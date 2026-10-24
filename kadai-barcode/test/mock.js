// Google スプレッドシート（Apps Script）の、テストに必要なところだけのまね
// Code.gs をそのまま読みこんで、Node で動かすために使う
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const TZ_OFFSET = 9 * 3600 * 1000; // Asia/Tokyo

function chain(target) {
  // 書式などの、テストでは気にしない呼び出しは何もせずに自分を返す
  return new Proxy(target, {
    get(t, k) {
      if (k in t) return t[k];
      if (typeof k === 'symbol') return undefined;
      return () => chain(t);
    },
  });
}

function a1(ref) {
  const m = ref.match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/);
  const col = (s) => s.split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
  const c0 = col(m[1]), r0 = m[2] ? Number(m[2]) : 1;
  const c1 = m[3] ? col(m[3]) : c0;
  const r1 = m[4] ? Number(m[4]) : (m[3] ? null : r0);
  return { r0, c0, r1, c1 };
}

class Sheet {
  constructor(ss, name) {
    this.ss = ss; this.name = name; this.id = ++ss.seq;
    this.data = []; this.fmt = {}; this.dv = {};
    this.maxRows = 1000; this.maxCols = 26;
  }
  getName() { return this.name; }
  getSheetId() { return this.id; }
  cell(r, c) { const row = this.data[r - 1]; return row && row[c - 1] !== undefined ? row[c - 1] : ''; }
  put(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    const row = this.data[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  getLastRow() {
    for (let r = this.data.length; r >= 1; r--) if (this.data[r - 1].some((v) => v !== '' && v !== null && v !== undefined)) return r;
    return 0;
  }
  getLastColumn() {
    let m = 0;
    this.data.forEach((row) => row.forEach((v, i) => { if (v !== '' && v !== null && v !== undefined) m = Math.max(m, i + 1); }));
    return m;
  }
  getMaxRows() { return Math.max(this.maxRows, this.data.length); }
  getMaxColumns() { return this.maxCols; }
  getRange(r, c, nr, nc) {
    if (typeof r === 'string') {
      const p = a1(r);
      return range(this, p.r0, p.c0, (p.r1 || this.getMaxRows()) - p.r0 + 1, p.c1 - p.c0 + 1);
    }
    return range(this, r, c, nr || 1, nc || 1);
  }
  appendRow(vals) { const r = this.getLastRow() + 1; vals.forEach((v, i) => this.put(r, i + 1, v)); this.ss.stats.writes++; return chain(this); }
  deleteRow(r) { this.data.splice(r - 1, 1); this.ss.stats.writes++; }
  insertRowsAfter(r, n) { this.maxRows += n; }
  insertColumnsAfter(c, n) { this.maxCols += n; }
  insertColumnAfter(c) { this.data.forEach((row) => { if (row.length > c) row.splice(c, 0, ''); }); }
  insertColumnBefore(c) { this.insertColumnAfter(c - 1); }
  clear() { this.data = []; this.fmt = {}; }
  getFilter() { return null; }
  createTextFinder(text) {
    const self = this;
    return { matchEntireCell() { return this; }, findNext() {
      for (let r = 1; r <= self.data.length; r++) for (let c = 1; c <= (self.data[r - 1] || []).length; c++) {
        const v = String(self.cell(r, c));
        if (v.indexOf(text) >= 0) return self.getRange(r, c);
      }
      return null;
    } };
  }
}

function range(sh, r, c, nr, nc) {
  const stats = sh.ss.stats;
  const t = {
    getSheet: () => sh, getRow: () => r, getColumn: () => c, getLastRow: () => r + nr - 1, getLastColumn: () => c + nc - 1,
    getValues() {
      stats.reads++; stats.cellsRead += nr * nc;
      const out = [];
      for (let i = 0; i < nr; i++) { const row = []; for (let j = 0; j < nc; j++) row.push(sh.cell(r + i, c + j)); out.push(row); }
      return out;
    },
    getValue() { stats.reads++; return sh.cell(r, c); },
    setValues(vals) { stats.writes++; vals.forEach((row, i) => row.forEach((v, j) => sh.put(r + i, c + j, v))); return p; },
    setValue(v) { stats.writes++; for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) sh.put(r + i, c + j, v); return p; },
    getNumberFormat() { return sh.fmt[r + ',' + c] || 'General'; },
    setNumberFormat(f) { for (let i = 0; i < Math.min(nr, 50); i++) for (let j = 0; j < nc; j++) sh.fmt[(r + i) + ',' + (c + j)] = f; return p; },
    getDataValidation() { return sh.dv[r + ',' + c] || null; },
    setDataValidation(v) { sh.dv[r + ',' + c] = v; return p; },
    insertCheckboxes() { for (let i = 0; i < nr; i++) if (sh.cell(r + i, c) === '') sh.put(r + i, c, false); return p; },
    uncheck() { for (let i = 0; i < nr; i++) sh.put(r + i, c, false); return p; },
    clearContent() { for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) if (sh.data[r + i - 1]) sh.put(r + i, c + j, ''); return p; },
  };
  const p = chain(t);
  return p;
}

class Spreadsheet {
  constructor() { this.sheets = []; this.seq = 0; this.active = null; this.stats = { reads: 0, cellsRead: 0, writes: 0 }; this.toasts = []; }
  getId() { return 'SSID'; }
  getSpreadsheetTimeZone() { return 'Asia/Tokyo'; }
  getSheetByName(n) { const s = this.sheets.find((x) => x.name === n); return s ? chain(s) : null; }
  getSheets() { return this.sheets.map(chain); }
  insertSheet(name, idx) { const s = new Sheet(this, name); this.sheets.splice(idx === undefined ? this.sheets.length : idx, 0, s); return chain(s); }
  deleteSheet(sh) { this.sheets = this.sheets.filter((x) => x.id !== sh.getSheetId()); }
  setActiveSheet(sh) { this.active = sh; return sh; }
  getActiveSheet() { return this.active || chain(this.sheets[0]); }
  moveActiveSheet() {}
  toast(msg) { this.toasts.push(msg); }
}

function props() {
  const store = {};
  return {
    store,
    getProperty: (k) => (k in store ? store[k] : null),
    setProperty(k, v) { store[k] = String(v); return this; },
    deleteProperty(k) { delete store[k]; },
    getProperties: () => Object.assign({}, store),
  };
}

function pad(n) { return ('0' + n).slice(-2); }

function formatDate(d, tz, f) {
  const j = new Date(d.getTime() + TZ_OFFSET);
  const y = j.getUTCFullYear(), M = j.getUTCMonth() + 1, D = j.getUTCDate(), H = j.getUTCHours(), m = j.getUTCMinutes();
  return f.replace('yyyy', y).replace('MM', pad(M)).replace('dd', pad(D)).replace('HH', pad(H)).replace('hh', pad(H)).replace('mm', pad(m))
    .replace(/\bM\b/, M).replace(/\bd\b/, D);
}

// Code.gs を新しい環境に読みこむ。users：先生ごとの UserProperties を切りかえるため
function load(opts) {
  opts = opts || {};
  const ss = new Spreadsheet();
  const users = {};
  let current = 'A';
  const docProps = props();
  const RealDate = Date;
  const now = opts.now ? new RealDate(opts.now) : new RealDate();
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length) super(...a); else super(now.getTime()); }
    static now() { return now.getTime(); }
    static [Symbol.hasInstance](x) { return x instanceof RealDate; }
  }
  const ctx = {
    console: { log() {}, warn: opts.quiet === false ? console.warn : () => {} },
    Date: FakeDate,
    SpreadsheetApp: {
      getActive: () => ss, flush() {},
      getUi: () => chain({ alert: () => 'OK', Button: { OK: 'OK' }, ButtonSet: {} }),
      BorderStyle: {}, DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST' },
      newDataValidation: () => chain({ build: () => ({ getCriteriaType: () => 'VALUE_IN_LIST' }) }),
    },
    Utilities: {
      formatDate,
      parseDate(key) { const [y, m, d] = key.split('-').map(Number); return new RealDate(RealDate.UTC(y, m - 1, d) - TZ_OFFSET); },
      getUuid: () => 'uuid',
    },
    PropertiesService: {
      getUserProperties: () => users[current] || (users[current] = props()),
      getDocumentProperties: () => docProps,
    },
    LockService: { getDocumentLock: () => ({ waitLock() {}, releaseLock() {} }) },
    HtmlService: { createHtmlOutputFromFile: () => chain({}), createTemplateFromFile: () => chain({}), createHtmlOutput: () => chain({}) },
  };
  vm.createContext(ctx);
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  // 1回の実行ごとに忘れる値（Apps Script では呼び出しのたびに新しく始まる）
  vm.runInContext(code + '\n;this.__reset = function () { SETTINGS_ = null; UPROPS_ = null; EXC_ = null; INC_COUNTS_ = null; HOLIDAYS_ = null; };', ctx);
  return {
    ctx, ss, docProps,
    as(user) { current = user; ctx.__reset(); },
    // Apps Script の1回の呼び出しのように、キャッシュを空にしてから呼ぶ
    call(name, ...args) { ctx.__reset(); return ctx[name](...args); },
    sheet: (n) => ss.sheets.find((x) => x.name === n),
  };
}

module.exports = { load, TZ_OFFSET };
