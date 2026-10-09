// 授業の出欠（教務手帳）：Google スプレッドシートに貼って使うプログラム
// 使い方は shukketsu/README.md を見てください。
//
// しくみ：講座ごとに「教務手帳」のシートを1枚作ります。
//  上の段に日付、欠席した生徒のマスに「欠」。名前の横の欠課時数・割合はシートの数式で数えるので、
//  シートを直接直しても（「欠」を消す・書く）すぐに反映されます。

const SHEET = {
  HOWTO: '使い方',
  STUDENTS: '名簿',
  COURSES: '講座',
  SETTINGS: '設定',
  PERIOD: '期間集計',
  SEATS: '座席表',
};
const RESERVED = [SHEET.HOWTO, SHEET.STUDENTS, SHEET.COURSES, SHEET.SETTINGS, SHEET.PERIOD, SHEET.SEATS];

// このプログラムの版（サイドバーのいちばん下に出ます。貼り直しが反映されたかの確認用）
const VERSION = '10/9-2';

const STUDENT_HEADERS = ['学籍番号', 'クラス', '番号', '氏名', 'ふりがな', '除外（転出などは ✓）'];
const COURSE_HEADERS = ['講座名（＝教務手帳のシート名）', '対象（クラス・学籍番号を「,」区切り）', '単位数', 'メモ'];
// 座席表：対象（講座の「対象」と同じ書き方）ごとに1行。同じクラスの講座は同じ座席表を使う
const SEAT_HEADERS = ['対象（クラス）', '写真（ドライブのファイルID）', '配置（さわらない）', '更新日時'];
const SEAT_FOLDER = '出欠_座席表の写真';

// 教務手帳のシートの形
//  1行目：講座名・単位数・年間時数　2行目：色の説明　3行目：見出し・日付　4行目：曜日　5行目から：生徒
//  A クラス／B 番号／C 学籍番号／D 氏名／E 欠課時数／F 割合／G それ以前の欠課（まとめて入力）／H すき間／I から右：日付
//  （H 列のすき間があるので、いちばん左に日付の列を入れても E 列の数式「H5:5」の範囲に入る）
const NB = { HEAD: 3, SUB: 4, TOP: 5, CLS: 1, NO: 2, KEY: 3, NAME: 4, TOTAL: 5, RATE: 6, BEFORE: 7, SEP: 8, FIRST: 9 };
const NB_HEADERS = ['クラス', '番号', '学籍番号', '氏名', '欠課時数', '割合', 'それ以前の欠課', ''];
const MARK = '欠';

const SETTING = {
  HOURS: '1単位あたりの年間授業時数',
  LINES: '色を変える欠課時数の割合（%・「,」区切り）',
  TERMS: '学期の区切り（期間集計のボタンになります）',
};
const SETTING_DEFAULTS = [
  [SETTING.HOURS, 35],
  [SETTING.LINES, '20, 25, 30, 50'],
  [SETTING.TERMS, '1学期 4/1-7/31, 2学期 8/1-12/31, 3学期 1/1-3/31'],
];

// 割合のラインの色（低い順）。ラインが5つ以上なら、いちばん濃い色をくり返す
const LINE_COLORS = [
  { bg: '#fff2cc', fg: '#7f6000' },
  { bg: '#fce5cd', fg: '#b45f06' },
  { bg: '#f4cccc', fg: '#990000' },
  { bg: '#cc0000', fg: '#ffffff' },
];
const COLOR = { header: '#cfe2f3', mark: '#f4cccc', markText: '#c5221f', sep: '#d9d9d9' };

// ───────── メニュー ─────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 出欠')
    .addItem('✋ 出欠をとる', 'openTake')
    .addItem('🔎 期間を指定して集計', 'openPeriod')
    .addSeparator()
    .addItem('⚙ 初期設定（名簿・講座・単位数）', 'openSetup')
    .addItem('📝 まとめて入力（使い始める前の欠課）', 'openBulk')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('その他')
      .addItem('教務手帳の色・数式を整える', 'refreshNotebooks')
      .addItem('お試しデータを入れる', 'insertSampleData')
      .addItem('シートを作り直す（消したシートを戻す）', 'setup'))
    .addToUi();
}

// 出欠をとる画面は、シートを見ながら使えるように「閉じなくてもシートをさわれる」ダイアログで開く
function openTake() {
  setup_();
  SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutputFromFile('Take').setWidth(1120).setHeight(640), '出欠をとる');
}

function openSetup(tab) {
  setup_();
  const t = HtmlService.createTemplateFromFile('Setup');
  t.tab = typeof tab === 'string' ? tab : '';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(960).setHeight(680), '初期設定');
}

function openBulk() {
  openSetup('bulk');
}

function openPeriod() {
  setup_();
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Period').setWidth(960).setHeight(680), '期間を指定して集計');
}

// ───────── 出欠をとる（Take.html） ─────────

function getTakeInit() {
  setup_();
  return {
    version: VERSION,
    today: todayKey_(),
    lines: lines_(),
    courses: getCourses_().map(function (c) { return { name: c.name, units: c.units, target: c.target }; }),
    students: getStudents_().filter(function (s) { return !s.excluded; }).length,
  };
}

// 講座を選んだとき：その講座の生徒と、その日の列（すでに記録した列）を返す。教務手帳のシートも開く
function getCourseDay(name, date) {
  const c = findCourse_(name);
  const sh = notebook_(c);
  syncRows_(sh, c);
  SpreadsheetApp.getActive().setActiveSheet(sh);
  return dayData_(sh, c, dateKey_(date) || todayKey_());
}

// o: { course, date, col（直すときの列。新しく記録するときは 0）, double（2時間続き）, absent: [学籍番号] }
function saveDay(o) {
  const c = findCourse_(o.course);
  const date = dateKey_(o.date);
  if (!date) throw new Error('日付がわかりません');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = notebook_(c);
    const rows = syncRows_(sh, c);
    let cols;
    if (o.col) {
      if (dateKey_(sh.getRange(NB.HEAD, o.col).getValue()) !== date) {
        throw new Error('教務手帳の列が動いたようです。もう一度、講座を選び直してください。');
      }
      cols = [Number(o.col)];
    } else {
      cols = insertDayColumns_(sh, date, o.double ? 2 : 1);
    }
    const absent = {};
    (o.absent || []).forEach(function (k) { absent[norm_(k)] = true; });
    const members = {};
    members_(c, getStudents_().filter(function (s) { return !s.excluded; })).forEach(function (s) { members[s.key] = true; });
    const last = sh.getLastRow();
    cols.forEach(function (col) {
      if (last < NB.TOP) return;
      const range = sh.getRange(NB.TOP, col, last - NB.TOP + 1, 1);
      const vals = range.getValues();
      Object.keys(rows).forEach(function (key) {
        const i = rows[key] - NB.TOP;
        if (!members[key]) return;
        const cur = String(vals[i][0]).trim();
        if (absent[key]) vals[i][0] = MARK;
        else if (cur === MARK) vals[i][0] = ''; // 手で書いたメモなど、「欠」以外はそのまま
      });
      range.setValues(vals);
    });
    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().setActiveSheet(sh);
    sh.getRange(NB.HEAD, cols[0], Math.max(1, last - NB.HEAD + 1), cols.length).activate();
    const d = dayData_(sh, c, date);
    d.saved = cols;
    return d;
  } finally {
    lock.releaseLock();
  }
}

function dayData_(sh, c, date) {
  const lastCol = lastDateCol_(sh);
  const last = sh.getLastRow();
  const width = Math.max(lastCol, NB.FIRST);
  const grid = last >= NB.HEAD ? sh.getRange(NB.HEAD, 1, last - NB.HEAD + 1, width).getValues() : [];
  const head = grid[0] || [];
  const dayCols = [];
  let lessons = 0;
  for (let col = NB.FIRST; col <= lastCol; col++) {
    const k = dateKey_(head[col - 1]);
    if (!k) continue;
    lessons++;
    if (k === date) dayCols.push(col);
  }
  const members = {};
  members_(c, getStudents_().filter(function (s) { return !s.excluded; })).forEach(function (s) { members[s.key] = s; });
  const students = [];
  for (let i = NB.TOP - NB.HEAD; i < grid.length; i++) {
    const r = grid[i];
    const key = rowKey_(r);
    if (!key || !members[key]) continue;
    let absentAll = 0;
    const day = {};
    for (let col = NB.FIRST; col <= lastCol; col++) {
      const v = String(r[col - 1]).trim();
      if (v === MARK && dateKey_(head[col - 1])) absentAll++;
      if (dayCols.indexOf(col) >= 0) day[col] = v;
    }
    students.push({ key: key, cls: String(r[NB.CLS - 1]), no: r[NB.NO - 1], name: String(r[NB.NAME - 1]), before: Number(r[NB.BEFORE - 1]) || 0, absentAll: absentAll, day: day });
  }
  return { course: c.name, target: c.target, units: c.units, hours: hoursOf_(sh), date: date, dayCols: dayCols, lessons: lessons, students: students, lines: lines_() };
}

// 日付の順になるように、新しい列を入れる（I列から右）。入れた列の番号を返す
function insertDayColumns_(sh, date, n) {
  const lastCol = lastDateCol_(sh);
  let pos = 0;
  if (lastCol >= NB.FIRST) {
    const head = sh.getRange(NB.HEAD, NB.FIRST, 1, lastCol - NB.FIRST + 1).getValues()[0];
    for (let i = 0; i < head.length; i++) {
      const k = dateKey_(head[i]);
      if (k && k > date) { pos = NB.FIRST + i; break; }
    }
  }
  if (pos) {
    sh.insertColumnsBefore(pos, n);
  } else {
    pos = Math.max(lastCol + 1, NB.FIRST);
    if (sh.getMaxColumns() < pos + n - 1) sh.insertColumnsAfter(sh.getMaxColumns(), pos + n - 1 - sh.getMaxColumns() + 50);
  }
  const d = toDate_(date);
  const wd = '日月火水木金土'.charAt(Number(Utilities.formatDate(d, tz_(), 'u')) % 7);
  const head = [], sub = [];
  for (let i = 0; i < n; i++) { head.push(d); sub.push(wd); }
  sh.getRange(NB.HEAD, pos, 1, n).setValues([head]).setNumberFormat('m/d').setFontWeight('bold').setBackground(COLOR.header);
  sh.getRange(NB.SUB, pos, 1, n).setValues([sub]).setBackground(COLOR.header);
  sh.getRange(NB.HEAD, pos, Math.max(2, sh.getMaxRows() - NB.HEAD + 1), n).setHorizontalAlignment('center');
  for (let i = 0; i < n; i++) sh.setColumnWidth(pos + i, 42);
  applyFormats_(sh);
  const out = [];
  for (let i = 0; i < n; i++) out.push(pos + i);
  return out;
}

// ───────── 座席表 ─────────

// target（講座の「対象」）の座席表：写真（data URL）と、生徒ごとのボタンの位置（写真の左上を 0、右下を 1）
function getSeat(target) {
  const row = seatRow_(target);
  const out = { target: target, image: '', spots: [], updated: '' };
  if (!row) return out;
  try { out.spots = JSON.parse(row.json || '[]'); } catch (e) { out.spots = []; }
  out.updated = row.updated instanceof Date ? Utilities.formatDate(row.updated, tz_(), 'yyyy/M/d') : String(row.updated || '');
  if (row.fileId) {
    try {
      const blob = DriveApp.getFileById(row.fileId).getBlob();
      out.image = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (e) {
      out.missing = true; // 写真がドライブから消された
    }
  }
  return out;
}

// o: { target, image（新しい写真の data URL。変えないときは空）, spots: [{ key, x, y }] }
function saveSeat(o) {
  const target = splitList_(o.target).join(', ');
  if (!target) throw new Error('どのクラスの座席表かわかりません');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const row = seatRow_(target);
    let fileId = row ? row.fileId : '';
    if (o.image) {
      const m = String(o.image).match(/^data:(image\/[\w+.-]+);base64,(.+)$/);
      if (!m) throw new Error('写真が読めませんでした');
      const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], '座席表_' + target.replace(/[,\s]+/g, '_') + '.jpg');
      const file = seatFolder_().createFile(blob);
      if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* 前の写真がもうない */ } }
      fileId = file.getId();
    }
    const spots = (o.spots || []).map(function (p) {
      const cl = function (v) { v = Number(v); return isNaN(v) ? 0.5 : Math.max(0, Math.min(1, Math.round(v * 10000) / 10000)); };
      return { key: norm_(p.key), x: cl(p.x), y: cl(p.y) };
    }).filter(function (p) { return p.key; });
    const sh = sheet_(SHEET.SEATS);
    const vals = [target, fileId, JSON.stringify(spots), new Date()];
    const r = row ? row.row : lastDataRow_(sh, 1) + 1;
    ensureSize_(sh, r, SEAT_HEADERS.length);
    sh.getRange(r, 1, 1, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, vals.length).setValues([vals]);
    SpreadsheetApp.flush();
    return getSeat(target);
  } finally {
    lock.releaseLock();
  }
}

function seatRow_(target) {
  const key = splitList_(target).join(', ');
  const sh = sheet_(SHEET.SEATS);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const vals = sh.getRange(2, 1, last - 1, SEAT_HEADERS.length).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (splitList_(vals[i][0]).join(', ') === key) return { row: i + 2, fileId: String(vals[i][1] || '').trim(), json: String(vals[i][2] || ''), updated: vals[i][3] };
  }
  return null;
}

// 写真は、このスプレッドシートと同じフォルダの「出欠_座席表の写真」に入れる
function seatFolder_() {
  const parents = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const found = parent.getFoldersByName(SEAT_FOLDER);
  return found.hasNext() ? found.next() : parent.createFolder(SEAT_FOLDER);
}

// ───────── 初期設定（Setup.html） ─────────

function getSetupData() {
  setup_();
  const st = getSettings_();
  const students = getStudents_().filter(function (s) { return !s.excluded; });
  return {
    version: VERSION,
    students: students.length,
    classes: classList_(students),
    courses: getCourses_().map(function (c) {
      return { name: c.name, target: c.target, units: c.units, memo: c.memo, members: members_(c, students).length };
    }),
    hoursPerUnit: hoursPerUnit_(st),
    lines: String(st[SETTING.LINES] || ''),
    terms: String(st[SETTING.TERMS] || ''),
  };
}

// o: { oldName（直すとき）, name, target, units, memo }
function saveCourse(o) {
  const name = String(o.name || '').trim();
  if (!name) throw new Error('講座名を入れてください');
  if (/[\[\]\*\?\/\\:]/.test(name)) throw new Error('講座名に [ ] * ? / \\ : は使えません（シートの名前になるため）');
  if (RESERVED.indexOf(name) >= 0) throw new Error('「' + name + '」は講座名に使えません');
  const target = splitList_(o.target).join(', ');
  if (!target) throw new Error('対象のクラス（または学籍番号）を選んでください');
  const units = Number(String(o.units || '').normalize('NFKC'));
  if (!(units > 0 && units <= 20)) throw new Error('単位数を入れてください（例：3）');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const courses = getCourses_();
    const old = o.oldName ? courses.filter(function (c) { return c.name === o.oldName; })[0] : null;
    if (courses.some(function (c) { return c.name === name && c !== old; })) throw new Error('「' + name + '」という講座はもうあります');
    if ((!old || old.name !== name) && ss.getSheetByName(name)) {
      throw new Error('「' + name + '」という名前のシートがもうあります。ちがう講座名にしてください');
    }
    const sh = sheet_(SHEET.COURSES);
    const row = [name, target, units, String(o.memo || '').trim()];
    let r;
    if (old) {
      r = old.row;
      const nb = ss.getSheetByName(old.name);
      if (nb && old.name !== name) nb.setName(name);
    } else {
      r = lastDataRow_(sh, 1) + 1;
      ensureSize_(sh, r, COURSE_HEADERS.length);
    }
    sh.getRange(r, 1, 1, 2).setNumberFormat('@');
    sh.getRange(r, 1, 1, row.length).setValues([row]);
    const c = { name: name, target: target, units: units, memo: row[3] };
    const nb = notebook_(c);
    nb.getRange(1, 1).setValue(name);
    nb.getRange(1, 5).setValue(units);
    syncRows_(nb, c);
    SpreadsheetApp.flush();
    return getSetupData();
  } finally {
    lock.releaseLock();
  }
}

// 講座の一覧から外す（教務手帳のシートは残す）
function deleteCourse(name) {
  const c = getCourses_().filter(function (x) { return x.name === name; })[0];
  if (c) sheet_(SHEET.COURSES).deleteRow(c.row);
  return getSetupData();
}

function saveSettings(o) {
  const hours = Number(String(o.hoursPerUnit || '').normalize('NFKC'));
  if (!(hours > 0 && hours <= 100)) throw new Error('1単位あたりの時数を入れてください（ふつうは 35）');
  const lines = parseLines_(o.lines);
  if (!lines.length) throw new Error('色を変える割合を入れてください（例：20, 25, 30, 50）');
  const terms = String(o.terms || '').trim();
  terms_(terms, fiscalYear_(todayKey_()), true); // 読めない書き方ならここでエラー
  putSetting_(SETTING.HOURS, hours);
  putSetting_(SETTING.LINES, lines.map(function (v) { return Math.round(v * 1000) / 10; }).join(', '));
  putSetting_(SETTING.TERMS, terms);
  refreshNotebooks_();
  return getSetupData();
}

// まとめて入力：その講座の生徒と、今の「それ以前の欠課」
function getBulk(name) {
  const c = findCourse_(name);
  const sh = notebook_(c);
  const rows = syncRows_(sh, c);
  const last = sh.getLastRow();
  const vals = last >= NB.TOP ? sh.getRange(NB.TOP, 1, last - NB.TOP + 1, NB.BEFORE).getValues() : [];
  const members = members_(c, getStudents_().filter(function (s) { return !s.excluded; }));
  return {
    course: c.name,
    until: dateKey_(sh.getRange(NB.SUB, NB.BEFORE).getValue()),
    students: members.map(function (s) {
      const r = vals[rows[s.key] - NB.TOP] || [];
      const v = r[NB.BEFORE - 1];
      return { key: s.key, cls: s.cls, no: s.no, name: s.name, before: v === '' || v === undefined ? '' : Number(v) || 0 };
    }),
  };
}

// o: { course, until: 'yyyy-MM-dd'（いつまでの分か・なくてもよい）, values: { 学籍番号: 数 } }
function saveBulk(o) {
  const c = findCourse_(o.course);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = notebook_(c);
    const rows = syncRows_(sh, c);
    const last = sh.getLastRow();
    if (last < NB.TOP) return getBulk(c.name);
    const range = sh.getRange(NB.TOP, NB.BEFORE, last - NB.TOP + 1, 1);
    const vals = range.getValues();
    Object.keys(o.values || {}).forEach(function (k) {
      const row = rows[norm_(k)];
      if (!row) return;
      const v = String(o.values[k] === null || o.values[k] === undefined ? '' : o.values[k]).normalize('NFKC').trim();
      vals[row - NB.TOP][0] = v === '' ? '' : Math.max(0, Math.round(Number(v) || 0));
    });
    range.setValues(vals);
    const until = dateKey_(o.until);
    sh.getRange(NB.SUB, NB.BEFORE).setValue(until ? toDate_(until) : '').setNumberFormat('m/d"まで"');
    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().setActiveSheet(sh);
    return getBulk(c.name);
  } finally {
    lock.releaseLock();
  }
}

// ───────── 名簿 ─────────

// rows: [{ gakuseki, name, kana }]（画面で貼り付けた表を読み取ったもの）
function previewRoster(rows) {
  return rosterPlan_(rows).map(function (p) { return { gakuseki: p.gakuseki, name: p.name, kana: p.kana, kind: p.kind, before: p.before }; });
}

function rosterPlan_(rows) {
  const existing = getStudents_();
  const byGk = {};
  existing.forEach(function (s) { if (s.gakuseki) byGk[s.gakuseki] = s; });
  return (rows || []).map(function (r) {
    const gk = norm_(r.gakuseki);
    const name = String(r.name || '').trim();
    const kana = String(r.kana || '').trim();
    const m = gk ? byGk[gk] : null;
    let kind = 'new';
    if (m) kind = (m.name === name && !m.excluded && (!kana || m.kana === kana)) ? 'same' : 'update';
    return { gakuseki: gk, name: name, kana: kana, kind: kind, before: m ? m.name : '', row: m ? m.row : 0 };
  }).filter(function (p) { return p.name && p.gakuseki; });
}

function importRoster(rows) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const plan = rosterPlan_(rows);
    const sh = sheet_(SHEET.STUDENTS);
    let updated = 0;
    plan.forEach(function (p) {
      if (p.kind !== 'update') return;
      sh.getRange(p.row, 4).setValue(p.name);
      if (p.kana) sh.getRange(p.row, 5).setValue(p.kana);
      sh.getRange(p.row, 6).insertCheckboxes().uncheck();
      updated++;
    });
    const fresh = plan.filter(function (p) { return p.kind === 'new'; }).map(function (p) { return [p.gakuseki, '', '', p.name, p.kana, false]; });
    if (fresh.length) {
      const start = lastDataRow_(sh, 5) + 1;
      ensureSize_(sh, start + fresh.length, STUDENT_HEADERS.length);
      sh.getRange(start, 1, fresh.length, 2).setNumberFormat('@');
      sh.getRange(start, 1, fresh.length, STUDENT_HEADERS.length).setValues(fresh);
      sh.getRange(start, 6, fresh.length, 1).insertCheckboxes();
    }
    SpreadsheetApp.flush();
    getStudents_(); // 学籍番号からクラス・番号をここで入れる
    // 講座の教務手帳に、新しい生徒の行を足す
    getCourses_().forEach(function (c) {
      const nb = SpreadsheetApp.getActive().getSheetByName(c.name);
      if (nb) syncRows_(nb, c);
    });
    const out = getSetupData();
    out.result = { added: fresh.length, updated: updated, same: plan.length - fresh.length - updated };
    return out;
  } finally {
    lock.releaseLock();
  }
}

// ───────── 期間を指定して集計（Period.html） ─────────

function getPeriodInit() {
  const today = todayKey_();
  const fy = fiscalYear_(today);
  return {
    today: today,
    fy: fy,
    terms: terms_(getSettings_()[SETTING.TERMS], fy),
    courses: getCourses_().map(function (c) { return c.name; }),
  };
}

// o: { course（空欄＝すべての講座）, from, to }
// 講座ごとに、期間の中の日付の列で「欠」を数える
function runPeriod(o) {
  const from = dateKey_(o.from) || '0000-00-00';
  const to = dateKey_(o.to) || '9999-12-31';
  const courses = getCourses_().filter(function (c) { return !o.course || c.name === o.course; });
  if (!courses.length) throw new Error('講座がありません');
  const students = getStudents_().filter(function (s) { return !s.excluded; });
  const byKey = {};
  const outCourses = [];
  courses.forEach(function (c) {
    const sh = SpreadsheetApp.getActive().getSheetByName(c.name);
    if (!sh) return;
    const lastCol = lastDateCol_(sh);
    const last = sh.getLastRow();
    const members = {};
    members_(c, students).forEach(function (s) { members[s.key] = s; });
    const grid = last >= NB.HEAD ? sh.getRange(NB.HEAD, 1, last - NB.HEAD + 1, Math.max(lastCol, NB.FIRST)).getValues() : [];
    const head = grid[0] || [];
    const cols = [];
    for (let col = NB.FIRST; col <= lastCol; col++) {
      const k = dateKey_(head[col - 1]);
      if (k && k >= from && k <= to) cols.push(col);
    }
    const until = dateKey_(sh.getRange(NB.SUB, NB.BEFORE).getValue());
    const counts = {}, before = {};
    for (let i = NB.TOP - NB.HEAD; i < grid.length; i++) {
      const r = grid[i];
      const key = rowKey_(r);
      if (!key || !members[key]) continue;
      let n = 0;
      cols.forEach(function (col) { if (String(r[col - 1]).trim() === MARK) n++; });
      counts[key] = n;
      if (Number(r[NB.BEFORE - 1])) before[key] = Number(r[NB.BEFORE - 1]);
      byKey[key] = members[key];
    }
    outCourses.push({ name: c.name, lessons: cols.length, until: until, counts: counts, before: before, hasBefore: Object.keys(before).length > 0 });
  });
  const rows = sortStudents_(Object.keys(byKey).map(function (k) { return byKey[k]; })).map(function (s) {
    return { key: s.key, cls: s.cls, no: s.no, name: s.name };
  });
  return { from: o.from || '', to: o.to || '', courses: outCourses, rows: rows };
}

// 期間集計をシートに書き出す（印刷用）
function exportPeriod(o) {
  const r = runPeriod(o);
  const ss = SpreadsheetApp.getActive();
  const sh = freshSheet_(ss, SHEET.PERIOD);
  const label = (r.from || r.to) ? shortDate_(dateKey_(r.from)) + '〜' + shortDate_(dateKey_(r.to)) : '全期間';
  const multi = r.courses.length > 1;
  const head = ['クラス', '番号', '学籍番号', '氏名'].concat(r.courses.map(function (c) { return c.name + '（' + c.lessons + 'コマ中）'; }));
  if (multi) head.push('合計');
  sh.getRange(1, 1, 2, 1).setValues([['期間集計：休んだコマ数'], [label + '　' + (o.course || 'すべての講座')]]);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(COLOR.header).setWrap(true);
  const vals = r.rows.map(function (s) {
    const row = [s.cls, s.no, s.key, s.name];
    let sum = 0;
    r.courses.forEach(function (c) {
      const n = c.counts[s.key];
      row.push(n === undefined ? '－' : n);
      sum += n || 0;
    });
    if (multi) row.push(sum);
    return row;
  });
  if (vals.length) {
    sh.getRange(5, 3, vals.length, 1).setNumberFormat('@');
    sh.getRange(5, 1, vals.length, head.length).setValues(vals);
  }
  sh.setFrozenRows(4);
  sh.setColumnWidth(4, 140);
  SpreadsheetApp.flush();
  ss.setActiveSheet(sh);
  return true;
}

// ───────── 教務手帳のシート ─────────

function notebook_(c) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(c.name);
  if (!sh) {
    sh = ss.insertSheet(c.name, ss.getSheets().length);
    ensureSize_(sh, 200, 200);
    sh.getRange(1, 1, 1, 6).setValues([[c.name, '', '', '単位数', c.units, '年間時数']]);
    sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
    sh.getRange(1, 4).setHorizontalAlignment('right');
    sh.getRange(1, 6).setHorizontalAlignment('right');
    sh.getRange(1, 5).setFontWeight('bold');
    sh.getRange(1, 7).setFormula('=E1*' + hoursPerUnit_(getSettings_())).setFontWeight('bold');
    sh.getRange(NB.HEAD, 1, 1, NB_HEADERS.length).setValues([NB_HEADERS]);
    sh.getRange(NB.HEAD, 1, 2, NB.SEP).setFontWeight('bold').setBackground(COLOR.header).setVerticalAlignment('middle').setWrap(true);
    sh.getRange(NB.SUB, NB.BEFORE).setNumberFormat('m/d"まで"');
    sh.getRange('C:C').setNumberFormat('@');
    sh.getRange(1, NB.SEP, sh.getMaxRows(), 1).setBackground(COLOR.sep);
    [50, 40, 70, 120, 64, 64, 70, 6].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setFrozenRows(NB.SUB);
    sh.setFrozenColumns(NB.SEP);
    writeLegend_(sh);
    applyFormats_(sh);
  }
  return sh;
}

function writeLegend_(sh) {
  const names = ['黄', '橙', '赤', '濃い赤'];
  const txt = '割合（欠課時数 ÷ 年間時数）の色：' + lines_().map(function (v, i) { return Math.round(v * 1000) / 10 + '%以上 ' + names[Math.min(i, names.length - 1)]; }).join('／') +
    '　　マスの「欠」を数えます。あとから欠席扱いでないとわかったら、そのマスの「欠」を消すだけで直ります';
  sh.getRange(2, 1).setValue(txt).setFontColor('#5f6368');
}

// 講座の生徒の行がなければ足す（名簿にあとから入った生徒など）。{ 学籍番号: 行 } を返す
function syncRows_(sh, c) {
  const students = getStudents_().filter(function (s) { return !s.excluded; });
  const members = members_(c, students);
  const last = sh.getLastRow();
  const rows = {};
  if (last >= NB.TOP) {
    sh.getRange(NB.TOP, 1, last - NB.TOP + 1, NB.NAME).getValues().forEach(function (r, i) {
      const key = rowKey_(r);
      if (key && !rows[key]) rows[key] = NB.TOP + i;
    });
  }
  const add = members.filter(function (s) { return !rows[s.key]; });
  if (add.length) {
    let start = NB.TOP;
    if (last >= NB.TOP) {
      const names = sh.getRange(NB.TOP, NB.NAME, last - NB.TOP + 1, 1).getValues();
      for (let i = names.length - 1; i >= 0; i--) if (String(names[i][0]).trim()) { start = NB.TOP + i + 1; break; }
    }
    const grow = sh.getMaxRows() < start + add.length;
    ensureSize_(sh, start + add.length, NB.FIRST);
    sh.getRange(start, NB.KEY, add.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, add.length, NB.NAME).setValues(add.map(function (s) { return [s.cls, s.no, s.key, s.name]; }));
    sh.getRange(start, NB.TOTAL, add.length, 2).setFormulas(add.map(function (s, i) {
      const r = start + i;
      return ['=IF(D' + r + '="","",N(G' + r + ')+COUNTIF(H' + r + ':' + r + ',"' + MARK + '"))', '=IF(OR(E' + r + '="",N($G$1)<=0),"",E' + r + '/$G$1)'];
    }));
    sh.getRange(start, NB.RATE, add.length, 1).setNumberFormat('0.0%');
    sh.getRange(start, NB.TOTAL, add.length, 3).setHorizontalAlignment('center');
    sh.getRange(start, NB.TOTAL, add.length, 1).setFontWeight('bold');
    add.forEach(function (s, i) { rows[s.key] = start + i; });
    if (grow) applyFormats_(sh);
  }
  return rows;
}

// 割合のマス（氏名〜割合）と「欠」のマスの色
function applyFormats_(sh) {
  const nRows = Math.max(1, sh.getMaxRows() - NB.TOP + 1);
  const nCols = Math.max(1, sh.getMaxColumns() - NB.FIRST + 1);
  const rateRange = sh.getRange(NB.TOP, NB.NAME, nRows, NB.RATE - NB.NAME + 1);
  const lines = lines_();
  const rules = [];
  for (let i = lines.length - 1; i >= 0; i--) {
    const c = LINE_COLORS[Math.min(i, LINE_COLORS.length - 1)];
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ISNUMBER($F' + NB.TOP + '),$F' + NB.TOP + '>=' + lines[i] + ')')
      .setBackground(c.bg).setFontColor(c.fg).setRanges([rateRange]).build());
  }
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(MARK).setBackground(COLOR.mark).setFontColor(COLOR.markText)
    .setRanges([sh.getRange(NB.TOP, NB.FIRST, nRows, nCols)]).build());
  sh.setConditionalFormatRules(rules);
}

function refreshNotebooks() {
  setup_();
  refreshNotebooks_();
  SpreadsheetApp.getUi().alert('教務手帳の色と数式を整えました。');
}

function refreshNotebooks_() {
  const hours = hoursPerUnit_(getSettings_());
  getCourses_().forEach(function (c) {
    const sh = SpreadsheetApp.getActive().getSheetByName(c.name);
    if (!sh) return;
    // 年間時数が「=E1*35」の形のときだけ、1単位あたりの時数を入れ直す（手で数を入れたときはそのまま）
    const f = sh.getRange(1, 7).getFormula();
    if (!f || /^=E1\*[\d.]+$/.test(f)) sh.getRange(1, 7).setFormula('=E1*' + hours);
    writeLegend_(sh);
    syncRows_(sh, c);
    applyFormats_(sh);
  });
}

// 3行目の、いちばん右の日付の列（なければ H 列）
function lastDateCol_(sh) {
  const maxCol = sh.getLastColumn();
  if (maxCol < NB.FIRST) return NB.SEP;
  const head = sh.getRange(NB.HEAD, NB.FIRST, 1, maxCol - NB.FIRST + 1).getValues()[0];
  for (let i = head.length - 1; i >= 0; i--) if (head[i] !== '' && head[i] !== null) return NB.FIRST + i;
  return NB.SEP;
}

function rowKey_(r) {
  const name = String(r[NB.NAME - 1] || '').trim();
  if (!name) return '';
  return norm_(r[NB.KEY - 1]) || norm_(normClass_(r[NB.CLS - 1]) + '-' + r[NB.NO - 1] + '-' + name);
}

// 年間時数（G1。手で数を入れてもよい）
function hoursOf_(sh) {
  const v = Number(sh.getRange(1, 7).getValue());
  if (v > 0) return v;
  const u = Number(sh.getRange(1, 5).getValue());
  return u > 0 ? u * hoursPerUnit_(getSettings_()) : 0;
}

// ───────── シートの準備 ─────────

function setup() {
  setup_();
  SpreadsheetApp.getUi().alert('シートを確認しました。足りないシートは作り直しました。');
}

function setup_() {
  const ss = SpreadsheetApp.getActive();
  function make(name, headers, widths) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, ss.getSheets().length);
    if (headers && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(COLOR.header).setWrap(true).setVerticalAlignment('middle');
      sh.setFrozenRows(1);
      (widths || []).forEach(function (w, i) { if (w) sh.setColumnWidth(i + 1, w); });
      return { sh: sh, fresh: true };
    }
    return { sh: sh, fresh: false };
  }
  const howto = make(SHEET.HOWTO);
  if (howto.sh.getLastRow() === 0) {
    writeHowTo_(howto.sh);
    ss.setActiveSheet(howto.sh);
    ss.moveActiveSheet(1);
  }
  const students = make(SHEET.STUDENTS, STUDENT_HEADERS, [90, 60, 50, 140, 160, 120]);
  if (students.fresh) students.sh.getRange('A:B').setNumberFormat('@');
  const courses = make(SHEET.COURSES, COURSE_HEADERS, [220, 280, 70, 200]);
  if (courses.fresh) courses.sh.getRange('A:B').setNumberFormat('@');
  const settings = make(SHEET.SETTINGS, ['項目', '値'], [380, 420]);
  if (settings.fresh) settings.sh.getRange(2, 1, SETTING_DEFAULTS.length, 2).setValues(SETTING_DEFAULTS);
  const seats = make(SHEET.SEATS, SEAT_HEADERS, [160, 280, 300, 140]);
  if (seats.fresh) seats.sh.getRange('A:B').setNumberFormat('@');
}

function writeHowTo_(sh) {
  const lines = [
    ['授業の出欠（教務手帳）'],
    [''],
    ['はじめに（1回だけ）：メニュー「📋 出欠」→「⚙ 初期設定」'],
    ['1. 名簿：Excel や校務システムの名簿をコピーして貼る（学籍番号と氏名）'],
    ['2. 講座：持っている授業の名前・対象のクラス・単位数を登録する → 講座ごとに教務手帳のシートができます'],
    ['3. まとめて入力：学期の途中から使うときは、それまでの欠課時数を生徒ごとに入れる'],
    [''],
    ['授業のたびに：メニュー「📋 出欠」→「✋ 出欠をとる」'],
    ['1. 講座を押す（日付は今日が入っています）'],
    ['2. 欠席した生徒を、座席表の写真の上か、右の名列で押す（赤くなります。もう一度押すと元に戻ります）'],
    ['3. 「教務手帳に書きこむ」を押す → 教務手帳のシートに日付の列ができ、欠席した生徒のマスに「欠」が入ります'],
    [''],
    ['教務手帳のシート'],
    ['・名前の横の「欠課時数」＝「それ以前の欠課」＋「欠」の数。「割合」＝欠課時数 ÷ 年間時数（単位数 × 35）'],
    ['・割合が 20%・25%・30%・50% 以上になると、名前と割合のマスの色が変わります（「設定」シートで変えられます）'],
    ['・マスを直接直してかまいません。あとから欠席扱いでなかったとわかったら、そのマスの「欠」を消すだけで合計と割合が直ります'],
    [''],
    ['期間を指定して集計：メニュー「📋 出欠」→「🔎 期間を指定して集計」で、その期間に何コマ休んだかを出せます'],
    [''],
    ['座席表：出欠をとる画面の「📷 座席表を登録」で、座席表を撮る（または写真を選ぶ）と、写真の上の名前がそのまま欠席を押すボタンになります'],
    ['・写真の名前は自動で読み取ります。読み取れなかった生徒は、右の名列で選んでから写真の上の席を押して置きます（ボタンはドラッグで動かせます）'],
    ['・席がえをしたら「座席表を直す」から撮り直してください。同じクラスの講座は同じ座席表を使います'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  [3, 8, 13, 18, 20].forEach(function (r) { sh.getRange(r, 1).setFontWeight('bold'); });
  sh.setColumnWidth(1, 900);
}

// ───────── お試しデータ ─────────

function insertSampleData() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('お試しの名簿（2クラス・20人）と講座・出欠を入れます。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  setup_();
  const family = ['青木', '石田', '上野', '江口', '大川', '加藤', '木村', '工藤', '小林', '佐藤'];
  const given = ['さくら', 'はると', 'ゆい', 'そうた', 'ひなた', 'りく', 'めい', 'ゆうと', 'あおい', 'こはる'];
  const rows = [];
  [1, 2].forEach(function (c) {
    family.forEach(function (f, i) { rows.push({ gakuseki: '1' + c + ('0' + (i + 1)).slice(-2), name: f + ' ' + given[(i + c * 3) % 10], kana: '' }); });
  });
  importRoster(rows);
  const names = getCourses_().map(function (c) { return c.name; });
  if (names.indexOf('1-1 英語コミュⅠ') < 0) saveCourse({ name: '1-1 英語コミュⅠ', target: '1-1', units: 3 });
  if (names.indexOf('1-2 論理表現Ⅰ') < 0) saveCourse({ name: '1-2 論理表現Ⅰ', target: '1-2', units: 2 });
  const today = todayKey_();
  saveBulk({ course: '1-1 英語コミュⅠ', until: addDaysKey_(today, -22), values: { '1103': 18, '1107': 22, '1104': 1 } });
  saveBulk({ course: '1-2 論理表現Ⅰ', until: addDaysKey_(today, -22), values: { '1205': 34 } });
  [21, 14, 7].forEach(function (back) {
    const d = addDaysKey_(today, -back);
    saveDay({ course: '1-1 英語コミュⅠ', date: d, absent: ['1103'] });
    saveDay({ course: '1-2 論理表現Ⅰ', date: d, absent: back === 7 ? ['1205', '1201'] : ['1205'] });
  });
  ui.alert('お試しデータを入れました。メニュー「📋 出欠」→「✋ 出欠をとる」でためしてください。');
}

// ───────── 読み書きの部品 ─────────

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」シートがありません。メニュー「📋 出欠」→「その他」→「シートを作り直す」を押してください。');
  return sh;
}

function getStudents_() {
  const sh = sheet_(SHEET.STUDENTS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, STUDENT_HEADERS.length).getValues();
  const list = [];
  vals.forEach(function (r, i) {
    const name = String(r[3]).trim();
    if (!name) return;
    const gakuseki = norm_(r[0]);
    let cls = normClass_(r[1]);
    let no = r[2];
    if (r[1] instanceof Date) sh.getRange(i + 2, 2).setNumberFormat('@').setValue(cls); // 日付になってしまったクラスを文字に戻す
    // 4けたの学籍番号（1101＝1年1組1番）から、空欄のクラス・番号を入れる
    const m = gakuseki.match(/^(\d)(\d)(\d\d)$/);
    if (m && (cls === '' || no === '')) {
      if (cls === '') cls = m[1] + '-' + m[2];
      if (no === '') no = Number(m[3]);
      sh.getRange(i + 2, 2).setNumberFormat('@').setValue(cls);
      sh.getRange(i + 2, 3).setValue(no);
    }
    const ex = r[5];
    list.push({
      key: gakuseki || norm_(cls + '-' + no + '-' + name),
      gakuseki: gakuseki, cls: cls, no: no === '' ? '' : Number(no) || no, name: name, kana: String(r[4]).trim(),
      excluded: ex === true || /^(✓|✔|○|〇|1|TRUE|除外|転出|×)$/i.test(String(ex).trim()),
      row: i + 2,
    });
  });
  return sortStudents_(list);
}

function getCourses_() {
  const sh = sheet_(SHEET.COURSES);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const out = [];
  sh.getRange(2, 1, last - 1, COURSE_HEADERS.length).getValues().forEach(function (r, i) {
    const name = String(r[0]).trim();
    if (!name) return;
    out.push({ name: name, target: splitList_(r[1]).join(', '), units: Number(r[2]) || 0, memo: String(r[3] || ''), row: i + 2 });
  });
  return out;
}

function findCourse_(name) {
  const c = getCourses_().filter(function (x) { return x.name === name; })[0];
  if (!c) throw new Error('講座「' + name + '」がありません（「講座」シートを確かめてください）');
  return c;
}

// 講座の対象の生徒：「1-1, 1-2」のようなクラスと、「1105」のような学籍番号（選択授業など）
function members_(course, students) {
  const classes = {}, ids = {};
  splitList_(course.target).forEach(function (t) {
    if (/^\d{4,}$/.test(t)) ids[norm_(t)] = true; else classes[normClass_(t)] = true;
  });
  return students.filter(function (s) { return classes[s.cls] || (s.gakuseki && ids[s.gakuseki]); });
}

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.SETTINGS);
  const out = {};
  SETTING_DEFAULTS.forEach(function (r) { out[r[0]] = r[1]; });
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0] !== '') out[String(r[0]).trim()] = r[1];
  });
  return out;
}

function putSetting_(key, value) {
  const sh = sheet_(SHEET.SETTINGS);
  const vals = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === key) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([key, value]);
}

function hoursPerUnit_(st) {
  const n = Number(String(st[SETTING.HOURS]).normalize('NFKC'));
  return n > 0 ? n : 35;
}

// 「20, 25, 30, 50」→ [0.2, 0.25, 0.3, 0.5]
function parseLines_(v) {
  return String(v || '').normalize('NFKC').split(/[,、\s]+/).map(function (s) { return Number(s.replace('%', '')); })
    .filter(function (n) { return n > 0 && n <= 100; }).map(function (n) { return n / 100; })
    .sort(function (a, b) { return a - b; });
}

function lines_() {
  const l = parseLines_(getSettings_()[SETTING.LINES]);
  return l.length ? l : [0.2, 0.25, 0.3, 0.5];
}

// 「1学期 4/1-7/31, 2学期 8/1-12/31, 3学期 1/1-3/31」→ その年度の日付の範囲（1〜3月は次の年）
function terms_(text, fy, strict) {
  const t = String(text || '').normalize('NFKC').replace(/(\d{1,2})月(\d{1,2})日/g, '$1/$2');
  const out = [];
  t.split(/[,、;\n]+/).forEach(function (part) {
    const p = part.trim();
    if (!p) return;
    const m = p.match(/^(.*?)\s*(?:(\d{4})\/)?(\d{1,2})\/(\d{1,2})\s*[-〜~ー]\s*(?:(\d{4})\/)?(\d{1,2})\/(\d{1,2})$/);
    if (!m) {
      if (strict) throw new Error('学期の区切りが読めません：「' + p + '」（例：1学期 4/1-7/31）');
      return;
    }
    const day = function (y, mo, d) {
      const year = y ? Number(y) : (Number(mo) >= 4 ? fy : fy + 1);
      return year + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
    };
    out.push({ name: m[1].trim() || ('期間' + (out.length + 1)), from: day(m[2], m[3], m[4]), to: day(m[5], m[6], m[7]) });
  });
  return out;
}

function fiscalYear_(key) {
  const p = key.split('-');
  return Number(p[1]) >= 4 ? Number(p[0]) : Number(p[0]) - 1;
}

function norm_(v) {
  return String(v === null || v === undefined ? '' : v).normalize('NFKC').replace(/\s+/g, '').toUpperCase();
}

function normClass_(v) {
  // 「1-1」はスプレッドシートが勝手に日付（1月1日）にしてしまうことがあるので、日付なら「月-日」に戻す
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'M-d');
  return String(v === null || v === undefined ? '' : v).normalize('NFKC').replace(/\s+/g, '');
}

function splitList_(v) {
  if (v instanceof Date && !isNaN(v)) return [normClass_(v)];
  return String(v || '').normalize('NFKC').split(/[,、，;；\/／\n]+/).map(function (s) { return s.replace(/\s+/g, ''); }).filter(String);
}

function sortStudents_(list) {
  return list.slice().sort(function (a, b) {
    const c = String(a.cls).localeCompare(String(b.cls), 'ja', { numeric: true });
    if (c) return c;
    return (Number(a.no) || 0) - (Number(b.no) || 0) || a.name.localeCompare(b.name, 'ja');
  });
}

function classList_(students) {
  const seen = {};
  students.forEach(function (s) { if (s.cls) seen[s.cls] = true; });
  return Object.keys(seen).sort(function (a, b) { return a.localeCompare(b, 'ja', { numeric: true }); });
}

function tz_() {
  return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Tokyo';
}

function todayKey_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

// 日付（セルの Date でも「2026/10/5」のような文字でも）を 'yyyy-MM-dd' にそろえる
function dateKey_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const m = String(v || '').normalize('NFKC').match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

// 'yyyy-MM-dd' を、スプレッドシートのタイムゾーンでその日の 0 時にする（dateKey_ と同じタイムゾーン）
function toDate_(key) {
  return Utilities.parseDate(key, tz_(), 'yyyy-MM-dd');
}

function addDaysKey_(key, n) {
  const p = key.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

function shortDate_(key) {
  if (!key) return '';
  const p = key.split('-');
  return Number(p[1]) + '/' + Number(p[2]);
}

function freshSheet_(ss, name) {
  const old = ss.getSheetByName(name);
  if (old) {
    if (ss.getActiveSheet().getSheetId() === old.getSheetId()) ss.setActiveSheet(ss.getSheets()[0]);
    ss.deleteSheet(old);
  }
  return ss.insertSheet(name, ss.getSheets().length);
}

// 実際に文字が入っている最後の行（チェックボックスの FALSE は数えない）
function lastDataRow_(sh, ncols) {
  const last = sh.getLastRow();
  if (last < 2) return Math.max(1, last);
  const vals = sh.getRange(2, 1, last - 1, ncols).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i].some(function (v) { return v !== '' && v !== false && v !== null; })) return i + 2;
  }
  return 1;
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}
