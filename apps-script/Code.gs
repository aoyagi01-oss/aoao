// AI英会話れんしゅう：教材スプレッドシートと先生用・生徒用ページをつなぐプログラム
// 使い方は README.md の「スプレッドシートとつなぐ」を見てください。

// ▼ 先生用ページで「公開」するときの合言葉。必ず自分だけの言葉に変えてください。
const PASSWORD = 'ここを合言葉に変える';

const HEADERS = ['公開日', 'タイトル', '場面', 'AIの役', '生徒の役', 'レベル', 'やりとりの回数', '目標表現', 'AIへの追加の指示', 'プロンプト（上級・任意）', '登録日時'];
const KEYS = ['date', 'title', 'scene', 'aiRole', 'studentRole', 'level', 'turns', 'phrases', 'extra', 'prompt', 'savedAt'];

function doGet(e) {
  const p = e.parameter || {};
  let out;
  try {
    if (p.action === 'save') {
      checkPassword_(p.key);
      out = { ok: true, lesson: save_(JSON.parse(decode_(p.data))) };
    } else if (p.action === 'check') {
      checkPassword_(p.key);
      out = { ok: true };
    } else if (p.action === 'list') {
      out = { ok: true, lessons: readAll_().reverse(), today: today_() };
    } else {
      out = { ok: true, lesson: todaysLesson_() };
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  const json = JSON.stringify(out);
  // ページからは <script> タグ（JSONP）で呼び出す
  if (p.callback && /^[\w$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function checkPassword_(key) {
  if (PASSWORD === 'ここを合言葉に変える') throw new Error('Apps Script の PASSWORD（合言葉）がまだ設定されていません');
  if (key !== PASSWORD) throw new Error('合言葉がちがいます');
}

function decode_(data) {
  let s = String(data || '');
  while (s.length % 4) s += '=';
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString('UTF-8');
}

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function tz_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'Asia/Tokyo';
}

function today_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v);
}

// "2026/10/1" や日付セルを "2026-10-01" にそろえる
function normDate_(v) {
  if (isDate_(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const m = String(v || '').match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function readAll_() {
  const values = sheet_().getDataRange().getValues();
  const head = values.shift().map(String);
  return values.map(function (r) {
    const o = {};
    KEYS.forEach(function (k, i) {
      const col = head.indexOf(HEADERS[i]);
      o[k] = col < 0 ? '' : r[col];
    });
    o.date = normDate_(o.date);
    o.savedAt = isDate_(o.savedAt) ? Utilities.formatDate(o.savedAt, tz_(), 'yyyy-MM-dd HH:mm') : String(o.savedAt || '');
    ['title', 'scene', 'aiRole', 'studentRole', 'level', 'phrases', 'extra', 'prompt'].forEach(function (k) { o[k] = String(o[k] || '').trim(); });
    o.turns = Number(o.turns) || 8;
    return o;
  }).filter(function (o) { return o.title; });
}

// 公開日が今日以前（または空欄）の行のうち、いちばん下の行が「今日の教材」
function todaysLesson_() {
  const t = today_();
  const rows = readAll_().filter(function (o) { return !o.date || o.date <= t; });
  return rows.length ? rows[rows.length - 1] : null;
}

function save_(f) {
  if (!f || !String(f.title || '').trim()) throw new Error('タイトルが空です');
  const date = normDate_(f.date);
  const row = [
    date ? Utilities.parseDate(date, tz_(), 'yyyy-MM-dd') : '',
    f.title, f.scene, f.aiRole, f.studentRole, f.level, Number(f.turns) || 8,
    f.phrases, f.extra, f.prompt || '', new Date()
  ].map(function (v) { return typeof v === 'string' ? v.replace(/^[=+\-@]/, "'$&") : v; });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_().appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return { title: f.title, date: date };
}
