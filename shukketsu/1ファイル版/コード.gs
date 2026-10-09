// ─────────────────────────────────────────────
// 授業の出欠（教務手帳）1ファイル版
// このファイルだけを Apps Script の「コード.gs」に貼れば動きます（HTML ファイルを作る必要はありません）。
// shukketsu/build_onefile.py で自動で作ったファイルです。直すときは Code.gs・Take.html などを直して作り直してください。
// ─────────────────────────────────────────────

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
const VERSION = '10/9-3';

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
// 2時間続きで書きこんだ列は、4行目（曜日）に「続き」と入れて、まとめて直せるようにする
const DOUBLE = '続き';

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

// 画面の HTML：「Take」などの HTML ファイルがあればそれを、なければ 1ファイル版に入っている HTML（HTML_FILES）を使う
function html_(name) {
  return typeof HTML_FILES !== 'undefined' && HTML_FILES[name] ? HtmlService.createHtmlOutput(HTML_FILES[name]) : HtmlService.createHtmlOutputFromFile(name);
}

function template_(name) {
  return typeof HTML_FILES !== 'undefined' && HTML_FILES[name] ? HtmlService.createTemplate(HTML_FILES[name]) : HtmlService.createTemplateFromFile(name);
}

// 出欠をとる画面は、シートを見ながら使えるように「閉じなくてもシートをさわれる」ダイアログで開く
function openTake() {
  setup_();
  SpreadsheetApp.getUi().showModelessDialog(html_('Take').setWidth(1120).setHeight(640), '出欠をとる');
}

function openSetup(tab) {
  setup_();
  const t = template_('Setup');
  t.tab = typeof tab === 'string' ? tab : '';
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(960).setHeight(680), '初期設定');
}

function openBulk() {
  openSetup('bulk');
}

function openPeriod() {
  setup_();
  SpreadsheetApp.getUi().showModalDialog(html_('Period').setWidth(960).setHeight(680), '期間を指定して集計');
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
  // 生徒の行を足すことがあるので、書きこみと重ならないようにする
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = notebook_(c);
    syncRows_(sh, c);
    SpreadsheetApp.getActive().setActiveSheet(sh);
    return dayData_(sh, c, dateKey_(date) || todayKey_());
  } finally {
    lock.releaseLock();
  }
}

// o: { course, date, cols（直すときの列。2時間続きなら2つ。新しく記録するときは空）, double（2時間続き）, absent: [学籍番号] }
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
    const fix = (o.cols || (o.col ? [o.col] : [])).map(Number).filter(function (n) { return n >= NB.FIRST; });
    if (fix.length) {
      fix.forEach(function (col) {
        if (dateKey_(sh.getRange(NB.HEAD, col).getValue()) !== date) {
          throw new Error('教務手帳の列が動いたようです。もう一度、講座を選び直してください。');
        }
      });
      cols = fix;
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
  const sub = grid[1] || [];
  const dayCols = [];
  let lessons = 0;
  for (let col = NB.FIRST; col <= lastCol; col++) {
    const k = dateKey_(head[col - 1]);
    if (!k) continue;
    lessons++;
    if (k === date) dayCols.push(col);
  }
  // その日の授業：となり合った「続き」の列は1つの授業（2時間続き）にまとめる
  const isDouble = function (col) { return String(sub[col - 1] || '').indexOf(DOUBLE) >= 0; };
  const groups = [];
  dayCols.forEach(function (col) {
    const g = groups[groups.length - 1];
    if (g && isDouble(col) && isDouble(g[g.length - 1]) && g[g.length - 1] === col - 1) g.push(col);
    else groups.push([col]);
  });
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
      if (v === MARK) absentAll++; // シートの数式（E列）と同じく、日付のない列の「欠」も数える
      if (dayCols.indexOf(col) >= 0) day[col] = v;
    }
    students.push({ key: key, cls: normClass_(r[NB.CLS - 1]), no: r[NB.NO - 1], name: String(r[NB.NAME - 1]), before: Number(r[NB.BEFORE - 1]) || 0, absentAll: absentAll, day: day });
  }
  return { course: c.name, target: c.target, units: c.units, hours: hoursOf_(sh), date: date, dayCols: dayCols, groups: groups, lessons: lessons, students: students, lines: lines_() };
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
    // 入れた列が左どなりの H 列（すき間）の灰色を引きつがないように、生徒の行の色を消す
    sh.getRange(NB.TOP, pos, Math.max(1, sh.getMaxRows() - NB.TOP + 1), n).setBackground(null);
  } else {
    pos = Math.max(lastCol + 1, NB.FIRST);
    if (sh.getMaxColumns() < pos + n - 1) sh.insertColumnsAfter(sh.getMaxColumns(), pos + n - 1 - sh.getMaxColumns() + 50);
  }
  const d = toDate_(date);
  const wd = '日月火水木金土'.charAt(Number(Utilities.formatDate(d, tz_(), 'u')) % 7);
  const head = [], sub = [];
  for (let i = 0; i < n; i++) { head.push(d); sub.push(n > 1 ? wd + '\n' + DOUBLE : wd); }
  sh.getRange(NB.HEAD, pos, 1, n).setValues([head]).setNumberFormat('m/d').setFontWeight('bold').setBackground(COLOR.header);
  sh.getRange(NB.SUB, pos, 1, n).setValues([sub]).setBackground(COLOR.header).setWrap(true);
  sh.getRange(NB.HEAD, pos, Math.max(2, sh.getMaxRows() - NB.HEAD + 1), n).setHorizontalAlignment('center');
  for (let i = 0; i < n; i++) sh.setColumnWidth(pos + i, 42);
  applyFormats_(sh);
  const out = [];
  for (let i = 0; i < n; i++) out.push(pos + i);
  return out;
}

// ───────── 座席表 ─────────

// target（講座の「対象」）の座席表：写真（data URL）と、生徒ごとのボタンの位置（写真の左上を 0、右下を 1）
//  desk：写真の中の教卓の位置（'top'・'bottom'・''）、flip：表を上下左右さかさまに出すか、view：'grid'（見やすい表）か 'photo'（写真）
function getSeat(target) {
  const row = seatRow_(target);
  const out = { target: target, image: '', spots: [], desk: '', flip: false, view: 'grid', updated: '' };
  if (!row) return out;
  Object.assign(out, seatLayout_(row.json));
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

// 「配置」の列の中身。はじめのころの形（ボタンの位置の配列だけ）も読む
function seatLayout_(json) {
  let v;
  try { v = JSON.parse(json || '[]'); } catch (e) { v = []; }
  if (Array.isArray(v)) v = { spots: v };
  return { spots: v.spots || [], desk: v.desk === 'top' || v.desk === 'bottom' ? v.desk : '', flip: !!v.flip, view: v.view === 'photo' ? 'photo' : 'grid' };
}

// 見やすい表・写真の切りかえ、向き（教卓から見た向き）だけを保存する
function saveSeatView(target, o) {
  const row = seatRow_(target);
  if (!row) return false;
  const lay = seatLayout_(row.json);
  if (o.flip !== undefined) lay.flip = !!o.flip;
  if (o.view !== undefined) lay.view = o.view === 'photo' ? 'photo' : 'grid';
  sheet_(SHEET.SEATS).getRange(row.row, 3).setValue(JSON.stringify(lay));
  return true;
}

// o: { target, image（新しい写真の data URL。変えないときは空）, spots: [{ key, x, y }], desk }
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
    const lay = row ? seatLayout_(row.json) : seatLayout_('');
    lay.spots = spots;
    if (o.desk !== undefined) lay.desk = o.desk === 'top' || o.desk === 'bottom' ? o.desk : '';
    const vals = [target, fileId, JSON.stringify(lay), new Date()];
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
  if (name.length > 100) throw new Error('講座名が長すぎます（100文字まで）');
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
    // シートの名前は大文字・小文字を区別しないので、名前の大文字・小文字だけを直すときは自分のシートを除く
    const same = ss.getSheetByName(name);
    const mine = old ? ss.getSheetByName(old.name) : null;
    // 一覧から外した講座を同じ名前で登録し直すときは、残っていた教務手帳のシートをそのまま使う
    const adopt = !old && same && String(same.getRange(NB.HEAD, 1).getValue()).trim() === NB_HEADERS[0];
    if (adopt && same.getName() !== name) same.setName(name);
    if (!adopt && (!old || old.name !== name) && same && !(mine && same.getSheetId() === mine.getSheetId())) {
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
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const c = getCourses_().filter(function (x) { return x.name === name; })[0];
    if (c) {
      const sh = sheet_(SHEET.COURSES);
      // 固定した行のほかに1行しかないと deleteRow はエラーになるので、そのときは中身を消す
      if (sh.getMaxRows() - sh.getFrozenRows() <= 1) sh.getRange(c.row, 1, 1, COURSE_HEADERS.length).clearContent();
      else sh.deleteRow(c.row);
    }
  } finally {
    lock.releaseLock();
  }
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
  const seen = {};
  return (rows || []).map(function (r) {
    const gk = norm_(r.gakuseki);
    const name = String(r.name || '').trim();
    const kana = String(r.kana || '').trim();
    const m = gk ? byGk[gk] : null;
    let kind = 'new';
    if (m) kind = (m.name === name && !m.excluded && (!kana || m.kana === kana)) ? 'same' : 'update';
    return { gakuseki: gk, name: name, kana: kana, kind: kind, before: m ? m.name : '', row: m ? m.row : 0 };
  }).filter(function (p) {
    // 同じ学籍番号が2回貼られていたら、はじめの行だけ使う（2人分登録されないように）
    if (!p.name || !p.gakuseki || seen[p.gakuseki]) return false;
    seen[p.gakuseki] = true;
    return true;
  });
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
    studentsCache_ = null; // 名簿を書きかえたので読み直す
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
    sh.getRange(5, 1, vals.length, 1).setNumberFormat('@'); // クラス「1-1」が日付にならないように
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
    sh.getRange('A:A').setNumberFormat('@');
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
    // クラス（1-1）は文字にしておかないと日付（1月1日）になってしまう
    sh.getRange(start, NB.CLS, add.length, 1).setNumberFormat('@');
    sh.getRange(start, NB.KEY, add.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, add.length, NB.NAME).setValues(add.map(function (s) { return [s.cls, s.no, s.key, s.name]; }));
    writeRowFormulas_(sh, start, add.length);
    sh.getRange(start, NB.TOTAL, add.length, 3).setHorizontalAlignment('center');
    sh.getRange(start, NB.TOTAL, add.length, 1).setFontWeight('bold');
    add.forEach(function (s, i) { rows[s.key] = start + i; });
    if (grow) applyFormats_(sh);
  }
  return rows;
}

// E 列（欠課時数）・F 列（割合）の数式を、start 行から n 行入れる
function writeRowFormulas_(sh, start, n) {
  if (n < 1) return;
  const f = [];
  for (let r = start; r < start + n; r++) {
    f.push(['=IF(D' + r + '="","",N(G' + r + ')+COUNTIF(H' + r + ':' + r + ',"' + MARK + '"))', '=IF(OR(E' + r + '="",N($G$1)<=0),"",E' + r + '/$G$1)']);
  }
  sh.getRange(start, NB.TOTAL, n, 2).setFormulas(f);
  sh.getRange(start, NB.RATE, n, 1).setNumberFormat('0.0%');
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
    // 欠課時数・割合の数式を入れ直す（手で消したり、行を動かしたりしたとき用）
    const last = sh.getLastRow();
    if (last >= NB.TOP) writeRowFormulas_(sh, NB.TOP, last - NB.TOP + 1);
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
    ['・名前の横の「欠課時数」＝「それ以前の欠課」＋「欠」の数。「割合」＝欠課時数 ÷ 年間時数（単位数 × 1単位あたりの時数。最初は35）'],
    ['・割合が 20%・25%・30%・50% 以上になると、名前と割合のマスの色が変わります（「設定」シートで変えられます）'],
    ['・マスを直接直してかまいません。あとから欠席扱いでなかったとわかったら、そのマスの「欠」を消すだけで合計と割合が直ります'],
    [''],
    ['期間を指定して集計：メニュー「📋 出欠」→「🔎 期間を指定して集計」で、その期間に何コマ休んだかを出せます'],
    [''],
    ['座席表：出欠をとる画面の「📷 座席表を登録する」で座席表を撮る（または写真を選ぶ）と、出席番号と名前を読み取って、席ごとのボタンを並べた見やすい座席表ができます'],
    ['・押すと欠席になります。「写真」に切りかえると写真の上でも押せます。「↕ 向きを変える」で教卓から見た向きにできます'],
    ['・読み取れなかった生徒は、右の名列で選んでから写真の上の席を押して置きます（ボタンはドラッグで動かせます）'],
    ['・席がえをしたら「✏ 直す・撮り直す」から撮り直してください。同じクラスの講座は同じ座席表を使います'],
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

// 1回の呼び出しの中で何度も名簿を読まないように、読んだ結果をとっておく（名簿に書きこんだら studentsCache_ = null）
let studentsCache_ = null;

function getStudents_() {
  if (studentsCache_) return studentsCache_;
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
  studentsCache_ = sortStudents_(list);
  return studentsCache_;
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

// ───────── 画面の HTML（Take.html・Setup.html・Period.html の中身） ─────────

var HTML_FILES = {};
HTML_FILES["Take"] = [
  "<!DOCTYPE html>",
  "<html>",
  "<head>",
  "<base target=\"_top\">",
  "<meta charset=\"utf-8\">",
  "<style>",
  "  * { box-sizing: border-box; }",
  "  html, body { height: 100%; }",
  "  body { font-family: \"Noto Sans JP\", sans-serif; margin: 0; font-size: 13px; color: #202124; background: #fff; display: flex; flex-direction: column; overflow: hidden; }",
  "  button { font-family: inherit; cursor: pointer; }",
  "  #top { display: flex; align-items: center; gap: 10px; padding: 6px 10px; border-bottom: 1px solid #dadce0; flex-wrap: wrap; }",
  "  .date { font-size: 15px; font-weight: bold; white-space: nowrap; }",
  "  .date.other { color: #b06000; }",
  "  .dates { display: flex; align-items: center; gap: 2px; }",
  "  .dates input { font-family: inherit; font-size: 13px; padding: 3px 4px; border: 1px solid #dadce0; border-radius: 6px; }",
  "  .dbtn { border: 1px solid #dadce0; background: #fff; border-radius: 6px; padding: 3px 7px; font-size: 11px; }",
  "  .lnk { border: none; background: none; color: #1a73e8; font-size: 12px; padding: 2px 4px; }",
  "  .courses { display: flex; gap: 4px; flex-wrap: wrap; flex: 1; }",
  "  .course { padding: 5px 12px; border: 1px solid #dadce0; border-radius: 16px; background: #fff; font-size: 13px; }",
  "  .course:hover { background: #f1f3f4; }",
  "  .course.on { background: #1a73e8; color: #fff; border-color: #1a73e8; font-weight: bold; }",
  "  .menu { display: flex; gap: 2px; }",
  "  #body { flex: 1; display: flex; min-height: 0; }",
  "  #left { flex: 1.6; border-right: 1px solid #dadce0; overflow: auto; padding: 8px 10px; position: relative; }",
  "  #right { flex: 1; min-width: 300px; overflow: auto; padding: 8px 10px; }",
  "  #foot { display: flex; align-items: center; gap: 12px; padding: 8px 10px; border-top: 1px solid #dadce0; background: #fff; }",
  "  #foot .cnt { flex: 1; font-size: 13px; }",
  "  #foot .cnt b { color: #d93025; font-size: 17px; }",
  "  .save { padding: 9px 26px; border: none; border-radius: 8px; background: #1a73e8; color: #fff; font-size: 15px; font-weight: bold; }",
  "  .save:disabled { opacity: .5; cursor: default; }",
  "  .btn { padding: 6px 14px; border-radius: 8px; border: 1px solid #1a73e8; background: #fff; color: #1a73e8; font-size: 13px; }",
  "  .btn.primary { background: #1a73e8; color: #fff; }",
  "  .btn:disabled { opacity: .5; cursor: default; }",
  "  h3 { font-size: 13px; margin: 2px 0 6px; display: flex; align-items: center; gap: 8px; }",
  "  h3 small { color: #5f6368; font-weight: normal; }",
  "  h3 .sp { flex: 1; }",
  "  .info { background: #f1f3f4; border-radius: 8px; padding: 6px 8px; margin-bottom: 6px; font-size: 12px; line-height: 1.6; }",
  "  .info.edit { background: #fef7e0; color: #7a4f01; }",
  "  .info button { font-size: 12px; margin: 2px 4px 0 0; padding: 2px 10px; border-radius: 12px; border: 1px solid #1a73e8; background: #fff; color: #1a73e8; }",
  "  .info button.on { background: #1a73e8; color: #fff; }",
  "  .opt { display: flex; align-items: center; gap: 6px; margin: 4px 0 6px; }",
  "  .legend { font-size: 11px; color: #5f6368; margin: 0 0 6px; display: flex; flex-wrap: wrap; gap: 4px; align-items: center; }",
  "  .legend span { padding: 0 6px; border-radius: 8px; }",
  "  .list { display: grid; grid-template-columns: 1fr; gap: 3px; }",
  "  .stu { display: flex; align-items: center; gap: 6px; width: 100%; padding: 5px 8px; border: 1px solid #dadce0; border-radius: 6px; background: #fff; text-align: left; font-size: 13px; }",
  "  .stu:hover { border-color: #9aa0a6; }",
  "  .stu .no { width: 2.4em; text-align: right; color: #5f6368; font-size: 11px; flex: 0 0 auto; }",
  "  .stu .nm { flex: 1; font-size: 14px; }",
  "  .stu .st { font-size: 11px; padding: 1px 6px; border-radius: 8px; white-space: nowrap; color: #5f6368; }",
  "  .stu .mk { width: 1.6em; text-align: center; font-weight: bold; flex: 0 0 auto; }",
  "  .stu.abs { background: #d93025; border-color: #d93025; color: #fff; }",
  "  .stu.abs .no { color: #fde7e5; }",
  "  .stu.abs .nm { font-weight: bold; }",
  "  .stu.sel { outline: 3px solid #fbbc04; }",
  "  .stu .placed { font-size: 11px; color: #188038; }",
  "  .stu .unplaced { font-size: 11px; color: #b06000; }",
  "  .msg { padding: 8px; border-radius: 8px; margin: 6px 0; font-size: 12px; line-height: 1.6; }",
  "  .msg.ok { background: #e6f4ea; color: #137333; }",
  "  .msg.warn { background: #fef7e0; color: #b06000; }",
  "  .empty { color: #5f6368; padding: 8px 0; }",
  "",
  "  /* 座席表 */",
  "  .seatbox { position: relative; display: inline-block; max-width: 100%; user-select: none; line-height: 0; }",
  "  .seatbox img { max-width: 100%; max-height: calc(100vh - 170px); display: block; border-radius: 6px; }",
  "  .seatbox.edit img { cursor: crosshair; }",
  "  .spot { position: absolute; transform: translate(-50%, -50%); line-height: 1.2; padding: 3px 8px; border-radius: 12px; border: 2px solid #1a73e8; background: rgba(255,255,255,.88); color: #174ea6; font-size: 12px; font-weight: bold; white-space: nowrap; box-shadow: 0 1px 3px rgba(0,0,0,.3); }",
  "  .spot:hover { background: #e8f0fe; }",
  "  .spot.abs { background: #d93025; border-color: #d93025; color: #fff; }",
  "  .spot.sel { outline: 3px solid #fbbc04; }",
  "  .seatbox.edit .spot { cursor: grab; }",
  "  .seatbox.dim .spot { opacity: .35; }",
  "  .vtabs { display: inline-flex; border: 1px solid #1a73e8; border-radius: 8px; overflow: hidden; }",
  "  .vtabs button { border: none; background: #fff; color: #1a73e8; padding: 5px 12px; font-size: 12px; }",
  "  .vtabs button.on { background: #1a73e8; color: #fff; font-weight: bold; }",
  "  .grid { display: grid; gap: 6px; margin: 4px 0; }",
  "  .seat { min-height: 62px; border: 1px solid #c4c7c5; border-radius: 8px; background: #fff; padding: 4px 2px; text-align: center; display: flex; flex-direction: column; align-items: center; justify-content: center; line-height: 1.25; }",
  "  button.seat:hover { border-color: #1a73e8; background: #f8fbff; }",
  "  .seat .sno { font-size: 11px; color: #5f6368; }",
  "  .seat .snm { font-size: 14px; font-weight: bold; word-break: keep-all; }",
  "  .seat .sst { font-size: 10px; padding: 0 5px; border-radius: 6px; margin-top: 2px; color: #5f6368; }",
  "  .seat.abs, button.seat.abs:hover { background: #d93025; border-color: #d93025; color: #fff; }",
  "  .seat.abs .sno, .seat.abs .sst { color: #fde7e5; background: none !important; }",
  "  .seat.empty { background: #f8f9fa; border-style: dashed; border-color: #e0e0e0; }",
  "  .seat.multi { padding: 2px; gap: 2px; }",
  "  .seat.multi button { width: 100%; border: 1px solid #dadce0; border-radius: 6px; background: #fff; font-size: 12px; padding: 2px; }",
  "  .seat.multi button.abs { background: #d93025; color: #fff; border-color: #d93025; }",
  "  .grid.small .seat { min-height: 34px; padding: 2px; }",
  "  .grid.small .snm { font-size: 11px; }",
  "  .grid.small .sno, .grid.small .sst { display: none; }",
  "  .grid.dense .snm { font-size: 12px; }",
  "  .desk { text-align: center; margin: 6px auto; width: 30%; min-width: 120px; border: 2px solid #5f6368; border-radius: 6px; padding: 3px; font-weight: bold; color: #3c4043; background: #f1f3f4; font-size: 13px; }",
  "  .place { border: 2px dashed #dadce0; border-radius: 10px; padding: 24px; text-align: center; color: #5f6368; line-height: 1.8; }",
  "  .place .btn { margin: 4px; font-size: 14px; padding: 8px 18px; }",
  "  .tools { display: flex; gap: 6px; flex-wrap: wrap; align-items: center; margin-bottom: 6px; }",
  "  video { max-width: 100%; max-height: calc(100vh - 190px); border-radius: 8px; background: #000; display: block; }",
  "  .bar { height: 6px; background: #eee; border-radius: 3px; overflow: hidden; margin-top: 4px; }",
  "  .bar span { display: block; height: 100%; background: #1a73e8; width: 0; transition: width .2s; }",
  "  .ver { font-size: 10px; color: #9aa0a6; margin-top: 10px; }",
  "</style>",
  "</head>",
  "<body>",
  "<div id=\"top\">",
  "  <span class=\"date\" id=\"dateLabel\">…</span>",
  "  <span class=\"dates\"><button class=\"dbtn\" id=\"prev\" title=\"前の日\">◀</button><input type=\"date\" id=\"dateInput\"><button class=\"dbtn\" id=\"next\" title=\"次の日\">▶</button><button class=\"lnk\" id=\"today\" style=\"display:none\">今日にもどす</button></span>",
  "  <div class=\"courses\" id=\"courses\"></div>",
  "  <div class=\"menu\">",
  "    <button class=\"lnk\" id=\"lPeriod\">🔎 期間集計</button>",
  "    <button class=\"lnk\" id=\"lBulk\">📝 まとめて入力</button>",
  "    <button class=\"lnk\" id=\"lSetup\">⚙ 初期設定</button>",
  "    <button class=\"lnk\" id=\"lReload\" title=\"初期設定を変えたあとに押す\">🔄</button>",
  "  </div>",
  "</div>",
  "<div id=\"body\">",
  "  <div id=\"left\"><div class=\"empty\">読み込み中…</div></div>",
  "  <div id=\"right\"></div>",
  "</div>",
  "<div id=\"foot\" style=\"display:none\">",
  "  <div class=\"cnt\" id=\"cnt\"></div>",
  "  <button class=\"btn\" id=\"cancelSeat\" style=\"display:none\">やめる</button>",
  "  <button class=\"save\" id=\"save\">教務手帳に書きこむ</button>",
  "</div>",
  "",
  "<script>",
  "  const $ = (id) => document.getElementById(id);",
  "  const run = (name, ...args) => new Promise((ok, ng) => google.script.run.withSuccessHandler(ok).withFailureHandler(ng)[name](...args));",
  "  const COLORS = [{ bg: '#fff2cc', fg: '#7f6000' }, { bg: '#fce5cd', fg: '#b45f06' }, { bg: '#f4cccc', fg: '#990000' }, { bg: '#cc0000', fg: '#ffffff' }];",
  "  const WD = '日月火水木金土';",
  "  let I = null;       // getTakeInit",
  "  let date = '';",
  "  let D = null;       // getCourseDay",
  "  let cols = [];      // 直している列（2時間続きなら2つ。空＝新しく記録）",
  "  let split = false;  // 2時間続きの列も、1コマずつ直す",
  "  let double = false; // 2時間続き",
  "  let absent = new Set();",
  "  let dirty = false;",
  "  let note = '';",
  "  const seats = {};   // 対象 → getSeat の結果",
  "  let mode = 'take';  // 'take'：出欠をとる／'seat'：座席表を登録・直す",
  "  let E = null;       // 座席表を直しているときの状態",
  "  let busy = '';      // 'pick'：講座・日付を読み込み中／'save'：書きこみ中（そのあいだは押しても変えない）",
  "  let pickSeq = 0;    // 講座を続けて押したとき、最後に押したものだけ使う",
  "",
  "  function el(tag, attrs, ...kids) {",
  "    const e = document.createElement(tag);",
  "    Object.entries(attrs || {}).forEach(([k, v]) => {",
  "      if (v == null || v === false) return;",
  "      if (k.startsWith('on')) e.addEventListener(k.slice(2), v);",
  "      else if (k === 'class') e.className = v;",
  "      else if (k === 'style') e.style.cssText = v;",
  "      else if (k in e) e[k] = v;",
  "      else e.setAttribute(k, v);",
  "    });",
  "    kids.flat().forEach((c) => { if (c != null && c !== false) e.append(c.nodeType ? c : String(c)); });",
  "    return e;",
  "  }",
  "  const errText = (e) => (e && e.message) || String(e);",
  "  function dayLabel(k) { const p = k.split('-').map(Number); return p[1] + '/' + p[2] + '（' + WD[new Date(p[0], p[1] - 1, p[2]).getDay()] + '）'; }",
  "  function lineIndex(rate) { let idx = -1; D.lines.forEach((v, i) => { if (rate >= v - 1e-9) idx = i; }); return idx; }",
  "  const pct = (r) => (Math.round(r * 1000) / 10).toFixed(1) + '%';",
  "  const shortName = (n) => String(n).split(/[\\s　]+/)[0] || n;",
  "",
  "  // ───────── 画面全体 ─────────",
  "  function drawTop() {",
  "    $('dateLabel').textContent = '📅 ' + dayLabel(date) + (date === I.today ? '' : ' の分を記録（後日の入力）');",
  "    $('dateLabel').className = 'date' + (date === I.today ? '' : ' other');",
  "    $('dateInput').value = date;",
  "    $('today').style.display = date === I.today ? 'none' : '';",
  "    $('courses').innerHTML = '';",
  "    I.courses.forEach((c) => $('courses').append(el('button', { class: 'course' + (D && D.course === c.name ? ' on' : ''), onclick: () => pick(c.name) }, c.name)));",
  "  }",
  "",
  "  function draw() {",
  "    drawTop();",
  "    if (!I.students || !I.courses.length) {",
  "      $('left').innerHTML = '';",
  "      $('left').append(el('div', { class: 'msg warn' }, !I.students ? 'はじめに「⚙ 初期設定」で名簿を登録してください。' : 'はじめに「⚙ 初期設定」で講座を登録してください。'),",
  "        el('button', { class: 'btn primary', onclick: () => run('openSetup', '').catch((e) => alert(errText(e))) }, '⚙ 初期設定を開く'));",
  "      $('right').innerHTML = '';",
  "      $('foot').style.display = 'none';",
  "      return;",
  "    }",
  "    if (!D) {",
  "      $('left').innerHTML = '';",
  "      $('left').append(el('div', { class: 'place' }, '上の講座を押してください。', el('br'), '日付は今日が入っています。'), el('div', { class: 'ver' }, '版 ' + I.version));",
  "      $('right').innerHTML = '';",
  "      $('foot').style.display = 'none';",
  "      return;",
  "    }",
  "    if (mode === 'seat') { drawSeatEdit(); return; }",
  "    drawSeat();",
  "    drawList();",
  "    drawFoot();",
  "  }",
  "",
  "  async function pick(name, keepCol) {",
  "    if (busy === 'save') return;",
  "    if (mode === 'seat' && !confirm('座席表を保存せずに講座を変えますか？')) return;",
  "    if (dirty && D && !confirm(D.course !== name ? '書きこんでいない変更があります。捨てて講座を変えますか？' : '書きこんでいない変更があります。捨てて読みなおしますか？')) return;",
  "    const wasSeat = mode === 'seat';",
  "    stopCamera();",
  "    mode = 'take'; E = null;",
  "    if (wasSeat) draw();",
  "    const my = ++pickSeq;",
  "    busy = 'pick';",
  "    $('save').disabled = true;",
  "    let d;",
  "    try { d = await run('getCourseDay', name, date); } catch (e) {",
  "      if (my !== pickSeq) return;",
  "      busy = '';",
  "      if (D && D.date) date = D.date; // 日付を変えて読めなかったら、もとの日付にもどす",
  "      alert(errText(e));",
  "      draw();",
  "      return;",
  "    }",
  "    if (my !== pickSeq) return; // あとから別の講座・日付を押した",
  "    busy = '';",
  "    D = d;",
  "    note = '';",
  "    split = false;",
  "    startEdit(keepCol);",
  "    loadSeat(D.target);",
  "  }",
  "",
  "  // その日の授業（2時間続きは2列で1つ）。「1コマずつ直す」のときは1列ずつ",
  "  function units() {",
  "    if (split) return D.dayCols.map((c) => [c]);",
  "    return D.groups || D.dayCols.map((c) => [c]);",
  "  }",
  "  // その日の列があれば、それを直す（いちばん右の授業）。なければ新しく記録",
  "  function startEdit(keepCol) {",
  "    const us = units();",
  "    cols = (keepCol && us.find((u) => u.includes(keepCol))) || us[us.length - 1] || [];",
  "    double = false;",
  "    loadMarks();",
  "    draw();",
  "  }",
  "  function setCols(u) {",
  "    if (dirty && !confirm('書きこんでいない変更を捨てますか？')) return;",
  "    cols = u; loadMarks(); draw();",
  "  }",
  "  const isAbs = (s) => cols.some((c) => s.day[c] === '欠');",
  "  function loadMarks() {",
  "    absent = new Set(D.students.filter(isAbs).map((s) => s.key));",
  "    dirty = false;",
  "  }",
  "  // 2時間続きの2列で、欠席がちがう生徒（シートを直接直したときなど）",
  "  const mismatch = () => cols.length > 1 ? D.students.filter((s) => cols.some((c) => (s.day[c] === '欠') !== (s.day[cols[0]] === '欠'))) : [];",
  "  function toggle(key) {",
  "    if (busy) return; // 書きこみ中・読み込み中は変えない",
  "    // キーボードで押したときも、かき直したあと同じ生徒にフォーカスをもどす",
  "    const a = document.activeElement;",
  "    const side = a && a.getAttribute && a.getAttribute('data-k') === key ? (a.closest('#left') ? 'left' : 'right') : '';",
  "    absent.has(key) ? absent.delete(key) : absent.add(key);",
  "    dirty = true; note = '';",
  "    drawSeat(); drawList(); drawFoot();",
  "    if (side) { const b = Array.from($(side).querySelectorAll('[data-k]')).find((x) => x.getAttribute('data-k') === key); if (b) b.focus({ preventScroll: true }); }",
  "  }",
  "",
  "  // 生徒ごとの欠課時数（今つけている分もふくむ）",
  "  function statOf(s) {",
  "    const per = cols.length || (double ? 2 : 1);",
  "    const others = s.before + s.absentAll - cols.filter((c) => s.day[c] === '欠').length;",
  "    const total = others + (absent.has(s.key) ? per : 0);",
  "    const rate = D.hours ? total / D.hours : 0;",
  "    const li = D.hours ? lineIndex(rate) : -1;",
  "    return { total, rate, color: li >= 0 ? COLORS[Math.min(li, COLORS.length - 1)] : null };",
  "  }",
  "",
  "  // ───────── 右：名列 ─────────",
  "  function drawList() {",
  "    const R = $('right');",
  "    const y = R.scrollTop;",
  "    R.innerHTML = '';",
  "    const info = el('div', { class: 'info' + (cols.length ? ' edit' : '') });",
  "    // 「1コマ目」「2・3コマ目（2時間続き）」のような名前",
  "    const label = (u) => u.map((c) => D.dayCols.indexOf(c) + 1).join('・') + 'コマ目' + (u.length > 1 ? '（2時間続き）' : '');",
  "    const us = units();",
  "    const hasDouble = (D.groups || []).some((g) => g.length > 1);",
  "    if (cols.length) {",
  "      info.append(dayLabel(date) + 'の分は記録ずみです' + (D.dayCols.length > 1 ? '（' + label(cols) + '）' : '') + '。直して書きこむと上書きします。', el('br'));",
  "      if (cols.length > 1) info.append(el('b', null, '2時間続きの2列をまとめて直します。'), el('br'));",
  "      if (us.length > 1) us.forEach((u) => info.append(el('button', { class: u.join() === cols.join() ? 'on' : '', onclick: () => setCols(u) }, label(u))));",
  "      if (hasDouble) info.append(el('button', { onclick: () => { if (dirty && !confirm('書きこんでいない変更を捨てますか？')) return; split = !split; startEdit(cols[0]); } }, split ? '2時間続きをまとめて直す' : '1コマずつ直す'));",
  "      info.append(el('button', { onclick: () => setCols([]) }, '＋ この日にもう1コマ記録する'));",
  "      const mm = mismatch();",
  "      if (mm.length) info.append(el('div', { class: 'msg warn' }, '2列で欠席がちがう生徒がいます（' + mm.map((s) => s.name).join('、') + '）。どちらかの列に「欠」があれば欠席として出しています。書きこむと2列とも同じになります。'));",
  "    } else {",
  "      info.append(D.dayCols.length ? dayLabel(date) + 'の' + (D.dayCols.length + 1) + 'コマ目として、新しい列を作ります。' : '欠席した生徒を押してください（座席表でも名列でも）。全員出席なら、そのまま書きこみます。');",
  "      if (D.dayCols.length) info.append(el('br'), el('button', { onclick: () => { cols = us[us.length - 1]; loadMarks(); draw(); } }, '記録ずみの列を直す'));",
  "    }",
  "    R.append(info);",
  "    if (!cols.length) R.append(el('label', { class: 'opt' }, el('input', { type: 'checkbox', checked: double, onchange: (e) => { double = e.target.checked; draw(); } }), '2時間続き（2コマ分として書きこむ）'));",
  "    R.append(el('div', { class: 'legend' }, '年間 ' + D.hours + '時間 ・ 色：', D.lines.map((v, i) => { const c = COLORS[Math.min(i, COLORS.length - 1)]; return el('span', { style: 'background:' + c.bg + ';color:' + c.fg }, pct(v).replace('.0', '')); })));",
  "    if (note) R.append(el('div', { class: 'msg ok' }, note));",
  "    if (!D.students.length) R.append(el('div', { class: 'empty' }, '対象の生徒がいません。「⚙ 初期設定」で講座の対象クラスを確かめてください。'));",
  "    const list = el('div', { class: 'list' });",
  "    D.students.forEach((s) => {",
  "      const on = absent.has(s.key);",
  "      const st = statOf(s);",
  "      const v0 = cols.length ? s.day[cols[0]] : '';",
  "      const other = v0 && v0 !== '欠' ? v0 : '';",
  "      list.append(el('button', {",
  "        class: 'stu' + (on ? ' abs' : ''), 'data-k': s.key,",
  "        title: '欠課時数 ' + st.total + '（それ以前 ' + s.before + '）' + (D.hours ? ' ・ ' + pct(st.rate) : ''),",
  "        onclick: () => toggle(s.key),",
  "      },",
  "        el('span', { class: 'no' }, s.no !== '' ? s.no : ''),",
  "        el('span', { class: 'nm' }, s.name, other ? '（' + other + '）' : ''),",
  "        el('span', { class: 'st', style: st.color ? 'background:' + st.color.bg + ';color:' + st.color.fg : '' }, '欠' + st.total + (D.hours ? ' ' + pct(st.rate) : '')),",
  "        el('span', { class: 'mk' }, on ? '欠' : '')));",
  "    });",
  "    R.append(list, el('div', { class: 'ver' }, '版 ' + I.version));",
  "    R.scrollTop = y;",
  "  }",
  "",
  "  function drawFoot() {",
  "    $('foot').style.display = '';",
  "    $('cancelSeat').style.display = 'none';",
  "    $('cnt').innerHTML = '';",
  "    $('cnt').append('欠席 ', el('b', null, absent.size), ' 人 ・ 出席 ' + (D.students.length - absent.size) + ' 人' + (cols.length > 1 || (!cols.length && double) ? '（2コマ分）' : ''));",
  "    $('save').textContent = cols.length ? '教務手帳を直す（上書き）' : '教務手帳に書きこむ';",
  "    $('save').disabled = !!busy;",
  "  }",
  "",
  "  // ───────── 左：座席表（出欠をとるとき） ─────────",
  "  async function loadSeat(target) {",
  "    if (seats[target] && !seats[target].error) { if (mode === 'take') drawSeat(); return; } // 読めなかったときは読み直す",
  "    seats[target] = { loading: true };",
  "    drawSeat();",
  "    try { seats[target] = await run('getSeat', target); }",
  "    catch (e) { seats[target] = { error: errText(e), image: '', spots: [] }; }",
  "    if (D && D.target === target && mode === 'take') drawSeat();",
  "  }",
  "",
  "  function drawSeat() {",
  "    const L = $('left');",
  "    L.innerHTML = '';",
  "    const seat = seats[D.target];",
  "    if (!seat || seat.loading) { L.append(el('div', { class: 'empty' }, '座席表を読み込み中…')); return; }",
  "    if (!seat.image) {",
  "      L.append(el('div', { class: 'place' },",
  "        seat.error ? el('div', { class: 'msg warn' }, '座席表を読み込めませんでした：' + seat.error) : null,",
  "        seat.missing ? el('div', { class: 'msg warn' }, '座席表の写真がドライブから消されています。撮り直してください。') : null,",
  "        el('b', null, D.target + ' の座席表がまだありません'), el('br'),",
  "        '座席表を撮って登録すると、写真の上の名前を押して欠席にできます。', el('br'),",
  "        el('button', { class: 'btn primary', onclick: () => startSeatEdit() }, '📷 座席表を登録する'), el('br'),",
  "        el('small', null, '登録しなくても、右の名列から押して記録できます。')));",
  "      return;",
  "    }",
  "    const byKey = Object.fromEntries(D.students.map((s) => [s.key, s]));",
  "    const placed = seat.spots.filter((p) => byKey[p.key]);",
  "    const notPlaced = D.students.filter((s) => !seat.spots.some((p) => p.key === s.key));",
  "    const setView = (o) => { Object.assign(seat, o); drawSeat(); run('saveSeatView', D.target, o).catch(() => {}); };",
  "    L.append(el('h3', null, '座席表', el('small', null, D.target + (seat.updated ? ' ・ ' + seat.updated + ' 更新' : '')), el('span', { class: 'sp' }),",
  "      el('span', { class: 'vtabs' },",
  "        el('button', { class: seat.view !== 'photo' ? 'on' : '', onclick: () => setView({ view: 'grid' }) }, '見やすい表'),",
  "        el('button', { class: seat.view === 'photo' ? 'on' : '', onclick: () => setView({ view: 'photo' }) }, '写真')),",
  "      seat.view !== 'photo' ? el('button', { class: 'btn', title: '上下・左右をさかさまにする（教卓から見た向き／生徒から見た向き）', onclick: () => setView({ flip: !seat.flip }) }, '↕ 向きを変える') : null,",
  "      el('button', { class: 'btn', onclick: () => startSeatEdit() }, '✏ 直す・撮り直す')));",
  "    if (seat.view === 'photo') {",
  "      // 写真は1回だけ作って使いまわす（押すたびに大きな写真を読み直さない）",
  "      if (!seat.imgEl) seat.imgEl = el('img', { src: seat.image, alt: '座席表', draggable: false });",
  "      const box = el('div', { class: 'seatbox' }, seat.imgEl);",
  "      placed.forEach((p) => {",
  "        const s = byKey[p.key];",
  "        box.append(el('button', {",
  "          class: 'spot' + (absent.has(s.key) ? ' abs' : ''), 'data-k': s.key,",
  "          style: 'left:' + (p.x * 100) + '%;top:' + (p.y * 100) + '%',",
  "          title: s.name, onclick: () => toggle(s.key),",
  "        }, (absent.has(s.key) ? '欠 ' : '') + shortName(s.name)));",
  "      });",
  "      L.append(box);",
  "    } else {",
  "      L.append(gridView(placed, seat.desk, seat.flip, false));",
  "    }",
  "    if (notPlaced.length) L.append(el('div', { class: 'msg warn' }, '座席表に置いていない生徒：' + notPlaced.map((s) => s.name).join('、') + '（右の名列から押せます）'));",
  "  }",
  "",
  "  // 見やすい表：写真の上のボタンの位置から、行と列にそろえて並べなおした座席表",
  "  //  small：座席表を直すときの下の見本（押せない）",
  "  function gridView(spots, desk, flip, small) {",
  "    const wrap = el('div');",
  "    const g = seatGrid(spots);",
  "    if (!g) { wrap.append(el('div', { class: 'empty' }, 'まだボタンが置かれていません。')); return wrap; }",
  "    const byKey = Object.fromEntries(D.students.map((s) => [s.key, s]));",
  "    // 教卓が写真の上にあれば表の上に出す（向きを変えたら反対側）",
  "    const deskTop = desk ? (desk === 'top') !== !!flip : null;",
  "    const deskEl = () => el('div', { class: 'desk' }, '教卓');",
  "    if (deskTop === true) wrap.append(deskEl());",
  "    const grid = el('div', { class: 'grid' + (small ? ' small' : '') + (g.cols >= 8 ? ' dense' : ''), style: 'grid-template-columns:repeat(' + g.cols + ', minmax(0, 1fr))' });",
  "    for (let rr = 0; rr < g.rows; rr++) {",
  "      for (let cc = 0; cc < g.cols; cc++) {",
  "        const r = flip ? g.rows - 1 - rr : rr, c = flip ? g.cols - 1 - cc : cc;",
  "        const keys = (g.at[r + ',' + c] || []).filter((k) => byKey[k]);",
  "        if (!keys.length) { grid.append(el('div', { class: 'seat empty' })); continue; }",
  "        if (keys.length > 1) {",
  "          // 同じマスに2人（ボタンの位置が近すぎる）：両方出す",
  "          grid.append(el('div', { class: 'seat multi', title: '同じ席に2人います。「直す・撮り直す」で位置を直してください' }, keys.map((k) => el('button', {",
  "            class: !small && absent.has(k) ? 'abs' : '', disabled: small, 'data-k': small ? null : k, onclick: () => toggle(k),",
  "          }, shortName(byKey[k].name)))));",
  "          continue;",
  "        }",
  "        const s = byKey[keys[0]];",
  "        const on = !small && absent.has(s.key);",
  "        const st = small ? null : statOf(s);",
  "        grid.append(el(small ? 'div' : 'button', {",
  "          class: 'seat' + (on ? ' abs' : ''), title: s.name, 'data-k': small ? null : s.key, onclick: small ? null : () => toggle(s.key),",
  "        },",
  "          el('span', { class: 'sno' }, s.no !== '' ? String(s.no).padStart(2, '0') : ''),",
  "          el('span', { class: 'snm' }, on ? '欠 ' + shortName(s.name) : s.name),",
  "          st ? el('span', { class: 'sst', style: st.color && !on ? 'background:' + st.color.bg + ';color:' + st.color.fg : '' }, '欠' + st.total + (D.hours ? ' ' + pct(st.rate) : '')) : null));",
  "      }",
  "    }",
  "    wrap.append(grid);",
  "    if (deskTop === false) wrap.append(deskEl());",
  "    return wrap;",
  "  }",
  "",
  "  // ───────── 座席表を登録・直す ─────────",
  "  function startSeatEdit() {",
  "    if (busy) return;",
  "    if (dirty &&!confirm('書きこんでいない出欠があります。座席表を直すと、いま押した出欠は消えます。よろしいですか？')) return;",
  "    const seat = seats[D.target] || {};",
  "    E = { image: seat.image || '', newImage: '', spots: new Map((seat.spots || []).map((p) => [p.key, { x: p.x, y: p.y }])), desk: seat.desk || '', sel: '', status: '', progress: -1, stream: null, camera: false };",
  "    mode = 'seat';",
  "    loadMarks();",
  "    E.sel = (D.students.find((s) => !E.spots.has(s.key)) || {}).key || '';",
  "    draw();",
  "  }",
  "",
  "  function drawSeatEdit() {",
  "    const L = $('left');",
  "    L.innerHTML = '';",
  "    E.ui = null;",
  "    const reading = E.progress >= 0; // 自動で読み取っている最中（写真は変えられない）",
  "    L.append(el('h3', null, '座席表を登録・直す', el('small', null, D.target + '（同じクラスの講座で共通）')));",
  "    if (E.camera) {",
  "      const v = el('video', { autoplay: true, playsInline: true, muted: true });",
  "      v.srcObject = E.stream;",
  "      L.append(el('div', { class: 'tools' },",
  "        el('button', { class: 'btn primary', onclick: () => shoot(v) }, '📸 撮る'),",
  "        el('button', { class: 'btn', onclick: () => { stopCamera(); drawSeatEdit(); } }, 'やめる'),",
  "        el('small', null, '座席表の全体が入るように写して、「撮る」を押してください。')), v);",
  "    } else if (!E.image) {",
  "      L.append(el('div', { class: 'place' },",
  "        el('b', null, '座席表の写真を用意してください'), el('br'),",
  "        el('button', { class: 'btn primary', onclick: startCamera }, '📷 カメラで撮る'),",
  "        el('button', { class: 'btn', onclick: pickFile }, '🖼 写真のファイルを選ぶ'), el('br'),",
  "        el('small', null, '写真をこの画面に貼り付け（Ctrl + V）・ドラッグしてもOKです。'), el('br'),",
  "        el('small', null, '名前を書いた座席表を撮ると、名前を自動で読み取ってボタンを置きます。')));",
  "    } else {",
  "      L.append(el('div', { class: 'tools' },",
  "        el('button', { class: 'btn primary', onclick: autoRead, disabled: reading }, '🔤 名前を自動で読み取る'),",
  "        el('button', { class: 'btn', onclick: startCamera, disabled: reading }, '📷 撮り直す'),",
  "        el('button', { class: 'btn', onclick: pickFile, disabled: reading }, '🖼 ほかの写真'),",
  "        el('button', { class: 'btn', onclick: rotate, disabled: reading, title: '写真を右に90度まわす' }, '↻ 回転'),",
  "        E.spots.size ? el('button', { class: 'btn', disabled: reading, onclick: () => { if (confirm('置いたボタンをすべて外しますか？')) { E.spots.clear(); E.sel = (D.students[0] || {}).key || ''; drawSeatEdit(); } } }, 'ボタンを全部外す') : null));",
  "      if (E.status) {",
  "        const txt = el('span', null, E.status);",
  "        const bar = reading ? el('span', { style: 'width:' + Math.round(E.progress * 100) + '%' }) : null;",
  "        L.append(el('div', { class: 'msg ' + (reading ? 'warn' : 'ok') }, txt, bar ? el('div', { class: 'bar' }, bar) : null));",
  "        if (bar) E.ui = { txt, bar };",
  "      }",
  "      // 写真は1回だけ作って使いまわす（かき直すたびに大きな写真を読み直さない）",
  "      if (!E.img || E.imgSrc !== E.image) {",
  "        const img = el('img', { src: E.image, alt: '座席表', draggable: false });",
  "        img.addEventListener('click', (e) => {",
  "          if (!E || E.img !== img) return;",
  "          if (!E.sel) { E.status = '右の名列で、置く生徒を選んでください。'; drawSeatEdit(); return; }",
  "          const r = img.getBoundingClientRect();",
  "          E.spots.set(E.sel, { x: (e.clientX - r.left) / r.width, y: (e.clientY - r.top) / r.height });",
  "          E.sel = nextUnplaced(E.sel);",
  "          if (E.progress < 0) E.status = ''; // 読み取り中は進みぐあいを消さない",
  "          drawSeatEdit();",
  "        });",
  "        E.img = img; E.imgSrc = E.image;",
  "      }",
  "      const img = E.img;",
  "      const box = el('div', { class: 'seatbox edit' }, img);",
  "      const byKey = Object.fromEntries(D.students.map((s) => [s.key, s]));",
  "      E.spots.forEach((p, key) => {",
  "        const s = byKey[key];",
  "        if (!s) return;",
  "        const b = el('div', { class: 'spot' + (E.sel === key ? ' sel' : ''), style: 'left:' + (p.x * 100) + '%;top:' + (p.y * 100) + '%', title: s.name + '（ドラッグで動かす）' }, shortName(s.name));",
  "        dragSpot(b, key, img);",
  "        box.append(b);",
  "      });",
  "      L.append(box);",
  "      // できあがる見やすい表の見本",
  "      if (E.spots.size) {",
  "        const spots = Array.from(E.spots.entries()).map(([key, p]) => ({ key, x: p.x, y: p.y }));",
  "        L.append(el('h3', { style: 'margin-top:10px' }, 'できあがる座席表（見やすい表）', el('small', null, '写真の上のボタンの位置から作ります'), el('span', { class: 'sp' }),",
  "          el('label', { style: 'font-weight:normal;font-size:12px' }, '教卓：',",
  "            el('select', { onchange: (e) => { E.desk = e.target.value; drawSeatEdit(); } },",
  "              [['', 'なし'], ['top', '写真の上'], ['bottom', '写真の下']].map(([v, t]) => el('option', { value: v, selected: E.desk === v }, t))))),",
  "          gridView(spots, E.desk, false, true));",
  "      }",
  "    }",
  "    drawPlaceList();",
  "    $('foot').style.display = '';",
  "    $('cancelSeat').style.display = '';",
  "    $('cnt').innerHTML = '';",
  "    $('cnt').append('座席表に置いた生徒 ', el('b', { style: 'color:#1a73e8' }, Array.from(E.spots.keys()).filter((k) => D.students.some((s) => s.key === k)).length), ' / ' + D.students.length + ' 人');",
  "    $('save').textContent = '座席表を保存する';",
  "    $('save').disabled = !E.image || reading || !!busy;",
  "  }",
  "  // 読み取りの進みぐあいだけを書きかえる（写真や名列はかき直さない）",
  "  function showProgress(me) {",
  "    if (E !== me || mode !== 'seat') return;",
  "    if (me.ui && me.ui.txt.isConnected) {",
  "      me.ui.txt.textContent = me.status;",
  "      me.ui.bar.style.width = Math.round(me.progress * 100) + '%';",
  "    } else drawSeatEdit();",
  "  }",
  "",
  "  // 右：座席に置く生徒を選ぶ",
  "  function drawPlaceList() {",
  "    const R = $('right');",
  "    const y = R.scrollTop;",
  "    R.innerHTML = '';",
  "    R.append(el('div', { class: 'info' }, E.image",
  "      ? '① 右で生徒を選ぶ（黄色の枠）→ ② 写真の上のその生徒の席を押すと、ボタンが置かれます。置いたボタンはドラッグで動かせます。'",
  "      : '左で座席表の写真を用意してください。'));",
  "    const list = el('div', { class: 'list' });",
  "    D.students.forEach((s) => {",
  "      const placed = E.spots.has(s.key);",
  "      list.append(el('button', { class: 'stu' + (E.sel === s.key ? ' sel' : ''), onclick: () => { E.sel = s.key; drawSeatEdit(); } },",
  "        el('span', { class: 'no' }, s.no !== '' ? s.no : ''),",
  "        el('span', { class: 'nm' }, s.name),",
  "        placed ? el('span', { class: 'placed' }, '✓ 置いた') : el('span', { class: 'unplaced' }, 'まだ'),",
  "        placed ? el('span', { class: 'mk', title: 'ボタンを外す', onclick: (e) => { e.stopPropagation(); E.spots.delete(s.key); E.sel = s.key; drawSeatEdit(); } }, '×') : el('span', { class: 'mk' })));",
  "    });",
  "    R.append(list);",
  "    R.scrollTop = y;",
  "  }",
  "",
  "  function nextUnplaced(after) {",
  "    const i = D.students.findIndex((s) => s.key === after);",
  "    const order = D.students.slice(i + 1).concat(D.students.slice(0, i + 1));",
  "    return (order.find((s) => !E.spots.has(s.key)) || {}).key || '';",
  "  }",
  "",
  "  // 置いたボタンをドラッグで動かす（少しも動かさなければ「選ぶ」）",
  "  function dragSpot(b, key, img) {",
  "    b.addEventListener('pointerdown', (e) => {",
  "      e.preventDefault();",
  "      const me = E;",
  "      b.setPointerCapture(e.pointerId);",
  "      const r = img.getBoundingClientRect();",
  "      const sx = e.clientX, sy = e.clientY;",
  "      let moved = false;",
  "      const move = (ev) => {",
  "        if (E !== me) return;",
  "        if (Math.abs(ev.clientX - sx) + Math.abs(ev.clientY - sy) > 3) moved = true;",
  "        if (!moved) return;",
  "        const x = Math.max(0, Math.min(1, (ev.clientX - r.left) / r.width));",
  "        const y = Math.max(0, Math.min(1, (ev.clientY - r.top) / r.height));",
  "        E.spots.set(key, { x, y });",
  "        b.style.left = (x * 100) + '%'; b.style.top = (y * 100) + '%';",
  "      };",
  "      const up = () => {",
  "        b.removeEventListener('pointermove', move);",
  "        b.removeEventListener('pointerup', up);",
  "        b.removeEventListener('pointercancel', up);",
  "        if (E !== me || mode !== 'seat') return; // ドラッグ中に座席表の画面を閉じた",
  "        E.sel = key;",
  "        drawSeatEdit();",
  "      };",
  "      b.addEventListener('pointermove', move);",
  "      b.addEventListener('pointerup', up);",
  "      b.addEventListener('pointercancel', up);",
  "    });",
  "  }",
  "",
  "  // ───────── 写真：カメラ・ファイル・貼り付け ─────────",
  "  async function startCamera() {",
  "    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) { alert('このブラウザではカメラを使えません。「写真のファイルを選ぶ」を使ってください。'); return; }",
  "    const me = E;",
  "    if (me.camera || me.starting) return; // 2回押したとき、カメラを2つ開かない",
  "    me.starting = true;",
  "    let stream;",
  "    try {",
  "      stream = await navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 1920 }, height: { ideal: 1080 }, facingMode: 'environment' }, audio: false });",
  "    } catch (e) {",
  "      me.starting = false;",
  "      if (E === me) alert('カメラを起動できませんでした。\\n・アドレスバー右のカメラのマークで「許可」にしてから、もう一度押してください。\\n・それでもだめなときは、カメラのアプリで撮って「写真のファイルを選ぶ」から選んでください。\\n\\n（' + (e.name || '') + ' ' + (e.message || '') + '）');",
  "      return;",
  "    }",
  "    me.starting = false;",
  "    // 許可を待っているあいだに、座席表の画面を閉じた・読み取りを始めた",
  "    if (E !== me || mode !== 'seat' || me.progress >= 0) { stream.getTracks().forEach((t) => t.stop()); return; }",
  "    me.stream = stream;",
  "    me.camera = true;",
  "    drawSeatEdit();",
  "  }",
  "  function stopCamera() {",
  "    if (E && E.stream) { E.stream.getTracks().forEach((t) => t.stop()); E.stream = null; }",
  "    if (E) E.camera = false;",
  "  }",
  "  function shoot(video) {",
  "    if (!video.videoWidth) { alert('まだカメラの準備ができていません。少し待ってから押してください。'); return; }",
  "    const url = toJpeg(video, video.videoWidth, video.videoHeight);",
  "    stopCamera(); // 先に止める（あとだと、撮ったあともカメラの画面のままになる）",
  "    setImage(url);",
  "  }",
  "  function pickFile() {",
  "    const inp = el('input', { type: 'file', accept: 'image/*' });",
  "    inp.onchange = () => { if (inp.files[0]) readFile(inp.files[0]); };",
  "    inp.click();",
  "  }",
  "  function readFile(file) {",
  "    const me = E;",
  "    const url = URL.createObjectURL(file);",
  "    const img = new Image();",
  "    img.onload = () => {",
  "      URL.revokeObjectURL(url);",
  "      if (E !== me || mode !== 'seat' || me.progress >= 0) return; // 読みこむあいだに画面を変えた・読み取りを始めた",
  "      stopCamera();",
  "      setImage(toJpeg(img, img.naturalWidth, img.naturalHeight));",
  "    };",
  "    img.onerror = () => { URL.revokeObjectURL(url); alert('この写真は読めませんでした（JPEG・PNG の写真を選んでください）'); };",
  "    img.src = url;",
  "  }",
  "  // 長い辺 1600px までに小さくして JPEG にする（保存とやりとりを軽くするため）",
  "  function toJpeg(src, w, h) {",
  "    const k = Math.min(1, 1600 / Math.max(w, h));",
  "    const c = document.createElement('canvas');",
  "    c.width = Math.round(w * k); c.height = Math.round(h * k);",
  "    c.getContext('2d').drawImage(src, 0, 0, c.width, c.height);",
  "    return c.toDataURL('image/jpeg', 0.85);",
  "  }",
  "  function setImage(url) {",
  "    const first = !E.image;",
  "    E.image = url; E.newImage = url;",
  "    E.status = first || !E.spots.size ? '写真を用意しました。「🔤 名前を自動で読み取る」を押すと、ボタンを自動で置きます。' : '写真を変えました。席がえをしたときは「🔤 名前を自動で読み取る」で置き直せます。';",
  "    E.progress = -1;",
  "    drawSeatEdit();",
  "  }",
  "  function rotate() {",
  "    const me = E, src = E.image;",
  "    if (me.rotating || me.progress >= 0) return; // 続けて押したとき、ボタンだけ2回まわさない",
  "    me.rotating = true;",
  "    const img = new Image();",
  "    img.onload = () => {",
  "      me.rotating = false;",
  "      if (E !== me || me.image !== src) return;",
  "      const c = document.createElement('canvas');",
  "      c.width = img.naturalHeight; c.height = img.naturalWidth;",
  "      const g = c.getContext('2d');",
  "      g.translate(c.width, 0); g.rotate(Math.PI / 2);",
  "      g.drawImage(img, 0, 0);",
  "      me.image = me.newImage = c.toDataURL('image/jpeg', 0.85);",
  "      me.spots.forEach((p, k) => me.spots.set(k, { x: 1 - p.y, y: p.x }));",
  "      me.desk = ''; // 90度まわすと教卓の「上・下」が合わなくなるので、選び直してもらう",
  "      drawSeatEdit();",
  "    };",
  "    img.onerror = () => { me.rotating = false; };",
  "    img.src = src;",
  "  }",
  "  document.addEventListener('paste', (e) => {",
  "    if (mode !== 'seat' || !E || !e.clipboardData) return;",
  "    const item = Array.from(e.clipboardData.items || []).find((i) => i.type.startsWith('image/'));",
  "    if (!item) return;",
  "    e.preventDefault();",
  "    if (E.progress >= 0) return; // 読み取り中は写真を変えない",
  "    readFile(item.getAsFile());",
  "  });",
  "  // 座席表の画面でなくても、写真を落としたときに画面が写真に変わってしまわないようにする",
  "  document.addEventListener('dragover', (e) => e.preventDefault());",
  "  document.addEventListener('drop', (e) => {",
  "    e.preventDefault();",
  "    if (mode !== 'seat' || !E || E.progress >= 0) return;",
  "    const f = Array.from(e.dataTransfer.files || []).find((x) => x.type.startsWith('image/'));",
  "    if (f) readFile(f);",
  "  });",
  "",
  "  // ───────── 名前の自動読み取り（Tesseract.js） ─────────",
  "  let tessLoading = null;",
  "  function loadTesseract() {",
  "    if (window.Tesseract) return Promise.resolve();",
  "    if (!tessLoading) {",
  "      tessLoading = new Promise((ok, ng) => {",
  "        const s = document.createElement('script');",
  "        s.src = 'https://cdn.jsdelivr.net/npm/tesseract.js@5.1.1/dist/tesseract.min.js';",
  "        s.onload = ok; s.onerror = () => { tessLoading = null; ng(new Error('読み取りの部品を読み込めませんでした（ネットワークを確かめてください）')); };",
  "        document.head.appendChild(s);",
  "      });",
  "    }",
  "    return tessLoading;",
  "  }",
  "",
  "  async function autoRead() {",
  "    if (E.progress >= 0 || E.camera) return;",
  "    const me = E, d0 = D, src = E.image;",
  "    // 読み取り中に講座を変えた・座席表の画面を閉じた・写真を変えたら、結果は使わない",
  "    const alive = () => E === me && D === d0 && mode === 'seat' && me.image === src;",
  "    const already = me.spots.size;",
  "    let keep = false;",
  "    if (already) keep = !confirm('いま置いてあるボタンも、読み取った位置に置き直しますか？\\n「OK」＝置き直す ／「キャンセル」＝まだ置いていない生徒だけ置く');",
  "    me.progress = 0; me.status = '読み取りの準備をしています…（はじめての時は少しかかります）';",
  "    drawSeatEdit();",
  "    let worker = null;",
  "    try {",
  "      await loadTesseract();",
  "      const img = new Image();",
  "      await new Promise((ok, ng) => { img.onload = ok; img.onerror = () => ng(new Error('写真を読めませんでした')); img.src = src; });",
  "      // 小さい写真だけ大きくしてから読む",
  "      const k = Math.max(1, 1200 / Math.max(img.naturalWidth, img.naturalHeight));",
  "      const c = document.createElement('canvas');",
  "      c.width = Math.round(img.naturalWidth * k); c.height = Math.round(img.naturalHeight * k);",
  "      c.getContext('2d').drawImage(img, 0, 0, c.width, c.height);",
  "      // 読み方を2通り（表の中のばらばらの文字向き・ふつうの文章向き）ためして、見つかった行を合わせる",
  "      // （'12' は向きの判定用のデータ osd がないと警告が出るが、そのまま '11'（ばらばらの文字）として読む。'11' より少しよく読める）",
  "      const passes = ['12', '6'];",
  "      let pass = 0;",
  "      worker = await Tesseract.createWorker('jpn', 1, {",
  "        logger: (m) => {",
  "          if (!alive() || pass >= passes.length) return; // マスの読み直しの進みぐあいは下で出す",
  "          if (m.status === 'recognizing text') {",
  "            me.progress = 0.3 + (pass + m.progress) / passes.length * 0.55;",
  "            me.status = '名前を読み取っています… ' + Math.round((pass + m.progress) / passes.length * 100) + '%';",
  "          } else if (m.progress != null && pass === 0) { me.progress = Math.min(0.3, m.progress * 0.3); me.status = '読み取りの準備をしています…'; }",
  "          else return;",
  "          showProgress(me);",
  "        },",
  "      });",
  "      let lines = [];",
  "      for (pass = 0; pass < passes.length; pass++) {",
  "        if (!alive()) return;",
  "        await worker.setParameters({ tessedit_pageseg_mode: passes[pass] });",
  "        const res = await worker.recognize(c, {}, { blocks: true });",
  "        lines = lines.concat(res.data.lines || []);",
  "      }",
  "      if (!alive()) return;",
  "      // 「教卓」「黒板」「教壇」が写真の上と下のどちらにあるか",
  "      const deskLine = lines.find((l) => /教卓|黒板|教壇/.test(String(l.text || '').replace(/\\s/g, '')));",
  "      let desk = '';",
  "      if (deskLine && deskLine.bbox) desk = (deskLine.bbox.y0 + deskLine.bbox.y1) / 2 < c.height / 2 ? 'top' : 'bottom';",
  "      // 1クラスだけの講座なら、出席番号でも照合する（合同授業は番号が重なるので名前だけ）",
  "      const oneClass = new Set(d0.students.map((s) => s.cls)).size === 1;",
  "      const all = d0.students.map((s) => ({ key: s.key, name: s.name, no: s.no }));",
  "      const found = matchNames(lines, all, { numbers: oneClass, size: Math.max(c.width, c.height) });",
  "      // 読めなかった生徒は、表の空いているマスを1マスずつ読み直す（1マスだけのほうが正しく読める）",
  "      // （番号だけで置いた席は名前より少し上にあるので、表の形は名前で置いた席から見つける）",
  "      const named = found.filter((f) => f.by !== 'no');",
  "      const grid = found.length < all.length ? gridCells(named.length >= 6 ? named : found, c.width, c.height, found) : null;",
  "      if (grid && grid.cells.length) {",
  "        await worker.setParameters({ tessedit_pageseg_mode: '6' });",
  "        for (let i = 0; i < grid.cells.length; i++) {",
  "          if (!alive()) return;",
  "          const rest = all.filter((s) => !found.some((f) => f.key === s.key));",
  "          if (!rest.length) break;",
  "          me.progress = 0.85 + 0.15 * i / grid.cells.length;",
  "          me.status = '空いているマスを1つずつ読み直しています… ' + (i + 1) + '/' + grid.cells.length;",
  "          showProgress(me);",
  "          const cell = grid.cells[i];",
  "          const left = Math.max(0, Math.round(cell.x - grid.dx * 0.55)), top = Math.max(0, Math.round(cell.y - grid.dy * 0.55));",
  "          const rect = { left, top, width: Math.min(c.width - left, Math.round(grid.dx * 1.1)), height: Math.min(c.height - top, Math.round(grid.dy * 1.1)) };",
  "          const r = await worker.recognize(c, { rectangle: rect }, { blocks: true });",
  "          const hit = matchNames(r.data.lines || [], rest, { numbers: oneClass, size: Math.max(c.width, c.height) }).sort((a, b) => b.score - a.score)[0];",
  "          if (hit) found.push({ key: hit.key, x: cell.x, y: cell.y, score: hit.score, by: hit.by === 'no' ? 'no' : 'cell' });",
  "        }",
  "        // 番号だけで置いた席に、マスの読み直しで別の生徒が見つかったら、番号のほうはまちがい",
  "        const kept = dedupe(found, Math.min(grid.dx, grid.dy) * 0.4);",
  "        found.length = 0;",
  "        kept.forEach((f) => found.push(f));",
  "      }",
  "      if (!alive()) return;",
  "      if (desk) me.desk = desk;",
  "      let put = 0;",
  "      found.forEach((f) => {",
  "        if (keep && me.spots.has(f.key)) return;",
  "        me.spots.set(f.key, { x: f.x / c.width, y: f.y / c.height });",
  "        put++;",
  "      });",
  "      const rest = d0.students.filter((s) => !me.spots.has(s.key));",
  "      me.progress = -1;",
  "      const byNo = found.filter((f) => f.by !== 'name').length;",
  "      me.status = '✓ ' + found.length + '人を読み取り（出席番号で ' + byNo + '人）、' + put + '人のボタンを置きました。' +",
  "        (rest.length ? 'のこり ' + rest.length + '人は、右で選んでから写真の上の席を押して置いてください。' : '全員置けました。') +",
  "        ' 位置がずれているボタンはドラッグで直せます。';",
  "      me.sel = rest.length ? rest[0].key : '';",
  "    } catch (e) {",
  "      if (!alive()) return;",
  "      me.progress = -1;",
  "      me.status = '自動で読み取れませんでした：' + errText(e) + '　右で生徒を選んでから、写真の上の席を押して置いてください。';",
  "    } finally {",
  "      if (worker) worker.terminate().catch(() => {}); // 読み取りの部品（大きなメモリ）を必ず閉じる",
  "      // とちゅうでやめたとき（写真を変えたなど）、ボタンを押せるようにもどす",
  "      if (E === me && me.progress >= 0) { me.progress = -1; me.status = '読み取りをやめました。'; if (mode === 'seat') drawSeatEdit(); }",
  "    }",
  "    if (E === me && mode === 'seat') drawSeatEdit();",
  "  }",
  "",
  "  // <match>",
  "  // OCR の結果（行ごとの単語と位置）から、生徒の出席番号と名前を探して位置を返す",
  "  //  lines: [{ words: [{ text, bbox: { x0, y0, x1, y1 } }] }]、students: [{ key, name, no }]",
  "  //  opt: { numbers: 出席番号でも探すか（1クラスだけの講座）, size: 画像の長い辺 }",
  "  //  返り値: [{ key, x, y, score, by }]（x・y は読み取った画像の中のピクセル、by は 'no'・'name'・'both'）",
  "  function matchNames(lines, students, opt) {",
  "    opt = opt || {};",
  "    const clean = (t) => String(t || '').normalize('NFKC').replace(/[\\s　・,.、。|｜\\[\\]()（）{}「」『』\\-_—―:：;；'\"`~〜!?！？*＊#＃=＝+＋/／\\\\]/g, '');",
  "    const L = lines.map((line) => {",
  "      const chars = [];",
  "      (line.words || []).forEach((w, wi) => {",
  "        const t = Array.from(clean(w.text));",
  "        if (!t.length || !w.bbox) return;",
  "        const b = w.bbox, cw = (b.x1 - b.x0) / t.length;",
  "        t.forEach((ch, i) => chars.push({ ch, w: wi, x0: b.x0 + cw * i, x1: b.x0 + cw * (i + 1), y0: b.y0, y1: b.y1 }));",
  "      });",
  "      return chars;",
  "    }).filter((c) => c.length);",
  "    const vs = students.map((s) => {",
  "      const parts = String(s.name).normalize('NFKC').trim().split(/\\s+/);",
  "      return { key: s.key, full: clean(s.name), fam: parts.length > 1 ? clean(parts[0]) : '', giv: parts.length > 1 ? clean(parts.slice(1).join('')) : '' };",
  "    });",
  "    const count = {};",
  "    vs.forEach((v) => { [v.fam, v.giv].forEach((t) => { if (t) count[t] = (count[t] || 0) + 1; }); });",
  "    // 名前（target）と、行の中の文字の並び（start から len 文字）の似ている度合い",
  "    const score = (target, chars, start, len) => {",
  "      const T = Array.from(target);",
  "      const bag = {};",
  "      T.forEach((c) => { bag[c] = (bag[c] || 0) + 1; });",
  "      let pos = 0, inter = 0;",
  "      for (let i = 0; i < len; i++) {",
  "        const c = chars[start + i].ch;",
  "        if (c === T[i]) pos++;",
  "        if (bag[c]) { bag[c]--; inter++; }",
  "      }",
  "      return Math.max(pos / T.length, 0.9 * inter / Math.max(T.length, len));",
  "    };",
  "    const cands = [];",
  "    vs.forEach((v) => {",
  "      const variants = [[v.full, 1]];",
  "      // 名字だけ・名前だけの座席表でも読めるように（そのクラスに1人だけのときだけ）",
  "      if (v.fam && Array.from(v.fam).length >= 2 && count[v.fam] === 1) variants.push([v.fam, 0.9]);",
  "      if (v.giv && Array.from(v.giv).length >= 2 && count[v.giv] === 1) variants.push([v.giv, 0.85]);",
  "      variants.forEach(([t, w]) => {",
  "        const n = Array.from(t).length;",
  "        if (!n) return;",
  "        const need = n <= 2 ? 1 : n === 3 ? 0.66 : 0.6;",
  "        L.forEach((chars, li) => {",
  "          [n, n - 1, n + 1].forEach((len) => {",
  "            if (len < 2 || len > chars.length) return;",
  "            for (let st = 0; st + len <= chars.length; st++) {",
  "              const sc = score(t, chars, st, len);",
  "              if (sc >= need) cands.push({ key: v.key, li, st, len, score: sc * w + (len === n ? 0.001 : 0) });",
  "            }",
  "          });",
  "        });",
  "      });",
  "    });",
  "    cands.sort((a, b) => b.score - a.score);",
  "    const used = {}, taken = {}, out = [];",
  "    cands.forEach((c) => {",
  "      if (used[c.key]) return;",
  "      for (let i = c.st; i < c.st + c.len; i++) if (taken[c.li + ':' + i]) return;",
  "      used[c.key] = true;",
  "      for (let i = c.st; i < c.st + c.len; i++) taken[c.li + ':' + i] = true;",
  "      const ch = L[c.li].slice(c.st, c.st + c.len);",
  "      const x0 = Math.min(...ch.map((q) => q.x0)), x1 = Math.max(...ch.map((q) => q.x1));",
  "      const y0 = Math.min(...ch.map((q) => q.y0)), y1 = Math.max(...ch.map((q) => q.y1));",
  "      out.push({ key: c.key, x: (x0 + x1) / 2, y: (y0 + y1) / 2, score: c.score, by: 'name' });",
  "    });",
  "    if (!opt.numbers) return out;",
  "",
  "    // 出席番号（1〜2けたの数字）。「1年1組」「3番」のような数字は使わない。同じ番号が2か所にあれば使わない",
  "    const byNo = {};",
  "    students.forEach((s) => { const n = Number(s.no); if (n > 0) byNo[n] = s.key; });",
  "    const seen = {};",
  "    const near = new Set(Array.from('年組番限月日時回列号階棟'));",
  "    // 読み取りがゆらいで同じ数字が2回出ても、近い位置なら同じ1か所として数える",
  "    const tol = (opt.size || 1000) * 0.03;",
  "    L.forEach((chars) => {",
  "      let i = 0;",
  "      while (i < chars.length) {",
  "        if (!/[0-9]/.test(chars[i].ch)) { i++; continue; }",
  "        let j = i;",
  "        // 数字は単語の区切りで分ける（「26」「28」が同じ行に並んでいても、別の番号）",
  "        while (j < chars.length && /[0-9]/.test(chars[j].ch) && chars[j].w === chars[i].w) j++;",
  "        const run = chars.slice(i, j);",
  "        const after = chars[j] ? chars[j].ch : '', before = i > 0 ? chars[i - 1].ch : '';",
  "        if (run.length <= 2 && !near.has(after) && !near.has(before)) {",
  "          const n = Number(run.map((q) => q.ch).join(''));",
  "          const x = (Math.min(...run.map((q) => q.x0)) + Math.max(...run.map((q) => q.x1))) / 2;",
  "          const y = (Math.min(...run.map((q) => q.y0)) + Math.max(...run.map((q) => q.y1))) / 2;",
  "          const list = seen[n] = seen[n] || [];",
  "          if (!list.some((p) => Math.abs(p.x - x) < tol && Math.abs(p.y - y) < tol)) list.push({ x, y, two: run.length === 2 });",
  "        }",
  "        i = j;",
  "      }",
  "    });",
  "    const byKey = {};",
  "    out.forEach((o) => { byKey[o.key] = o; });",
  "    const limit = (opt.size || 1000) * 0.12; // 番号と名前がこれ以上はなれていたら、別の席とみなす",
  "    const offs = [], onlyNo = [];",
  "    Object.keys(seen).forEach((n) => {",
  "      const key = byNo[n];",
  "      if (!key || seen[n].length !== 1) return;",
  "      const p = seen[n][0];",
  "      const nm = byKey[key];",
  "      if (nm && Math.hypot(nm.x - p.x, nm.y - p.y) < limit) { nm.by = 'both'; offs.push([nm.x - p.x, nm.y - p.y]); return; } // 名前の位置のまま",
  "      // 番号だけで置くのは2けた（「04」など）のときだけ。1けたは読みまちがい（9→4 など）が多い",
  "      if (!nm && p.two) onlyNo.push({ key, x: p.x, y: p.y, score: 0.95, by: 'no' });",
  "    });",
  "    // 番号だけで置いた席は、番号から名前までのいつものずれ（両方読めた席から求める）だけ動かして、名前の位置にそろえる",
  "    const med = (a) => { const v = a.slice().sort((x, y) => x - y); return v.length ? v[Math.floor(v.length / 2)] : 0; };",
  "    const ox = med(offs.map((o) => o[0])), oy = med(offs.map((o) => o[1]));",
  "    onlyNo.forEach((o) => { o.x += ox; o.y += oy; byKey[o.key] = o; });",
  "    return dedupe(Object.values(byKey), limit * 0.25);",
  "  }",
  "",
  "  // 近すぎる（同じ席の）2人は、たしかなほう（番号と名前の両方 → 名前・マス → 番号だけ）を残す",
  "  function dedupe(list, dist) {",
  "    const rank = { both: 3, name: 2, cell: 2, no: 1 };",
  "    const sorted = list.slice().sort((a, b) => (rank[b.by] || 2) - (rank[a.by] || 2) || b.score - a.score);",
  "    const out = [];",
  "    sorted.forEach((o) => { if (!out.some((q) => Math.hypot(q.x - o.x, q.y - o.y) < dist)) out.push(o); });",
  "    return out;",
  "  }",
  "",
  "  // 読めた席の位置から、座席表の行と列を見つけて、まだだれも置いていないマスの中心を返す",
  "  //  points: [{ x, y }]（ピクセル）、W・H: 画像の大きさ",
  "  //  返り値: { cells: [{ x, y }], dx, dy }（dx・dy は1マスの幅・高さ）。表の形が見つからなければ null",
  "  //  occupied：すでに置いた席（表の形を見つけるのには使わないが、空いているかどうかには使う）",
  "  function gridCells(points, W, H, occupied) {",
  "    if (points.length < 6) return null;",
  "    const groups = (vals, tol) => {",
  "      const v = vals.slice().sort((a, b) => a - b);",
  "      const out = [];",
  "      v.forEach((x) => {",
  "        const g = out[out.length - 1];",
  "        if (g && x - g.last <= tol) { g.sum += x; g.n++; g.last = x; } else out.push({ sum: x, n: 1, last: x });",
  "      });",
  "      return out.map((g) => g.sum / g.n);",
  "    };",
  "    const step = (cs) => {",
  "      const d = [];",
  "      for (let i = 1; i < cs.length; i++) d.push(cs[i] - cs[i - 1]);",
  "      d.sort((a, b) => a - b);",
  "      return d.length ? d[Math.floor(d.length / 2)] : 0;",
  "    };",
  "    let cols = groups(points.map((p) => p.x), W * 0.035);",
  "    let rows = groups(points.map((p) => p.y), H * 0.035);",
  "    const dx = step(cols), dy = step(rows);",
  "    if (cols.length < 2 || rows.length < 2 || !dx || !dy) return null;",
  "    // 1列・1行まるごと読めなかったところ（間があいている）をうめる",
  "    const fill = (cs, d) => {",
  "      const out = [cs[0]];",
  "      for (let i = 1; i < cs.length; i++) {",
  "        const k = Math.round((cs[i] - cs[i - 1]) / d);",
  "        for (let j = 1; j < k; j++) out.push(cs[i - 1] + (cs[i] - cs[i - 1]) * j / k);",
  "        out.push(cs[i]);",
  "      }",
  "      return out;",
  "    };",
  "    cols = fill(cols, dx); rows = fill(rows, dy);",
  "    const cells = [];",
  "    rows.forEach((y) => cols.forEach((x) => {",
  "      if (!(occupied || points).some((p) => Math.abs(p.x - x) < dx * 0.45 && Math.abs(p.y - y) < dy * 0.45)) cells.push({ x, y });",
  "    }));",
  "    return { cells, dx, dy };",
  "  }",
  "",
  "  // ボタンの位置（写真の左上を 0、右下を 1）から、座席表の行と列を見つける",
  "  //  返り値: { rows, cols, at: { 'r,c': [key] } }（ボタンがなければ null）",
  "  function seatGrid(spots) {",
  "    if (!spots.length) return null;",
  "    const centers = (vals) => {",
  "      const v = vals.slice().sort((a, b) => a - b);",
  "      // 席と席の間（大きいすき間）の半分くらいまでは同じ列とみなす（番号の位置と名前の位置の差はまとめる）",
  "      const all = [];",
  "      for (let i = 1; i < v.length; i++) all.push(v[i] - v[i - 1]);",
  "      const max = Math.max(0, ...all);",
  "      const gaps = all.filter((g) => g > max * 0.4).sort((a, b) => a - b);",
  "      const typical = gaps.length ? gaps[Math.floor(gaps.length / 2)] : 1;",
  "      const tol = Math.max(0.015, typical * 0.45);",
  "      const out = [];",
  "      v.forEach((x) => {",
  "        const g = out[out.length - 1];",
  "        if (g && x - g.last <= tol) { g.sum += x; g.n++; g.last = x; } else out.push({ sum: x, n: 1, last: x });",
  "      });",
  "      const cs = out.map((g) => g.sum / g.n);",
  "      // 1列・1行まるごとだれもいないところ（間があいている）をうめる",
  "      const d = [];",
  "      for (let i = 1; i < cs.length; i++) d.push(cs[i] - cs[i - 1]);",
  "      d.sort((a, b) => a - b);",
  "      const step = d.length ? d[Math.floor(d.length / 2)] : 0;",
  "      const full = [cs[0]];",
  "      for (let i = 1; i < cs.length; i++) {",
  "        const k = step ? Math.round((cs[i] - cs[i - 1]) / step) : 1;",
  "        for (let j = 1; j < k; j++) full.push(cs[i - 1] + (cs[i] - cs[i - 1]) * j / k);",
  "        full.push(cs[i]);",
  "      }",
  "      return full;",
  "    };",
  "    const rowsC = centers(spots.map((p) => p.y)), colsC = centers(spots.map((p) => p.x));",
  "    const near = (cs, v) => { let b = 0; cs.forEach((c, i) => { if (Math.abs(c - v) < Math.abs(cs[b] - v)) b = i; }); return b; };",
  "    const at = {};",
  "    spots.forEach((p) => { const k = near(rowsC, p.y) + ',' + near(colsC, p.x); (at[k] = at[k] || []).push(p.key); });",
  "    return { rows: rowsC.length, cols: colsC.length, at };",
  "  }",
  "  // </match>",
  "",
  "  // ───────── 保存 ─────────",
  "  $('save').onclick = async () => {",
  "    if (!D || busy) return; // 2回押したとき・読み込み中は書きこまない",
  "    busy = 'save';",
  "    $('save').disabled = true;",
  "    if (mode === 'seat') {",
  "      const me = E, target = D.target;",
  "      try {",
  "        const spots = Array.from(me.spots.entries()).map(([key, p]) => ({ key, x: p.x, y: p.y }));",
  "        const r = await run('saveSeat', { target, image: me.newImage, spots, desk: me.desk });",
  "        seats[target] = r;",
  "        busy = '';",
  "        if (E === me) {",
  "          stopCamera();",
  "          mode = 'take'; E = null;",
  "          note = '✓ 座席表を保存しました。座席表の名前を押すと欠席になります。';",
  "        }",
  "        draw();",
  "      } catch (e) {",
  "        busy = '';",
  "        alert('座席表を保存できませんでした：' + errText(e));",
  "        if (E === me && mode === 'seat') drawSeatEdit(); else if (D) draw();",
  "      }",
  "      return;",
  "    }",
  "    const d0 = D;",
  "    const n = absent.size;",
  "    try {",
  "      const r = await run('saveDay', { course: D.course, date, cols, double: !cols.length && double, absent: Array.from(absent) });",
  "      busy = '';",
  "      if (D !== d0) { draw(); return; }",
  "      const wasNew = !cols.length;",
  "      D = r;",
  "      note = '✓ 教務手帳に' + (wasNew ? '書きこみました' : '上書きしました') + '（' + dayLabel(r.date || date) + '・欠席 ' + n + '人' + (r.saved.length > 1 ? '・2コマ分' : '') + '）';",
  "      startEdit(r.saved[0]);",
  "    } catch (e) {",
  "      busy = '';",
  "      alert('書きこめませんでした：' + errText(e));",
  "      if (D) drawFoot();",
  "    }",
  "  };",
  "  $('cancelSeat').onclick = () => {",
  "    if (busy || !confirm('座席表の変更を保存せずにやめますか？')) return;",
  "    stopCamera();",
  "    mode = 'take'; E = null;",
  "    draw();",
  "  };",
  "",
  "  // 日付を変える（後日、休んだ日の分を入れるとき）。教務手帳には日付の順に列が入る",
  "  function setDate(k) {",
  "    if (!k || k === date || busy === 'save') { $('dateInput').value = date; return; }",
  "    if (mode === 'seat') { alert('座席表を保存するか「やめる」を押してから、日付を変えてください。'); $('dateInput').value = date; return; }",
  "    if (dirty && !confirm('書きこんでいない変更を捨てて、日付を変えますか？')) { $('dateInput').value = date; return; }",
  "    date = k;",
  "    dirty = false; note = '';",
  "    if (D) { drawTop(); pick(D.course); } else draw(); // 読めなかったときは pick がもとの日付にもどす",
  "  }",
  "  function shift(n) {",
  "    const p = date.split('-').map(Number);",
  "    const d = new Date(p[0], p[1] - 1, p[2] + n);",
  "    setDate(d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2));",
  "  }",
  "  $('dateInput').onchange = (e) => setDate(e.target.value);",
  "  $('prev').onclick = () => shift(-1);",
  "  $('next').onclick = () => shift(1);",
  "  $('today').onclick = () => setDate(I.today);",
  "  const leaving = (fn) => () => {",
  "    if (busy === 'save') return;",
  "    if ((dirty || mode === 'seat') && !confirm('保存していない変更は消えます。よろしいですか？')) return;",
  "    if (E && E.camera) { stopCamera(); if (mode === 'seat') drawSeatEdit(); }",
  "    fn();",
  "  };",
  "  const openOther = (name, ...args) => run(name, ...args).catch((e) => alert(errText(e)));",
  "  $('lPeriod').onclick = leaving(() => openOther('openPeriod'));",
  "  $('lBulk').onclick = leaving(() => openOther('openBulk'));",
  "  $('lSetup').onclick = leaving(() => openOther('openSetup', ''));",
  "  $('lReload').onclick = leaving(() => {",
  "    run('getTakeInit').then((d) => {",
  "      I = d; const name = D && D.course; D = null; mode = 'take'; E = null; dirty = false; note = '';",
  "      pickSeq++; busy = ''; // 読み込み中の講座は使わない",
  "      Object.keys(seats).forEach((k) => delete seats[k]);",
  "      draw();",
  "      if (name && I.courses.some((c) => c.name === name)) pick(name);",
  "    }).catch((e) => alert('読み込めませんでした：' + errText(e)));",
  "  });",
  "",
  "  run('getTakeInit').then((d) => {",
  "    I = d; date = d.today;",
  "    draw();",
  "    if (I.courses.length === 1) pick(I.courses[0].name);",
  "  }).catch((e) => { $('left').innerHTML = ''; $('left').append(el('div', { class: 'msg warn' }, '読み込めませんでした：' + errText(e))); });",
  "</script>",
  "</body>",
  "</html>",
].join('\n');
HTML_FILES["Setup"] = [
  "<!DOCTYPE html>",
  "<html>",
  "<head>",
  "<base target=\"_top\">",
  "<meta charset=\"utf-8\">",
  "<style>",
  "  body { font-family: \"Noto Sans JP\", sans-serif; margin: 0; font-size: 13px; color: #202124; }",
  "  .tabs { display: flex; gap: 4px; border-bottom: 2px solid #dadce0; padding: 8px 16px 0; position: sticky; top: 0; background: #fff; z-index: 2; }",
  "  .tab { padding: 8px 16px; border: 1px solid #dadce0; border-bottom: none; border-radius: 8px 8px 0 0; background: #f8f9fa; cursor: pointer; font-size: 14px; }",
  "  .tab.on { background: #fff; color: #1a73e8; font-weight: bold; border-color: #1a73e8; position: relative; top: 2px; }",
  "  .tab .ck { color: #188038; margin-left: 4px; }",
  "  .pane { padding: 12px 16px 20px; }",
  "  .lead { margin: 0 0 8px; line-height: 1.6; }",
  "  .lead b { color: #1967d2; }",
  "  textarea { width: 100%; box-sizing: border-box; height: 150px; font-size: 13px; padding: 8px; border: 2px dashed #1a73e8; border-radius: 8px; font-family: inherit; }",
  "  .ex { color: #5f6368; font-size: 12px; margin: 4px 0 10px; line-height: 1.6; }",
  "  .bar { display: flex; gap: 10px; align-items: center; margin: 8px 0; flex-wrap: wrap; }",
  "  button { font-family: inherit; font-size: 14px; padding: 7px 16px; border-radius: 8px; border: 1px solid #1a73e8; background: #fff; color: #1a73e8; cursor: pointer; }",
  "  button.primary { background: #1a73e8; color: #fff; }",
  "  button.danger { border-color: #d93025; color: #d93025; }",
  "  button:disabled { opacity: .5; cursor: default; }",
  "  .count span { display: inline-block; margin-right: 8px; padding: 2px 8px; border-radius: 10px; font-size: 12px; }",
  "  .new { background: #e6f4ea; color: #137333; }",
  "  .update { background: #fef7e0; color: #b06000; }",
  "  .same { background: #f1f3f4; color: #5f6368; }",
  "  .bad { background: #fce8e6; color: #c5221f; }",
  "  .tbl { max-height: 300px; overflow: auto; border: 1px solid #dadce0; border-radius: 8px; }",
  "  table { border-collapse: collapse; width: 100%; }",
  "  th, td { padding: 5px 8px; border-bottom: 1px solid #eee; text-align: left; font-size: 13px; }",
  "  th { position: sticky; top: 0; background: #f8f9fa; }",
  "  .ok { background: #e6f4ea; color: #137333; padding: 10px; border-radius: 8px; margin: 8px 0; }",
  "  .warn { background: #fef7e0; color: #b06000; padding: 10px; border-radius: 8px; margin: 8px 0; }",
  "  .cols { display: grid; grid-template-columns: 1fr 1.2fr; gap: 16px; }",
  "  .list .item { display: flex; align-items: center; gap: 8px; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 8px; margin-bottom: 6px; cursor: pointer; }",
  "  .list .item:hover { background: #f1f3f4; }",
  "  .list .item.on { border-color: #1a73e8; background: #e8f0fe; }",
  "  .list .item b { flex: 1; }",
  "  .list .item small { color: #5f6368; }",
  "  .form { border: 1px solid #dadce0; border-radius: 10px; padding: 12px 14px; }",
  "  .form h3 { margin: 0 0 6px; font-size: 15px; }",
  "  label.f { display: block; margin: 10px 0 4px; font-size: 12px; color: #5f6368; }",
  "  input[type=text], input[type=number], input[type=date] { font-family: inherit; font-size: 14px; padding: 6px 8px; border: 1px solid #dadce0; border-radius: 6px; box-sizing: border-box; }",
  "  input.wide { width: 100%; }",
  "  .chips { display: flex; flex-wrap: wrap; gap: 6px; }",
  "  .chip { padding: 5px 12px; border: 1px solid #dadce0; border-radius: 14px; background: #fff; font-size: 13px; color: #202124; }",
  "  .chip.on { background: #e8f0fe; border-color: #1a73e8; color: #1967d2; font-weight: bold; }",
  "  .units { display: flex; gap: 6px; align-items: center; flex-wrap: wrap; }",
  "  .units .chip { min-width: 40px; }",
  "  .note { font-size: 12px; color: #5f6368; margin-top: 4px; }",
  "  .bulk td input { width: 70px; text-align: center; }",
  "  .bulk td.n { color: #5f6368; }",
  "  .empty { color: #5f6368; padding: 8px 0; }",
  "</style>",
  "</head>",
  "<body data-tab=\"<?= tab ?>\">",
  "<div class=\"tabs\" id=\"tabs\"></div>",
  "<div class=\"pane\" id=\"pane\"><div class=\"empty\">読み込み中…</div></div>",
  "",
  "<script>",
  "  // 最初に開くタブ（openSetup('bulk') など）。JS の文字列の中ではなく HTML の属性で受け取る（エスケープのされ方がはっきりしているため）",
  "  const START_TAB = document.body.dataset.tab || '';",
  "  const $ = (id) => document.getElementById(id);",
  "  const run = (name, ...args) => new Promise((ok, ng) => google.script.run.withSuccessHandler(ok).withFailureHandler(ng)[name](...args));",
  "  const errText = (e) => (e && e.message) || String(e);",
  "  let S = null;",
  "  let tab = '';",
  "",
  "  function el(tag, attrs, ...kids) {",
  "    const e = document.createElement(tag);",
  "    Object.entries(attrs || {}).forEach(([k, v]) => {",
  "      if (v == null || v === false) return;",
  "      if (k.startsWith('on')) e.addEventListener(k.slice(2), v);",
  "      else if (k === 'class') e.className = v;",
  "      else if (k === 'style') e.style.cssText = v;",
  "      else if (k in e) e[k] = v;",
  "      else e.setAttribute(k, v);",
  "    });",
  "    kids.flat().forEach((c) => { if (c != null && c !== false) e.append(c.nodeType ? c : String(c)); });",
  "    return e;",
  "  }",
  "",
  "  const TABS = [",
  "    ['roster', '① 名簿', () => S.students > 0],",
  "    ['course', '② 講座・単位数', () => S.courses.length > 0],",
  "    ['bulk', '③ まとめて入力', null],",
  "    ['settings', '設定', null],",
  "  ];",
  "  function drawTabs() {",
  "    $('tabs').innerHTML = '';",
  "    TABS.forEach(([id, label, done]) => $('tabs').append(el('div', { class: 'tab' + (tab === id ? ' on' : ''), onclick: () => { tab = id; draw(); } },",
  "      label, done && done() ? el('span', { class: 'ck' }, '✓') : null)));",
  "  }",
  "  function draw() {",
  "    drawTabs();",
  "    const p = $('pane');",
  "    p.innerHTML = '';",
  "    if (!TABS.some((t) => t[0] === tab)) tab = 'roster';",
  "    ({ roster: rosterPane, course: coursePane, bulk: bulkPane, settings: settingsPane })[tab](p);",
  "  }",
  "",
  "  // ───────── ① 名簿 ─────────",
  "  const isKana = (c) => /^[ぁ-んァ-ヶー・\\s]+$/.test(c);",
  "  const hasKanji = (c) => /[一-鿿々]/.test(c);",
  "  // 1行から 学籍番号・氏名・ふりがな を見つける（課題提出バーコード管理と同じ読み取り方）",
  "  function parseLine(line) {",
  "    // 「1−1」「1ー1」のような数字の間の横線は「-」にそろえる（NFKC では「−」「ー」は変わらないため）",
  "    const n = line.normalize('NFKC').replace(/　/g, ' ').replace(/(\\d)\\s*[‐‑‒–—―−ー]\\s*(?=\\d)/g, '$1-').trim();",
  "    if (!n) return null;",
  "    let cells;",
  "    if (n.includes('\\t')) cells = n.split('\\t');",
  "    else if (n.includes(',')) cells = n.split(',');",
  "    else cells = n.split(/\\s+/);",
  "    const byTab = n.includes('\\t') || n.includes(',');",
  "    cells = cells.map((c) => c.trim().replace(/^\"(.*)\"$/, '$1').trim()).filter(Boolean); // CSV の「\"」を外す",
  "    if (!/\\d/.test(n) && /氏名|名前|学籍|番号|ふりがな|フリガナ|クラス/.test(n)) return { header: true };",
  "    let gk = '', grade = '', cls = '', no = '';",
  "    const small = [], text = [];",
  "    cells.forEach((c) => {",
  "      let m;",
  "      if (!gk && /^\\d{4}$/.test(c)) gk = c;",
  "      else if ((m = c.match(/^(\\d)\\s*(?:-|年)\\s*(\\d{1,2})\\s*(?:組|-)\\s*(\\d{1,2})\\s*番?$/))) { grade = m[1]; cls = m[2]; small.unshift(m[3]); } // 「1年1組1番」「1-1-1」が1つのマス",
  "      else if ((m = c.match(/^(\\d)\\s*(?:-|年)\\s*(\\d{1,2})\\s*組?$/))) { grade = m[1]; cls = m[2]; }",
  "      else if ((m = c.match(/^(\\d)\\s*年$/))) grade = m[1];",
  "      else if ((m = c.match(/^(\\d{1,2})\\s*組$/))) cls = m[1];",
  "      else if ((m = c.match(/^(\\d{1,2})\\s*番?$/))) { if (/番$/.test(c)) small.unshift(m[1]); else small.push(m[1]); } // 「番」が付いた数を番号に",
  "      else if (/^[\\d\\-\\/]+$/.test(c)) { /* 生徒ID・電話番号などの数字は使わない */ }",
  "      else if (/^(男|女|男子|女子)$/.test(c)) { /* 性別の列は使わない */ }",
  "      else text.push(c.replace(/\\s+/g, ' '));",
  "    });",
  "    if (!gk) {",
  "      if (grade && cls && small.length) no = small[0];",
  "      else if (small.length >= 3) { grade = small[0]; cls = small[1]; no = small[2]; }",
  "      if (grade && cls && no && Number(grade) <= 9 && Number(cls) <= 9) gk = grade + cls + String(no).padStart(2, '0');",
  "    }",
  "    let kana = '';",
  "    if (byTab && text.length >= 2) {",
  "      const L = text.length - 1, rest = text.slice(0, L);",
  "      let i = text.findIndex((c, k) => isKana(c) && (c.includes(' ') || c.length >= 5) && text.some((o, j) => j !== k && hasKanji(o)));",
  "      // 「原 舞｜はらまい」「原｜舞｜はらまい」：漢字の氏名のあとの短いかなもふりがな（「青木｜さくら」は名なのでそのまま）",
  "      // 「ユーリ ノーマン｜ゆーり のーまん」：カタカナの氏名のあとの、ひらがなの氏名",
  "      if (i < 0 && isKana(text[L]) && (rest.length >= 2 ? rest.every(hasKanji) : rest[0].includes(' ') && (hasKanji(rest[0]) || (text[L].includes(' ') && /^[ぁ-ん\\s・ー]+$/.test(text[L]) && !/^[ぁ-ん\\s・ー]+$/.test(rest[0]))))) i = L;",
  "      if (i >= 0) kana = text.splice(i, 1)[0];",
  "      // 「林｜誠｜はやし｜まこと」：ふりがなが姓・名の2列",
  "      else if (text.length >= 3 && isKana(text[L]) && isKana(text[L - 1]) && text.slice(0, -2).some(hasKanji)) kana = text.splice(-2).join(' ');",
  "    } else if (!byTab && text.length >= 3 && isKana(text[text.length - 1]) && isKana(text[text.length - 2]) && text.slice(0, -2).some(hasKanji)) {",
  "      kana = text.splice(-2).join(' ');",
  "    }",
  "    const name = text.join(' ').trim();",
  "    if (!name || !gk) return { bad: line };",
  "    return { gakuseki: gk, name, kana };",
  "  }",
  "",
  "  function rosterPane(p) {",
  "    let parsed = [], bad = [];",
  "    const ta = el('textarea', { placeholder: 'ここに貼り付け（Ctrl + V）' });",
  "    const count = el('span', { class: 'count' });",
  "    const wrap = el('div', { class: 'tbl', style: 'display:none' });",
  "    const done = el('div');",
  "    const save = el('button', { class: 'primary', disabled: true }, '② この内容で名簿に登録する');",
  "    const check = el('button', { class: 'primary' }, '① 読み取って確認する');",
  "    const nowCount = el('b', null, S.students + '人');",
  "    p.append(",
  "      el('p', { class: 'lead' }, 'Excel・校務システム・名簿のファイルから、生徒の行を', el('b', null, 'まとめてコピーして、下の枠に貼り付け'), 'てください。列の順番はそのままでかまいません。', el('br'),",
  "        '登録ずみ：', nowCount, S.classes.length ? '（' + S.classes.join('・') + '）' : ''),",
  "      ta,",
  "      el('div', { class: 'ex' }, '読み取れる形の例：「1101　青木 さくら」／「1　1　1　青木　さくら」（学年・組・番号が別の列）／「1年1組　1番　青木さくら　あおき さくら」', el('br'),",
  "        '学籍番号（1101＝1年1組1番）から、クラスと番号は自動で入ります。転出した生徒は、シート「名簿」の「除外」に ✓ を付けてください。'),",
  "      el('div', { class: 'bar' }, check, count), wrap, el('div', { class: 'bar' }, save), done);",
  "    ta.addEventListener('input', () => { save.disabled = true; });",
  "    check.onclick = async () => {",
  "      parsed = []; bad = [];",
  "      const seen = {};",
  "      ta.value.split('\\n').forEach((l) => {",
  "        const r = parseLine(l);",
  "        if (!r || r.header) return;",
  "        if (r.bad) bad.push(r.bad);",
  "        else if (seen[r.gakuseki]) bad.push(l.trim() + '（学籍番号 ' + r.gakuseki + ' が「' + seen[r.gakuseki] + '」と重なっています）');",
  "        else { seen[r.gakuseki] = r.name; parsed.push(r); }",
  "      });",
  "      if (!parsed.length) { alert('生徒の行が見つかりませんでした。学籍番号（1101 など）か「学年・組・番号」と、氏名が入った表を貼り付けてください。'); return; }",
  "      check.disabled = true;",
  "      let plan;",
  "      try { plan = await run('previewRoster', parsed); } catch (e) { alert(errText(e)); check.disabled = false; return; }",
  "      check.disabled = false;",
  "      const cnt = { new: 0, update: 0, same: 0 };",
  "      const body = el('tbody');",
  "      plan.forEach((r) => {",
  "        cnt[r.kind]++;",
  "        body.append(el('tr', null, el('td', null, r.gakuseki), el('td', null, r.name), el('td', null, r.kana),",
  "          el('td', { class: r.kind }, { new: '新しく追加', update: '更新（前：' + r.before + '）', same: '登録ずみ' }[r.kind])));",
  "      });",
  "      bad.forEach((l) => body.append(el('tr', null, el('td', { class: 'bad', colSpan: 4 }, '読み取れない行（登録しません）：' + l))));",
  "      wrap.innerHTML = '';",
  "      wrap.append(el('table', null, el('thead', null, el('tr', null, ['学籍番号', '氏名', 'ふりがな', '登録のされ方'].map((t) => el('th', null, t)))), body));",
  "      wrap.style.display = '';",
  "      count.innerHTML = '';",
  "      [['new', '追加 ' + cnt.new + '人'], ['update', '更新 ' + cnt.update + '人'], ['same', '変更なし ' + cnt.same + '人']].forEach(([k, t]) => count.append(el('span', { class: k }, t)));",
  "      save.disabled = !(cnt.new + cnt.update);",
  "    };",
  "    save.onclick = async () => {",
  "      save.disabled = true;",
  "      try {",
  "        S = await run('importRoster', parsed);",
  "        const r = S.result;",
  "        drawTabs();",
  "        nowCount.textContent = S.students + '人';",
  "        done.innerHTML = '';",
  "        done.append(el('div', { class: 'ok' }, '✓ 名簿に登録しました（追加 ' + r.added + '人・更新 ' + r.updated + '人）。',",
  "          el('button', { class: 'primary', style: 'margin-left:10px', onclick: () => { tab = 'course'; draw(); } }, '次へ：講座を登録する')));",
  "      } catch (e) { alert(errText(e)); save.disabled = false; }",
  "    };",
  "  }",
  "",
  "  // ───────── ② 講座・単位数 ─────────",
  "  let editing = null; // 直している講座の名前（null＝新しく）",
  "  function coursePane(p) {",
  "    const list = el('div', { class: 'list' });",
  "    if (!S.courses.length) list.append(el('div', { class: 'empty' }, 'まだありません。右の欄から登録してください。'));",
  "    S.courses.forEach((c) => list.append(el('div', { class: 'item' + (editing === c.name ? ' on' : ''), onclick: () => { editing = c.name; draw(); } },",
  "      el('b', null, c.name), el('small', null, c.units + '単位 ・ ' + c.target + ' ・ ' + c.members + '人'))));",
  "    list.append(el('button', { onclick: () => { editing = null; draw(); } }, '＋ 新しい講座'));",
  "",
  "    const c = editing ? S.courses.find((x) => x.name === editing) : null;",
  "    const name = el('input', { type: 'text', class: 'wide', value: c ? c.name : '', placeholder: '例：1-1 英語コミュⅠ' });",
  "    const chosen = new Set();",
  "    const extra = [];",
  "    (c ? c.target.split(/\\s*,\\s*/) : []).forEach((t) => { if (S.classes.includes(t)) chosen.add(t); else if (t) extra.push(t); });",
  "    const chips = el('div', { class: 'chips' });",
  "    function drawChips() {",
  "      chips.innerHTML = '';",
  "      if (!S.classes.length) chips.append(el('div', { class: 'warn' }, '先に「① 名簿」を登録してください。'));",
  "      S.classes.forEach((k) => chips.append(el('button', { class: 'chip' + (chosen.has(k) ? ' on' : ''), onclick: () => { chosen.has(k) ? chosen.delete(k) : chosen.add(k); drawChips(); } }, k)));",
  "    }",
  "    drawChips();",
  "    const more = el('input', { type: 'text', class: 'wide', value: extra.join(', '), placeholder: '例：1105, 1203, 2111（選択授業などで、クラスの一部だけのとき）' });",
  "    let units = c ? c.units : '';",
  "    const unitInput = el('input', { type: 'number', min: 1, max: 20, value: units, style: 'width:70px', oninput: (e) => { units = e.target.value; drawUnits(); } });",
  "    const unitBox = el('div', { class: 'units' });",
  "    const hoursNote = el('div', { class: 'note' });",
  "    function drawUnits() {",
  "      unitBox.innerHTML = '';",
  "      [1, 2, 3, 4, 5, 6].forEach((u) => unitBox.append(el('button', { class: 'chip' + (Number(units) === u ? ' on' : ''), onclick: () => { units = u; unitInput.value = u; drawUnits(); } }, u)));",
  "      unitBox.append(unitInput, '単位');",
  "      hoursNote.textContent = Number(units) > 0 ? '年間時数 ' + Number(units) * S.hoursPerUnit + '時間（' + units + '単位 × ' + S.hoursPerUnit + '）→ 20% は ' + Math.round(Number(units) * S.hoursPerUnit * 0.2 * 10) / 10 + '時間' : '';",
  "    }",
  "    drawUnits();",
  "    const msg = el('div');",
  "    const save = el('button', { class: 'primary', onclick: async () => {",
  "      const target = Array.from(chosen).concat(more.value.split(/[,、\\s]+/).filter(Boolean)).join(', ');",
  "      save.disabled = true;",
  "      try {",
  "        S = await run('saveCourse', { oldName: c ? c.name : '', name: name.value, target, units, memo: c ? c.memo : '' });",
  "        editing = name.value.trim();",
  "        if (c && bulkCourse === c.name) bulkCourse = editing; // 名前を変えたら「まとめて入力」の選択もついていく",
  "        draw();",
  "        $('pane').prepend(el('div', { class: 'ok' }, '✓ 「' + editing + '」を保存しました。教務手帳のシート「' + editing + '」ができています。'));",
  "      } catch (e) { alert(errText(e)); save.disabled = false; }",
  "    } }, c ? '保存する' : 'この講座を登録する');",
  "    const del = c ? el('button', { class: 'danger', onclick: async () => {",
  "      if (!confirm('「' + c.name + '」を講座の一覧から外しますか？\\n（教務手帳のシートは消さずに残します。いらなければシートを削除してください）')) return;",
  "      try { S = await run('deleteCourse', c.name); editing = null; draw(); } catch (e) { alert(errText(e)); }",
  "    } }, '一覧から外す') : null;",
  "",
  "    const form = el('div', { class: 'form' },",
  "      el('h3', null, c ? '「' + c.name + '」を直す' : '新しい講座を登録する'),",
  "      el('label', { class: 'f' }, '講座名（教務手帳のシートの名前になります）'), name,",
  "      el('label', { class: 'f' }, '対象のクラス（押して選ぶ・合同授業なら複数）'), chips,",
  "      el('label', { class: 'f' }, '個別の生徒（学籍番号を「,」区切りで・なくてもよい）'), more,",
  "      el('label', { class: 'f' }, '単位数'), unitBox, hoursNote,",
  "      el('div', { class: 'bar', style: 'margin-top:14px' }, save, del), msg);",
  "    p.append(el('p', { class: 'lead' }, '持っている授業を1つずつ登録します。登録すると、講座ごとに', el('b', null, '教務手帳のシート'), 'ができます。'),",
  "      el('div', { class: 'cols' }, list, form));",
  "    if (S.courses.length) p.append(el('div', { class: 'bar' }, el('button', { onclick: () => { tab = 'bulk'; draw(); } }, '次へ：まとめて入力（学期の途中から使うとき）')));",
  "  }",
  "",
  "  // ───────── ③ まとめて入力 ─────────",
  "  let bulkCourse = '';",
  "  function bulkPane(p) {",
  "    p.append(el('p', { class: 'lead' }, '学期の途中から使い始めるときは、', el('b', null, 'それまでの欠課時数'), 'を生徒ごとにまとめて入れてください。',",
  "      '教務手帳の「それ以前の欠課」の列に入り、欠課時数と割合に足されます。'));",
  "    if (!S.courses.length) { p.append(el('div', { class: 'warn' }, '先に「② 講座・単位数」を登録してください。')); return; }",
  "    if (!bulkCourse || !S.courses.some((c) => c.name === bulkCourse)) bulkCourse = S.courses[0].name;",
  "    p.append(el('div', { class: 'chips' }, S.courses.map((c) => el('button', { class: 'chip' + (c.name === bulkCourse ? ' on' : ''), onclick: () => { bulkCourse = c.name; draw(); } }, c.name))));",
  "    const box = el('div', null, el('div', { class: 'empty' }, '読み込み中…'));",
  "    p.append(box);",
  "    run('getBulk', bulkCourse).then((b) => {",
  "      box.innerHTML = '';",
  "      const until = el('input', { type: 'date', value: b.until || '' });",
  "      const inputs = []; // [学籍番号, 入力欄]（表の順。{} だと数字のキーが数の順に並びかわるので配列）",
  "      const body = el('tbody');",
  "      b.students.forEach((s, i) => {",
  "        const inp = el('input', { type: 'number', min: 0, value: s.before, placeholder: '0' });",
  "        inp.addEventListener('keydown', (e) => {",
  "          if (e.key === 'Enter' || e.key === 'ArrowDown') { e.preventDefault(); const n = (inputs[i + 1] || [])[1]; if (n) { n.focus(); n.select(); } }",
  "          if (e.key === 'ArrowUp') { e.preventDefault(); const n = (inputs[i - 1] || [])[1]; if (n) { n.focus(); n.select(); } }",
  "        });",
  "        inputs.push([s.key, inp]);",
  "        body.append(el('tr', null, el('td', { class: 'n' }, s.cls), el('td', { class: 'n' }, s.no), el('td', null, s.key), el('td', null, s.name), el('td', null, inp)));",
  "      });",
  "      const msg = el('div');",
  "      const save = el('button', { class: 'primary', onclick: async () => {",
  "        const values = {};",
  "        inputs.forEach(([k, inp]) => { values[k] = inp.value; });",
  "        save.disabled = true;",
  "        try {",
  "          await run('saveBulk', { course: b.course, until: until.value, values });",
  "          msg.innerHTML = '';",
  "          msg.append(el('div', { class: 'ok' }, '✓ 「' + b.course + '」の教務手帳に入れました。'));",
  "        } catch (e) { alert(errText(e)); }",
  "        save.disabled = false;",
  "      } }, 'この講座の分を保存する');",
  "      box.append(",
  "        el('div', { class: 'bar' }, el('span', null, 'いつまでの分か（なくてもよい）：'), until),",
  "        el('div', { class: 'tbl bulk', style: 'max-height:calc(100vh - 300px);min-height:160px' }, el('table', null,",
  "          el('thead', null, el('tr', null, ['クラス', '番号', '学籍番号', '氏名', 'それまでの欠課時数'].map((t) => el('th', null, t)))), body)),",
  "        el('div', { class: 'note' }, '欠課のない生徒は空欄のままでOK。Enter で次の生徒に進みます。'),",
  "        el('div', { class: 'bar' }, save), msg);",
  "      if (inputs[0]) inputs[0][1].focus();",
  "    }).catch((e) => { box.innerHTML = ''; box.append(el('div', { class: 'warn' }, errText(e))); });",
  "  }",
  "",
  "  // ───────── 設定 ─────────",
  "  function settingsPane(p) {",
  "    const hours = el('input', { type: 'number', min: 1, value: S.hoursPerUnit, style: 'width:80px' });",
  "    const lines = el('input', { type: 'text', value: S.lines, style: 'width:240px' });",
  "    const terms = el('input', { type: 'text', class: 'wide', value: S.terms });",
  "    const save = el('button', { class: 'primary', onclick: async () => {",
  "      save.disabled = true;",
  "      try {",
  "        S = await run('saveSettings', { hoursPerUnit: hours.value, lines: lines.value, terms: terms.value });",
  "        draw();",
  "        $('pane').prepend(el('div', { class: 'ok' }, '✓ 保存しました。教務手帳の色・年間時数も新しい設定にしました。'));",
  "      } catch (e) { alert(errText(e)); save.disabled = false; }",
  "    } }, '保存する');",
  "    p.append(",
  "      el('label', { class: 'f' }, '1単位あたりの年間授業時数（年間時数 ＝ 単位数 × この数）'), el('div', null, hours, ' 時間'),",
  "      el('div', { class: 'note' }, '講座ごとに年間時数を変えたいときは、教務手帳のシートの G1（年間時数）に数を直接入れてください。'),",
  "      el('label', { class: 'f' }, '色を変える割合（%・「,」区切り）'), lines,",
  "      el('div', { class: 'note' }, '低い順に 黄・橙・赤・濃い赤。例：20, 25, 30, 50'),",
  "      el('label', { class: 'f' }, '学期の区切り（期間集計のボタンになります）'), terms,",
  "      el('div', { class: 'note' }, '例：1学期 4/1-7/31, 2学期 8/1-12/31, 3学期 1/1-3/31（前期・後期でも可）'),",
  "      el('div', { class: 'bar', style: 'margin-top:14px' }, save));",
  "  }",
  "",
  "  run('getSetupData').then((d) => {",
  "    S = d;",
  "    tab = TABS.some((t) => t[0] === START_TAB) ? START_TAB : !S.students ? 'roster' : 'course';",
  "    draw();",
  "  }).catch((e) => { $('pane').innerHTML = ''; $('pane').append(el('div', { class: 'warn' }, '読み込めませんでした：' + errText(e))); });",
  "</script>",
  "</body>",
  "</html>",
].join('\n');
HTML_FILES["Period"] = [
  "<!DOCTYPE html>",
  "<html>",
  "<head>",
  "<base target=\"_top\">",
  "<meta charset=\"utf-8\">",
  "<style>",
  "  body { font-family: \"Noto Sans JP\", sans-serif; margin: 0; padding: 12px 16px; font-size: 13px; color: #202124; }",
  "  h2 { font-size: 12px; color: #5f6368; margin: 12px 0 6px; }",
  "  .chips { display: flex; flex-wrap: wrap; gap: 6px; }",
  "  .chip { padding: 5px 12px; border: 1px solid #dadce0; border-radius: 14px; background: #fff; font-size: 13px; cursor: pointer; font-family: inherit; }",
  "  .chip.on { background: #e8f0fe; border-color: #1a73e8; color: #1967d2; font-weight: bold; }",
  "  .range { display: flex; gap: 8px; align-items: center; margin-top: 8px; }",
  "  .range input { font-family: inherit; font-size: 14px; padding: 5px 8px; border: 1px solid #dadce0; border-radius: 6px; }",
  "  .tbl { max-height: calc(100vh - 290px); min-height: 160px; overflow: auto; border: 1px solid #dadce0; border-radius: 8px; margin-top: 8px; }",
  "  table { border-collapse: collapse; width: 100%; }",
  "  th, td { padding: 5px 8px; border-bottom: 1px solid #eee; font-size: 13px; text-align: right; white-space: nowrap; }",
  "  th { position: sticky; top: 0; background: #f8f9fa; }",
  "  th.l, td.l { text-align: left; }",
  "  td.z { color: #bdc1c6; }",
  "  td.hit { color: #c5221f; font-weight: bold; }",
  "  .note { font-size: 12px; color: #5f6368; margin-top: 6px; }",
  "  .bar { display: flex; gap: 10px; align-items: center; margin-top: 10px; }",
  "  button.btn { font-family: inherit; font-size: 14px; padding: 7px 16px; border-radius: 8px; border: 1px solid #1a73e8; background: #fff; color: #1a73e8; cursor: pointer; }",
  "  .ok { background: #e6f4ea; color: #137333; padding: 8px; border-radius: 8px; }",
  "  .warn { background: #fef7e0; color: #b06000; padding: 8px; border-radius: 8px; margin-top: 8px; }",
  "  label { font-size: 13px; }",
  "</style>",
  "</head>",
  "<body>",
  "<h2>講座</h2>",
  "<div class=\"chips\" id=\"courses\"></div>",
  "<h2>期間</h2>",
  "<div class=\"chips\" id=\"ranges\"></div>",
  "<div class=\"range\"><input type=\"date\" id=\"from\"> 〜 <input type=\"date\" id=\"to\"></div>",
  "<div id=\"out\"></div>",
  "",
  "<script>",
  "  const $ = (id) => document.getElementById(id);",
  "  const run = (name, ...args) => new Promise((ok, ng) => google.script.run.withSuccessHandler(ok).withFailureHandler(ng)[name](...args));",
  "  const errText = (e) => (e && e.message) || String(e);",
  "  let P = null;",
  "  const st = { course: '', range: '', from: '', to: '' };",
  "  let onlyAbsent = false;",
  "  let last = null;",
  "  let seq = 0; // 集計を続けて押したとき、前の結果があとから届いても使わないため",
  "",
  "  function el(tag, attrs, ...kids) {",
  "    const e = document.createElement(tag);",
  "    Object.entries(attrs || {}).forEach(([k, v]) => {",
  "      if (v == null || v === false) return;",
  "      if (k.startsWith('on')) e.addEventListener(k.slice(2), v);",
  "      else if (k === 'class') e.className = v;",
  "      else if (k in e) e[k] = v;",
  "      else e.setAttribute(k, v);",
  "    });",
  "    kids.flat().forEach((c) => { if (c != null && c !== false) e.append(c.nodeType ? c : String(c)); });",
  "    return e;",
  "  }",
  "  const pad = (n) => ('0' + n).slice(-2);",
  "  const key = (d) => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());",
  "",
  "  function ranges() {",
  "    const t = new Date(P.today.replace(/-/g, '/'));",
  "    const m0 = new Date(t.getFullYear(), t.getMonth(), 1), m1 = new Date(t.getFullYear(), t.getMonth() + 1, 0);",
  "    const p0 = new Date(t.getFullYear(), t.getMonth() - 1, 1), p1 = new Date(t.getFullYear(), t.getMonth(), 0);",
  "    const w0 = new Date(t); w0.setDate(t.getDate() - ((t.getDay() + 6) % 7));",
  "    return [{ name: '今週', from: key(w0), to: P.today }, { name: '今月', from: key(m0), to: key(m1) }, { name: '先月', from: key(p0), to: key(p1) }]",
  "      .concat(P.terms).concat([{ name: P.fy + '年度', from: P.fy + '-04-01', to: (P.fy + 1) + '-03-31' }]);",
  "  }",
  "",
  "  function drawChips() {",
  "    $('courses').innerHTML = '';",
  "    [['', 'すべての講座']].concat(P.courses.map((c) => [c, c])).forEach(([v, label]) => $('courses').append(",
  "      el('button', { class: 'chip' + (st.course === v ? ' on' : ''), onclick: () => { st.course = v; drawChips(); load(); } }, label)));",
  "    $('ranges').innerHTML = '';",
  "    ranges().forEach((r) => $('ranges').append(el('button', { class: 'chip' + (st.range === r.name ? ' on' : ''), onclick: () => {",
  "      Object.assign(st, { range: r.name, from: r.from, to: r.to }); $('from').value = r.from; $('to').value = r.to; drawChips(); load();",
  "    } }, r.name)));",
  "  }",
  "",
  "  async function load() {",
  "    const out = $('out');",
  "    const my = ++seq;",
  "    out.innerHTML = '';",
  "    if (st.from && st.to && st.from > st.to) { out.append(el('div', { class: 'warn' }, '期間の始めが終わりよりあとになっています。日付を直してください。')); return; }",
  "    out.append(el('div', { class: 'note' }, '集計しています…'));",
  "    let r;",
  "    try { r = await run('runPeriod', { course: st.course, from: st.from, to: st.to }); }",
  "    catch (e) { if (my !== seq) return; out.innerHTML = ''; out.append(el('div', { class: 'warn' }, errText(e))); return; }",
  "    if (my !== seq) return;",
  "    last = r;",
  "    draw();",
  "  }",
  "",
  "  function draw() {",
  "    const r = last;",
  "    const out = $('out');",
  "    out.innerHTML = '';",
  "    const multi = r.courses.length > 1;",
  "    const head = el('tr', null, el('th', { class: 'l' }, 'クラス'), el('th', { class: 'l' }, '番号'), el('th', { class: 'l' }, '氏名'),",
  "      r.courses.map((c) => el('th', null, c.name, el('br'), el('small', null, c.lessons + 'コマ中'))), multi ? el('th', null, '合計') : null);",
  "    const body = el('tbody');",
  "    let shown = 0;",
  "    r.rows.forEach((s) => {",
  "      let sum = 0;",
  "      const cells = r.courses.map((c) => {",
  "        const n = c.counts[s.key];",
  "        sum += n || 0;",
  "        return el('td', { class: n === undefined ? 'z' : n ? 'hit' : 'z' }, n === undefined ? '－' : n);",
  "      });",
  "      if (onlyAbsent && !sum) return;",
  "      shown++;",
  "      body.append(el('tr', null, el('td', { class: 'l' }, s.cls), el('td', { class: 'l' }, s.no), el('td', { class: 'l' }, s.name), cells,",
  "        multi ? el('td', { class: sum ? 'hit' : 'z' }, sum) : null));",
  "    });",
  "    const label = (st.from || st.to) ? st.from.replace(/-/g, '/') + ' 〜 ' + st.to.replace(/-/g, '/') : '全期間';",
  "    out.append(el('div', { class: 'bar' }, el('b', null, label),",
  "      el('label', null, el('input', { type: 'checkbox', checked: onlyAbsent, onchange: (e) => { onlyAbsent = e.target.checked; draw(); } }), ' 休んだ生徒だけ出す')));",
  "    out.append(el('div', { class: 'tbl' }, el('table', null, el('thead', null, head), body)));",
  "    if (!shown) out.append(el('div', { class: 'note' }, 'この期間に休んだ生徒はいません。'));",
  "    out.append(el('div', { class: 'note' }, '教務手帳の日付の列のうち、期間に入るものの「欠」を数えています。「まとめて入力」した「それ以前の欠課」は日付がないので入りません。'));",
  "    const done = el('span');",
  "    out.append(el('div', { class: 'bar' }, el('button', { class: 'btn', onclick: async (e) => {",
  "      e.target.disabled = true;",
  "      try { await run('exportPeriod', { course: st.course, from: st.from, to: st.to }); done.innerHTML = ''; done.append(el('span', { class: 'ok' }, '✓ 「期間集計」シートに書き出しました')); }",
  "      catch (err) { alert(errText(err)); }",
  "      e.target.disabled = false;",
  "    } }, '📄 シートに書き出す（印刷用）'), done));",
  "  }",
  "",
  "  $('from').onchange = (e) => { st.from = e.target.value; st.range = ''; drawChips(); load(); };",
  "  $('to').onchange = (e) => { st.to = e.target.value; st.range = ''; drawChips(); load(); };",
  "",
  "  run('getPeriodInit').then((d) => {",
  "    P = d;",
  "    const r = ranges()[1]; // 今月",
  "    Object.assign(st, { range: r.name, from: r.from, to: r.to });",
  "    $('from').value = r.from; $('to').value = r.to;",
  "    drawChips();",
  "    if (!P.courses.length) { $('out').append(el('div', { class: 'warn' }, 'まだ講座がありません。「⚙ 初期設定」で登録してください。')); return; }",
  "    load();",
  "  }).catch((e) => { $('out').append(el('div', { class: 'warn' }, '読み込めませんでした：' + errText(e))); });",
  "</script>",
  "</body>",
  "</html>",
].join('\n');
