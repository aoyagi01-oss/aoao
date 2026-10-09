// 授業の出欠：Google スプレッドシートに貼り、スマホで開くウェブページとして使うプログラム
// 使い方は shukketsu/README.md を見てください。
//
// しくみ：記録するのは「授業（日付・時限・講座）」と、その授業で出席以外だった生徒だけ。
// 何も付けなかった生徒は出席です。

const SHEET = {
  HOWTO: '使い方',
  STUDENTS: '名簿',
  COURSES: '講座',
  LESSONS: '授業記録',
  MARKS: '出欠',
  SETTINGS: '設定',
};

// このプログラムの版（ページのいちばん下に出ます。貼り直しが反映されたかの確認用）
const VERSION = '10/9-1';

const STUDENT_HEADERS = ['学籍番号', 'クラス', '番号', '氏名', 'ふりがな', '除外（転出などは ✓）'];
const COURSE_HEADERS = ['講座名', '対象（クラス・学籍番号を「,」区切り）', '注意する欠課時数（空欄＝なし）', 'メモ'];
const LESSON_HEADERS = ['授業ID', '日付', '時限', '講座名', '対象', '出席', '欠課', '遅刻', '早退', '公欠', '出停・忌引', '授業メモ', '記録日時', '記録者'];
const MARK_HEADERS = ['授業ID', '日付', '時限', '講座名', '学籍番号', 'クラス', '番号', '氏名', '区分', '記録日時'];

// 出席以外の区分。公欠・出停・忌引は欠課時数に数えない
const KINDS = [
  { code: 'A', label: '欠課', short: '欠' },
  { code: 'L', label: '遅刻', short: '遅' },
  { code: 'E', label: '早退', short: '早' },
  { code: 'K', label: '公欠', short: '公' },
  { code: 'S', label: '出停・忌引', short: '停' },
];
const KIND_CODES = KINDS.map(function (k) { return k.code; });

const SETTING = {
  PERIODS: '1日の時限の数',
  LATE: '遅刻・早退を何回で欠課1にするか（空欄＝しない）',
  TERMS: '学期の区切り（集計のボタンになります）',
};
const SETTING_DEFAULTS = [
  [SETTING.PERIODS, 6],
  [SETTING.LATE, ''],
  [SETTING.TERMS, '1学期 4/1-7/31, 2学期 8/1-12/31, 3学期 1/1-3/31'],
];

const COLOR = { header: '#cfe2f3', warn: '#f4cccc' };

// ───────── メニューとページ ─────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📋 出欠')
    .addItem('📱 出欠ページを開く', 'showUrl')
    .addSeparator()
    .addItem('お試しデータを入れる', 'insertSampleData')
    .addItem('シートを作り直す（消したシートを戻す）', 'setup')
    .addToUi();
}

function doGet() {
  setup_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('授業の出欠')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function showUrl() {
  setup_();
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? '<p style="font-family:sans-serif;font-size:14px">スマホでは、このURLをブックマーク（ホーム画面に追加）しておくと便利です。</p>' +
      '<p style="font-family:sans-serif;font-size:14px;word-break:break-all"><a href="' + url + '" target="_blank">' + url + '</a></p>'
    : '<p style="font-family:sans-serif;font-size:14px">まだウェブページとして公開されていません。<br>拡張機能 → Apps Script →「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」で公開してください（README の手順）。</p>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(200), '出欠ページ');
}

// ───────── ページから呼ぶ処理 ─────────

// はじめに読みこむデータ（名簿・講座・授業の一覧・今年度の生徒ごとの回数）
function getInit() {
  setup_();
  const st = getSettings_();
  const students = getStudents_().filter(function (s) { return !s.excluded; });
  const courses = getCourses_();
  const lessons = readLessons_();
  const today = todayKey_();
  const fy = fiscalYear_(today);
  const from = fy + '-04-01', to = (fy + 1) + '-03-31';

  // 今年度の、講座ごと・生徒ごとの回数（出欠をとる画面に「これまで 欠課3」と出す）
  const inYear = {};
  const held = {};
  lessons.forEach(function (l) {
    if (l.date < from || l.date > to) return;
    inYear[l.id] = l;
    held[l.course] = (held[l.course] || 0) + 1;
  });
  const tally = {};
  readMarks_().forEach(function (m) {
    const l = inYear[m.id];
    if (!l) return;
    const byKey = tally[l.course] = tally[l.course] || {};
    const c = byKey[m.key] = byKey[m.key] || {};
    c[m.code] = (c[m.code] || 0) + 1;
  });

  return {
    version: VERSION,
    today: today,
    fy: fy,
    periods: periodCount_(st),
    lateRule: lateRule_(st),
    terms: terms_(st, fy),
    termsText: String(st[SETTING.TERMS] || ''),
    kinds: KINDS,
    classes: classList_(students),
    students: students.map(function (s) { return { key: s.key, cls: s.cls, no: s.no, name: s.name, kana: s.kana }; }),
    courses: courses.map(function (c) {
      return { name: c.name, target: c.target, limit: c.limit, memo: c.memo, members: members_(c, students).map(function (s) { return s.key; }) };
    }),
    lessons: lessons.map(function (l) {
      return { id: l.id, date: l.date, period: l.period, course: l.course, memo: l.memo, n: l.n };
    }),
    tally: tally,
    held: held,
    sheetUrl: SpreadsheetApp.getActive().getUrl(),
  };
}

// 1つの授業の記録（出席以外の生徒と、授業メモ）
function getLesson(id) {
  const l = readLessons_().filter(function (x) { return x.id === id; })[0];
  if (!l) throw new Error('この授業の記録が見つかりません（シートで消されたかもしれません）');
  const marks = {};
  readMarks_().forEach(function (m) { if (m.id === id) marks[m.key] = m.code; });
  return { id: l.id, date: l.date, period: l.period, course: l.course, memo: l.memo, marks: marks };
}

// o: { id（直すときだけ）, date: 'yyyy-MM-dd', period, course, marks: { 学籍番号: 'A' など }, memo }
// 同じ日・時限・講座の記録がすでにあれば、それを書きかえる
function saveLesson(o) {
  const date = dateKey_(o.date);
  const period = Number(o.period);
  if (!date) throw new Error('日付がわかりません');
  if (!(period >= 1 && period <= 20)) throw new Error('時限を選んでください');
  const course = getCourses_().filter(function (c) { return c.name === o.course; })[0];
  if (!course) throw new Error('講座「' + o.course + '」がありません');

  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const all = getStudents_();
    const byKey = {};
    all.forEach(function (s) { byKey[s.key] = s; });
    const members = members_(course, all.filter(function (s) { return !s.excluded; }));

    const lessons = readLessons_();
    let cur = o.id ? lessons.filter(function (l) { return l.id === o.id; })[0] : null;
    if (!cur) cur = lessons.filter(function (l) { return l.date === date && l.period === period && l.course === course.name; })[0];
    const id = cur ? cur.id : newId_(date, period);
    deleteMarks_(id);

    const now = new Date();
    const n = { A: 0, L: 0, E: 0, K: 0, S: 0 };
    const rows = [];
    Object.keys(o.marks || {}).forEach(function (key) {
      const code = o.marks[key];
      const s = byKey[key];
      if (!s || KIND_CODES.indexOf(code) < 0) return;
      n[code]++;
      rows.push([id, toDate_(date), period, course.name, s.key, s.cls, s.no, s.name, kindLabel_(code), now]);
    });
    if (rows.length) {
      const sh = sheet_(SHEET.MARKS);
      const start = sh.getLastRow() + 1;
      ensureSize_(sh, start + rows.length, MARK_HEADERS.length);
      sh.getRange(start, 5, rows.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, rows.length, MARK_HEADERS.length).setValues(rows);
    }

    const target = members.length;
    const lessonRow = [id, toDate_(date), period, course.name, target, Math.max(0, target - n.A - n.K - n.S),
      n.A, n.L, n.E, n.K, n.S, String(o.memo || '').trim(), now, Session.getActiveUser().getEmail() || ''];
    const lsh = sheet_(SHEET.LESSONS);
    if (cur) {
      lsh.getRange(cur.row, 1, 1, LESSON_HEADERS.length).setValues([lessonRow]);
    } else {
      lsh.appendRow(lessonRow);
    }
    SpreadsheetApp.flush();
    return { id: id, init: getInit() };
  } finally {
    lock.releaseLock();
  }
}

function deleteLesson(id) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    deleteMarks_(id);
    const l = readLessons_().filter(function (x) { return x.id === id; })[0];
    if (l) sheet_(SHEET.LESSONS).deleteRow(l.row);
    SpreadsheetApp.flush();
    return getInit();
  } finally {
    lock.releaseLock();
  }
}

// 講座の集計。o: { course, from: 'yyyy-MM-dd', to: 'yyyy-MM-dd' }
function getSummary(o) {
  const st = getSettings_();
  const rule = lateRule_(st);
  const course = getCourses_().filter(function (c) { return c.name === o.course; })[0];
  if (!course) throw new Error('講座「' + o.course + '」がありません');
  const from = dateKey_(o.from) || '0000-00-00';
  const to = dateKey_(o.to) || '9999-12-31';
  const ids = {};
  let held = 0;
  readLessons_().forEach(function (l) {
    if (l.course !== course.name || l.date < from || l.date > to) return;
    ids[l.id] = true;
    held++;
  });
  const counts = {};
  readMarks_().forEach(function (m) {
    if (!ids[m.id]) return;
    const c = counts[m.key] = counts[m.key] || { A: 0, L: 0, E: 0, K: 0, S: 0 };
    c[m.code]++;
  });
  const rows = members_(course, getStudents_().filter(function (s) { return !s.excluded; })).map(function (s) {
    const c = counts[s.key] || { A: 0, L: 0, E: 0, K: 0, S: 0 };
    const kekka = c.A + (rule ? Math.floor((c.L + c.E) / rule) : 0);
    return { key: s.key, cls: s.cls, no: s.no, name: s.name, A: c.A, L: c.L, E: c.E, K: c.K, S: c.S, kekka: kekka, should: held - c.K - c.S };
  });
  return { course: course.name, limit: course.limit, from: o.from || '', to: o.to || '', held: held, lateRule: rule, rows: rows };
}

// 集計をシートに書き出す（印刷・成績処理用）
function exportSummary(o) {
  const sum = getSummary(o);
  const ss = SpreadsheetApp.getActive();
  const name = ('集計（' + sum.course + '）').replace(/[\[\]\*\?\/\\:]/g, '_').slice(0, 90);
  const sh = freshSheet_(ss, name);
  const period = (sum.from || sum.to) ? shortDate_(sum.from) + '〜' + shortDate_(sum.to) : '全期間';
  const head = ['クラス', '番号', '学籍番号', '氏名', '欠課', '遅刻', '早退', '公欠', '出停・忌引', '欠課時数', '出席すべき時数'];
  const note = sum.lateRule ? '（欠課時数＝欠課＋遅刻・早退' + sum.lateRule + '回で1）' : '（欠課時数＝欠課）';
  sh.getRange(1, 1, 2, 1).setValues([[sum.course + '　出欠の集計'], [period + '　授業 ' + sum.held + '回　' + note]]);
  sh.getRange(1, 1).setFontSize(14).setFontWeight('bold');
  sh.getRange(4, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground(COLOR.header);
  if (sum.rows.length) {
    const vals = sum.rows.map(function (r) { return [r.cls, r.no, r.key, r.name, r.A, r.L, r.E, r.K, r.S, r.kekka, r.should]; });
    sh.getRange(5, 3, vals.length, 1).setNumberFormat('@');
    sh.getRange(5, 1, vals.length, head.length).setValues(vals);
    if (sum.limit) {
      sum.rows.forEach(function (r, i) { if (r.kekka >= sum.limit) sh.getRange(5 + i, 1, 1, head.length).setBackground(COLOR.warn); });
    }
  }
  sh.setFrozenRows(4);
  sh.setColumnWidth(4, 140);
  SpreadsheetApp.flush();
  return ss.getUrl() + '#gid=' + sh.getSheetId();
}

// 講座を追加・変更する。o: { oldName（変更のとき）, name, target, limit, memo }
function saveCourse(o) {
  const name = String(o.name || '').trim();
  if (!name) throw new Error('講座名を入れてください');
  const target = splitList_(o.target).join(', ');
  if (!target) throw new Error('対象のクラス（または学籍番号）を選んでください');
  const limit = Number(String(o.limit || '').normalize('NFKC')) || '';
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const courses = getCourses_();
    const old = o.oldName ? courses.filter(function (c) { return c.name === o.oldName; })[0] : null;
    if (courses.some(function (c) { return c.name === name && c !== old; })) throw new Error('「' + name + '」という講座はもうあります');
    const sh = sheet_(SHEET.COURSES);
    const row = [name, target, limit, String(o.memo || '').trim()];
    if (old) {
      sh.getRange(old.row, 2).setNumberFormat('@');
      sh.getRange(old.row, 1, 1, row.length).setValues([row]);
      if (old.name !== name) renameCourse_(old.name, name);
    } else {
      const r = lastDataRow_(sh, 1) + 1;
      ensureSize_(sh, r, COURSE_HEADERS.length);
      sh.getRange(r, 2).setNumberFormat('@');
      sh.getRange(r, 1, 1, row.length).setValues([row]);
    }
    SpreadsheetApp.flush();
    return getInit();
  } finally {
    lock.releaseLock();
  }
}

// 講座の設定だけを消す（授業記録・出欠はシートに残る）
function deleteCourse(name) {
  const c = getCourses_().filter(function (x) { return x.name === name; })[0];
  if (c) sheet_(SHEET.COURSES).deleteRow(c.row);
  return getInit();
}

function saveSettings(o) {
  const periods = Math.round(Number(String(o.periods || '').normalize('NFKC')));
  if (!(periods >= 1 && periods <= 12)) throw new Error('時限の数は 1〜12 で入れてください');
  const late = Number(String(o.lateRule || '').normalize('NFKC')) || '';
  const terms = String(o.terms || '').trim();
  terms_({ [SETTING.TERMS]: terms }, fiscalYear_(todayKey_()), true); // 読めない書き方ならここでエラー
  putSetting_(SETTING.PERIODS, periods);
  putSetting_(SETTING.LATE, late);
  putSetting_(SETTING.TERMS, terms);
  return getInit();
}

// ───────── 名簿 ─────────

// rows: [{ gakuseki, name, kana }]（ページで貼り付けた表を読み取ったもの）
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
  }).filter(function (p) { return p.name; });
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
    return { added: fresh.length, updated: updated, same: plan.length - fresh.length - updated, init: getInit() };
  } finally {
    lock.releaseLock();
  }
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
  const courses = make(SHEET.COURSES, COURSE_HEADERS, [180, 280, 140, 200]);
  if (courses.fresh) courses.sh.getRange('A:B').setNumberFormat('@');
  const lessons = make(SHEET.LESSONS, LESSON_HEADERS, [150, 90, 50, 160]);
  if (lessons.fresh) {
    lessons.sh.getRange('A:A').setNumberFormat('@');
    lessons.sh.getRange('D:D').setNumberFormat('@');
    lessons.sh.getRange('B:B').setNumberFormat('yyyy/mm/dd');
    lessons.sh.getRange('M:M').setNumberFormat('yyyy/mm/dd hh:mm');
  }
  const marks = make(SHEET.MARKS, MARK_HEADERS, [150, 90, 50, 160, 80, 60, 50, 140, 90]);
  if (marks.fresh) {
    marks.sh.getRange('A:A').setNumberFormat('@');
    marks.sh.getRange('D:D').setNumberFormat('@');
    marks.sh.getRange('B:B').setNumberFormat('yyyy/mm/dd');
    marks.sh.getRange('E:F').setNumberFormat('@');
    marks.sh.getRange('J:J').setNumberFormat('yyyy/mm/dd hh:mm');
    marks.sh.getRange(2, 9, marks.sh.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(KINDS.map(function (k) { return k.label; }), true).setAllowInvalid(true).build());
  }
  const settings = make(SHEET.SETTINGS, ['項目', '値'], [380, 420]);
  if (settings.fresh) settings.sh.getRange(2, 1, SETTING_DEFAULTS.length, 2).setValues(SETTING_DEFAULTS);
}

function writeHowTo_(sh) {
  const lines = [
    ['授業の出欠'],
    [''],
    ['ふだんは、スマホで「出欠ページ」を開いて記録します（メニュー「📋 出欠」→「📱 出欠ページを開く」でURLが出ます）。'],
    [''],
    ['はじめに（1回だけ）'],
    ['1. 出欠ページの「⚙ 設定」→「名簿を貼り付ける」で、生徒の名簿を貼る（Excel や校務システムの表をコピーして貼るだけ）'],
    ['2. 「⚙ 設定」→「講座を追加」で、自分の授業（例：1-1 英語コミュⅠ）と対象のクラスを登録する'],
    [''],
    ['授業のたびに'],
    ['1. 時限を押す → 講座を押す（先週の同じ曜日・時限の講座には印が付きます）'],
    ['2. 全員「出席」から始まるので、休んだ・遅れた生徒の名前だけ押す（上のボタンで 欠課／遅刻／早退／公欠／出停・忌引 を切りかえ）'],
    ['3. 「保存」を押す'],
    [''],
    ['シートについて'],
    ['・「授業記録」：1回の授業が1行。「出欠」：出席以外だった生徒が1行ずつ（出席の生徒は記録しません）'],
    ['・シートを直接直してもかまいません。区分は「欠課」「遅刻」「早退」「公欠」「出停・忌引」のどれかで書きます'],
    ['・公欠・出停・忌引は欠課時数に数えません。「設定」で、遅刻・早退を何回で欠課1にするかを決められます'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  [5, 9, 14].forEach(function (r) { sh.getRange(r, 1).setFontWeight('bold'); });
  sh.setColumnWidth(1, 900);
}

// ───────── お試しデータ ─────────

function insertSampleData() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('お試しの名簿（2クラス・20人）と講座・授業記録を入れます。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  setup_();
  const family = ['青木', '石田', '上野', '江口', '大川', '加藤', '木村', '工藤', '小林', '佐藤'];
  const given = ['さくら', 'はると', 'ゆい', 'そうた', 'ひなた', 'りく', 'めい', 'ゆうと', 'あおい', 'こはる'];
  const rows = [];
  [1, 2].forEach(function (c) {
    family.forEach(function (f, i) { rows.push({ gakuseki: '1' + c + ('0' + (i + 1)).slice(-2), name: f + ' ' + given[(i + c * 3) % 10], kana: '' }); });
  });
  importRoster(rows);
  if (!getCourses_().some(function (c) { return c.name === '1-1 英語コミュⅠ'; })) saveCourse({ name: '1-1 英語コミュⅠ', target: '1-1', limit: 10 });
  if (!getCourses_().some(function (c) { return c.name === '1-2 英語コミュⅠ'; })) saveCourse({ name: '1-2 英語コミュⅠ', target: '1-2', limit: 10 });
  const today = todayKey_();
  [3, 2, 1].forEach(function (back, i) {
    const d = addDaysKey_(today, -7 * back);
    saveLesson({ date: d, period: 2, course: '1-1 英語コミュⅠ', marks: i === 0 ? { '1103': 'A', '1107': 'L' } : { '1103': 'A' }, memo: 'お試し' });
    saveLesson({ date: d, period: 4, course: '1-2 英語コミュⅠ', marks: { '1205': i === 2 ? 'K' : 'E' }, memo: 'お試し' });
  });
  ui.alert('お試しデータを入れました。出欠ページを開いて（または再読み込みして）ためしてください。');
}

// ───────── 読み書きの部品 ─────────

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」シートがありません。スプレッドシートのメニュー「📋 出欠」→「シートを作り直す」を押してください。');
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
      key: gakuseki || norm_(cls && no !== '' ? cls + '-' + no : name),
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
    out.push({ name: name, target: splitList_(r[1]).join(', '), limit: Number(r[2]) || 0, memo: String(r[3] || ''), row: i + 2 });
  });
  return out;
}

// 講座の対象の生徒：「1-1, 1-2」のようなクラスと、「1105」のような学籍番号（選択授業など）
function members_(course, students) {
  const classes = {}, ids = {};
  splitList_(course.target).forEach(function (t) {
    if (/^\d{4,}$/.test(t)) ids[norm_(t)] = true; else classes[normClass_(t)] = true;
  });
  return students.filter(function (s) { return classes[s.cls] || (s.gakuseki && ids[s.gakuseki]); });
}

function readLessons_() {
  const sh = sheet_(SHEET.LESSONS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const out = [];
  sh.getRange(2, 1, last - 1, LESSON_HEADERS.length).getValues().forEach(function (r, i) {
    const id = String(r[0]).trim();
    const date = dateKey_(r[1]);
    if (!id || !date) return;
    out.push({
      id: id, date: date, period: Number(r[2]) || 0, course: String(r[3]).trim(), memo: String(r[11] || ''),
      n: { target: Number(r[4]) || 0, A: Number(r[6]) || 0, L: Number(r[7]) || 0, E: Number(r[8]) || 0, K: Number(r[9]) || 0, S: Number(r[10]) || 0 },
      row: i + 2,
    });
  });
  return out;
}

function readMarks_() {
  const sh = sheet_(SHEET.MARKS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const out = [];
  sh.getRange(2, 1, last - 1, MARK_HEADERS.length).getValues().forEach(function (r, i) {
    const id = String(r[0]).trim();
    const code = kindCode_(r[8]);
    if (!id || !code) return;
    out.push({ id: id, key: norm_(r[4]), code: code, row: i + 2 });
  });
  return out;
}

// その授業の「出欠」の行を消す（つながった行はまとめて消す）
function deleteMarks_(id) {
  const sh = sheet_(SHEET.MARKS);
  const rows = readMarks_().filter(function (m) { return m.id === id; }).map(function (m) { return m.row; });
  // 区分が読めない行も、同じ授業IDなら消す
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) {
      if (String(r[0]).trim() === id && rows.indexOf(i + 2) < 0) rows.push(i + 2);
    });
  }
  rows.sort(function (a, b) { return b - a; });
  let i = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
    sh.deleteRows(rows[j], rows[i] - rows[j] + 1);
    i = j + 1;
  }
}

function renameCourse_(from, to) {
  [[SHEET.LESSONS, 4], [SHEET.MARKS, 4]].forEach(function (p) {
    const sh = sheet_(p[0]);
    const last = sh.getLastRow();
    if (last < 2) return;
    const range = sh.getRange(2, p[1], last - 1, 1);
    const vals = range.getValues();
    let changed = false;
    vals.forEach(function (r) { if (String(r[0]).trim() === from) { r[0] = to; changed = true; } });
    if (changed) range.setValues(vals);
  });
}

function newId_(date, period) {
  return date.replace(/-/g, '') + '-' + period + '-' + Utilities.getUuid().slice(0, 6);
}

function kindCode_(v) {
  const t = String(v || '').normalize('NFKC').trim();
  if (!t) return '';
  for (let i = 0; i < KINDS.length; i++) {
    const k = KINDS[i];
    if (t === k.label || t === k.short || t.toUpperCase() === k.code) return k.code;
  }
  if (/欠席|欠課/.test(t)) return 'A';
  if (/出停|忌引|出席停止/.test(t)) return 'S';
  return '';
}

function kindLabel_(code) {
  return KINDS.filter(function (k) { return k.code === code; })[0].label;
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

function periodCount_(st) {
  const n = Math.round(Number(String(st[SETTING.PERIODS]).normalize('NFKC')));
  return n >= 1 && n <= 12 ? n : 6;
}

function lateRule_(st) {
  const n = Math.round(Number(String(st[SETTING.LATE] || '').normalize('NFKC')));
  return n >= 1 ? n : 0;
}

// 「1学期 4/1-7/31, 2学期 8/1-12/31, 3学期 1/1-3/31」→ その年度の日付の範囲（1〜3月は次の年）
function terms_(st, fy, strict) {
  const text = String(st[SETTING.TERMS] || '').normalize('NFKC').replace(/(\d{1,2})月(\d{1,2})日/g, '$1/$2');
  const out = [];
  text.split(/[,、;\n]+/).forEach(function (part) {
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
