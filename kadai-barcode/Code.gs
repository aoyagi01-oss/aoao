// 課題提出バーコード管理：Google スプレッドシートに貼って使うプログラム
// 使い方は kadai-barcode/README.md を見てください。
//
// しくみ：生徒のバーコードには「生徒ID」だけが入っています。
// どの課題の提出かは読み取るときに画面で選ぶので、課題や教材が変わっても同じバーコードを使い続けられます。

const SHEET = {
  HOWTO: '使い方',
  SETTINGS: '設定',
  STUDENTS: '生徒名簿',
  TASKS: '課題一覧',
  LOG: '提出記録',
  STATUS: '提出状況',
  SUMMARY: '集計',
  PRINT_MISSING: '印刷_未提出者',
  PRINT_NOTICE: '印刷_提出のお願い',
  PRINT_BARCODE: '印刷_バーコード',
};

const STUDENT_HEADERS = ['生徒ID（バーコード）', 'クラス', '番号', '氏名', 'ふりがな', '除外（転出などは ✓）'];
const TASK_HEADERS = ['課題ID', '教科', '課題名', '対象クラス（空欄＝全員）', '出した日', '締切日', 'メモ', '対象', '提出', 'うち遅れ', '未提出', '提出率'];
const LOG_HEADERS = ['読み取り日時', '課題ID', '教科', '課題名', '生徒ID', 'クラス', '番号', '氏名', '判定'];

const SETTING = {
  TITLE: '学校名・担当（印刷物の見出し）',
  SUBJECT: '提出状況・集計に出す教科（空欄＝すべて／複数は「,」区切り）',
  CLASS: '提出状況・集計に出すクラス（空欄＝すべて／複数は「,」区切り）',
  NOTICE: '「提出のお願い」に書く文',
};
const SETTING_DEFAULTS = [
  [SETTING.TITLE, ''],
  [SETTING.SUBJECT, ''],
  [SETTING.CLASS, ''],
  [SETTING.NOTICE, '次の課題がまだ提出されていません。できるだけ早く提出してください。'],
];

const COLOR = { header: '#cfe2f3', ok: '#d9ead3', late: '#fff2cc', overdue: '#f4cccc', notYet: '#ffffff', none: '#eeeeee' };

// ───────── メニュー ─────────

function onOpen() {
  SpreadsheetApp.getUi().createMenu('課題バーコード')
    .addItem('📷 読み取りを始める', 'openScanner')
    .addItem('🔄 提出状況・集計を更新', 'refreshAll')
    .addSeparator()
    .addItem('🖨 バーコードを印刷する（生徒・課題）', 'openBarcodePrint')
    .addItem('🖨 未提出者リスト・提出のお願いを作る', 'openListDialog')
    .addItem('🖨 バーコードをシートに作る（印刷画面が出ないとき）', 'makeBarcodeSheet')
    .addSeparator()
    .addItem('生徒IDを付ける（空欄の人だけ）', 'assignStudentIds')
    .addItem('初期設定（シートを作る）', 'setup')
    .addItem('お試しデータを入れる', 'insertSampleData')
    .addToUi();
}

// 課題一覧に課題名を書いたら、すぐに課題IDを付ける
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() === SHEET.TASKS && e.range.getRow() > 1) getTasks_();
  } catch (err) {
    // 単純トリガーでは失敗しても何もしない（読み取り時や集計時にも ID は付く）
  }
}

function openScanner() {
  setup_(false);
  const html = HtmlService.createHtmlOutputFromFile('Scan').setTitle('課題の読み取り');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openBarcodePrint() {
  setup_(false);
  const students = sortStudents_(getStudents_().filter(function (s) { return !s.excluded && s.id; }));
  const tasks = getTasks_().reverse();
  const t = HtmlService.createTemplateFromFile('Print');
  t.payload = toScriptJson_({
    title: getSettings_()[SETTING.TITLE] || '',
    students: students.map(function (s) { return { id: s.id, cls: s.cls, no: s.no, name: s.name }; }),
    tasks: tasks.map(function (k) { return { id: k.id, subject: k.subject, name: k.name, due: k.due }; }),
  });
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1000).setHeight(720), 'バーコードの印刷');
}

function openListDialog() {
  setup_(false);
  const tasks = getTasks_().reverse();
  const t = HtmlService.createTemplateFromFile('Select');
  t.payload = toScriptJson_({
    today: todayKey_(),
    tasks: tasks.map(function (k) { return { id: k.id, subject: k.subject, name: k.name, due: k.due, classes: k.classes.join(',') }; }),
    classes: classList_(getStudents_()),
  });
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(720).setHeight(640), '未提出者リスト・提出のお願い');
}

// ───────── 初期設定 ─────────

function setup() {
  setup_(true);
}

function setup_(showMessage) {
  const ss = SpreadsheetApp.getActive();
  let created = false;
  function make(name, headers, widths) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name, ss.getSheets().length);
      created = true;
    }
    if (headers && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground(COLOR.header).setWrap(true).setVerticalAlignment('middle');
      sh.setFrozenRows(1);
      (widths || []).forEach(function (w, i) { if (w) sh.setColumnWidth(i + 1, w); });
      created = true;
      return { sh: sh, fresh: true };
    }
    return { sh: sh, fresh: false };
  }

  const howto = make(SHEET.HOWTO);
  if (howto.sh.getLastRow() === 0) writeHowTo_(howto.sh);

  const settings = make(SHEET.SETTINGS, ['項目', '値'], [380, 460]);
  if (settings.fresh) settings.sh.getRange(2, 1, SETTING_DEFAULTS.length, 2).setValues(SETTING_DEFAULTS);

  const students = make(SHEET.STUDENTS, STUDENT_HEADERS, [140, 80, 60, 140, 160, 120]);
  if (students.fresh) {
    students.sh.getRange('A2:A').setNumberFormat('@'); // 先頭の 0 が消えないように文字として扱う
    students.sh.getRange('F2:F').insertCheckboxes();
  }

  const tasks = make(SHEET.TASKS, TASK_HEADERS, [70, 80, 220, 160, 90, 90, 160, 50, 50, 60, 60, 60]);
  if (tasks.fresh) {
    const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('日付を入れてください（例：2026/10/5）').build();
    tasks.sh.getRange('E2:F').setDataValidation(dateRule).setNumberFormat('yyyy/mm/dd');
    tasks.sh.getRange('L2:L').setNumberFormat('0%');
    tasks.sh.getRange('H1:L1').setBackground('#e0e0e0');
    tasks.sh.getRange('H1').setNote('H〜L列は「提出状況・集計を更新」で自動で書き込まれます。');
  }

  const log = make(SHEET.LOG, LOG_HEADERS, [150, 60, 80, 200, 110, 70, 50, 130, 70]);
  if (log.fresh) {
    log.sh.getRange('A2:A').setNumberFormat('yyyy/mm/dd hh:mm');
    log.sh.getRange('E2:E').setNumberFormat('@');
  }

  make(SHEET.STATUS);
  make(SHEET.SUMMARY);

  // 新しいスプレッドシートに最初からある空の「シート1」は消す
  ['シート1', 'Sheet1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  if (showMessage) {
    ss.setActiveSheet(ss.getSheetByName(SHEET.STUDENTS));
    SpreadsheetApp.getUi().alert(created
      ? '準備ができました。\n\n1. 「生徒名簿」にクラス・番号・氏名を貼り付ける（生徒IDは空欄でOK。自動で付きます）\n2. 「課題一覧」に課題を書く\n3. メニュー「課題バーコード」→「バーコードを印刷する」'
      : 'シートはすでにそろっています。');
  }
}

function writeHowTo_(sh) {
  const lines = [
    ['課題提出バーコード管理　使い方'],
    [''],
    ['■ 最初に1回だけ'],
    ['1. 「生徒名簿」にクラス・番号・氏名を貼り付ける。生徒IDは空欄でよい（自動で6けたの番号が付く）。学籍番号を使ってもよい。'],
    ['2. メニュー「課題バーコード」→「バーコードを印刷する」で生徒のバーコードを印刷し、ノートやファイルにはる（先生の手元用の名簿型もある）。'],
    ['   バーコードには生徒IDだけが入っているので、課題・教材・教科が変わっても、学年が上がっても同じバーコードを使える。'],
    [''],
    ['■ 課題を出すたび'],
    ['1. 「課題一覧」に 教科・課題名・対象クラス・締切日 を書く（課題IDは自動で付く）。対象クラスは「1-1,1-2」のように生徒名簿のクラスと同じ書き方で。空欄なら全員。'],
    [''],
    ['■ 回収するとき'],
    ['1. メニュー「課題バーコード」→「読み取りを始める」。右に出る画面で課題を選ぶ（課題バーコードを読んでも切りかわる）。'],
    ['2. 読み取り欄をクリックしてから、バーコードリーダーで次々に読む。ピッ＝提出、ピピッ＝締切後の提出、ブー＝エラー（未登録・2回目など）。'],
    ['3. 画面の「未提出」に、まだ出していない生徒が出る。間違えたら「直前の1件を取り消す」か「取り消しモード」。'],
    ['   ※ 読み取り欄からカーソルが外れると、シートのセルに文字が入ってしまうので注意（画面が赤くなって知らせる）。'],
    [''],
    ['■ 提出を促す・集計'],
    ['・「未提出者リスト・提出のお願いを作る」：課題とクラスを選ぶと、未提出者の一覧、または生徒ごとに切り取って渡す「提出のお願い」ができ、PDFでまとめて印刷できる。'],
    ['・「提出状況・集計を更新」：「提出状況」（生徒×課題の表）、「集計」（課題別・クラス別・教科別の提出率）、「課題一覧」の提出率が新しくなる。'],
    ['・「設定」で、表に出す教科・クラスをしぼれる。'],
    [''],
    ['■ ほかの先生に配るとき'],
    ['・このスプレッドシートを「ファイル → コピーを作成」してもらう。プログラムも一緒にコピーされる。'],
    ['・コピーした人は、生徒名簿・課題一覧・提出記録の中身を消して使う（最初の実行時に Google の許可画面が出たら、自分のアカウントで許可する）。'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  lines.forEach(function (l, i) { if (/^■/.test(l[0])) sh.getRange(i + 1, 1).setFontWeight('bold').setBackground('#eef3fb'); });
}

function insertSampleData() {
  setup_(false);
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('お試しデータ（架空の生徒6人・課題3つ）を入れます。あとで行ごと消してください。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName(SHEET.STUDENTS);
  const sample = [
    ['', '1-1', 1, '青木 さくら', 'あおき さくら', false],
    ['', '1-1', 2, '石田 はると', 'いしだ はると', false],
    ['', '1-1', 3, '上野 ゆい', 'うえの ゆい', false],
    ['', '1-2', 1, '江藤 そうた', 'えとう そうた', false],
    ['', '1-2', 2, '小川 めい', 'おがわ めい', false],
    ['', '1-2', 3, '加藤 れん', 'かとう れん', false],
  ];
  st.getRange(st.getLastRow() + 1, 1, sample.length, sample[0].length).setValues(sample);
  const now = new Date();
  function day(n) { const d = new Date(now); d.setDate(d.getDate() + n); return d; }
  const tk = ss.getSheetByName(SHEET.TASKS);
  tk.getRange(tk.getLastRow() + 1, 1, 3, 7).setValues([
    ['', '国語', '漢字ノート 第3回', '', day(-10), day(-3), ''],
    ['', '数学', 'ワーク p.20〜25', '1-1', day(-5), day(-1), ''],
    ['', '英語', 'Unit 2 ワークシート', '1-1,1-2', day(-1), day(5), ''],
  ]);
  getStudents_();
  getTasks_();
  ss.setActiveSheet(st);
  ui.alert('お試しデータを入れました。メニュー「読み取りを始める」で、画面の読み取り欄に生徒ID（例：100001）を打って Enter を押すと試せます。');
}

// ───────── データの読み書き ─────────

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」シートがありません。メニュー「初期設定（シートを作る）」を実行してください。');
  return sh;
}

function norm_(v) {
  return String(v === null || v === undefined ? '' : v).normalize('NFKC').replace(/\s+/g, '').toUpperCase();
}

function normClass_(v) {
  return String(v === null || v === undefined ? '' : v).normalize('NFKC').replace(/\s+/g, '');
}

function splitList_(v) {
  return String(v || '').normalize('NFKC').split(/[,、，;；\/／\n]+/).map(function (s) { return s.replace(/\s+/g, ''); }).filter(String);
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

function shortDate_(key) {
  if (!key) return '';
  const p = key.split('-');
  return Number(p[1]) + '/' + Number(p[2]);
}

function getStudents_() {
  const sh = sheet_(SHEET.STUDENTS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const range = sh.getRange(2, 1, last - 1, STUDENT_HEADERS.length);
  const vals = range.getValues();
  let maxNo = 100000;
  vals.forEach(function (r) {
    const id = norm_(r[0]);
    if (/^\d+$/.test(id) && Number(id) > maxNo && Number(id) < 1e12) maxNo = Number(id);
  });
  const list = [];
  vals.forEach(function (r, i) {
    const name = String(r[3]).trim();
    let id = norm_(r[0]);
    if (!name && !id) return;
    if (!id && name) {
      maxNo += 1;
      id = String(maxNo);
      sh.getRange(i + 2, 1).setNumberFormat('@').setValue(id);
    }
    const ex = r[5];
    list.push({
      id: id, cls: normClass_(r[1]), no: r[2], name: name, kana: String(r[4]).trim(),
      excluded: ex === true || /^(✓|✔|○|〇|1|TRUE|除外|転出|×)$/i.test(String(ex).trim()),
      row: i + 2,
    });
  });
  return list;
}

function sortStudents_(list) {
  return list.slice().sort(function (a, b) {
    const c = a.cls.localeCompare(b.cls, 'ja', { numeric: true });
    if (c) return c;
    return (Number(a.no) || 0) - (Number(b.no) || 0) || a.name.localeCompare(b.name, 'ja');
  });
}

function classList_(students) {
  const seen = {};
  students.forEach(function (s) { if (s.cls && !s.excluded) seen[s.cls] = true; });
  return Object.keys(seen).sort(function (a, b) { return a.localeCompare(b, 'ja', { numeric: true }); });
}

function getTasks_() {
  const sh = sheet_(SHEET.TASKS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 7).getValues();
  let maxNo = 0;
  vals.forEach(function (r) {
    const m = norm_(r[0]).match(/^K(\d+)$/);
    if (m && Number(m[1]) > maxNo) maxNo = Number(m[1]);
  });
  const seen = {};
  const tasks = [];
  vals.forEach(function (r, i) {
    const name = String(r[2]).trim();
    let id = norm_(r[0]);
    if (!name && !id) return;
    if (!id || seen[id]) { // 空欄、または行をコピーして ID が重なったときは新しい ID
      maxNo += 1;
      id = 'K' + ('00' + maxNo).slice(-3);
      sh.getRange(i + 2, 1).setValue(id);
    }
    seen[id] = true;
    tasks.push({
      id: id, subject: String(r[1]).trim(), name: name || '（課題名なし）',
      classes: splitList_(r[3]), given: dateKey_(r[4]), due: dateKey_(r[5]), memo: String(r[6]), row: i + 2,
    });
  });
  return tasks;
}

function taskLabel_(t) {
  return (t.subject ? '【' + t.subject + '】' : '') + t.name + (t.due ? '（締切 ' + shortDate_(t.due) + '）' : '');
}

function isTarget_(task, student) {
  if (student.excluded) return false;
  return !task.classes.length || task.classes.indexOf(student.cls) >= 0;
}

function readLog_() {
  const sh = sheet_(SHEET.LOG);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues().map(function (r, i) {
    return { time: r[0], taskId: norm_(r[1]), studentId: norm_(r[4]), row: i + 2 };
  }).filter(function (r) { return r.taskId && r.studentId; });
}

function isLate_(task, time) {
  return !!task.due && time instanceof Date && dateKey_(time) > task.due;
}

// 課題ごとに「だれが・いつ」出したか（同じ生徒が2回あれば最初の記録）
function submissionsOf_(task, logs) {
  const map = {};
  logs.forEach(function (r) {
    if (r.taskId === task.id && !map[r.studentId]) map[r.studentId] = { time: r.time, late: isLate_(task, r.time) };
  });
  return map;
}

function taskStats_(task, students, logs) {
  const subs = submissionsOf_(task, logs);
  let target = 0, submitted = 0, late = 0;
  const missing = [];
  sortStudents_(students).forEach(function (s) {
    if (!isTarget_(task, s)) return;
    target++;
    const sub = subs[s.id];
    if (sub) {
      submitted++;
      if (sub.late) late++;
    } else {
      missing.push({ cls: s.cls, no: s.no, name: s.name });
    }
  });
  return {
    taskId: task.id, label: taskLabel_(task), target: target, submitted: submitted, late: late,
    rate: target ? submitted / target : 0, missing: missing,
    overdue: !!task.due && todayKey_() > task.due,
  };
}

// ───────── 読み取り画面（Scan.html）から呼ばれる ─────────

function getScanInit() {
  setup_(false);
  const tasks = getTasks_().reverse(); // 新しい課題を上に
  const last = PropertiesService.getUserProperties().getProperty('lastTask:' + SpreadsheetApp.getActive().getId());
  const lastTask = tasks.some(function (t) { return t.id === last; }) ? last : (tasks[0] ? tasks[0].id : '');
  return {
    tasks: tasks.map(function (t) { return { id: t.id, label: t.id + ' ' + taskLabel_(t) }; }),
    taskId: lastTask,
    stats: lastTask ? getTaskStats(lastTask) : null,
  };
}

function getTaskStats(taskId) {
  const task = findTask_(getTasks_(), taskId);
  if (!task) return null;
  PropertiesService.getUserProperties().setProperty('lastTask:' + SpreadsheetApp.getActive().getId(), task.id);
  return taskStats_(task, getStudents_(), readLog_());
}

function findTask_(tasks, id) {
  const key = norm_(id);
  for (let i = 0; i < tasks.length; i++) if (tasks[i].id === key) return tasks[i];
  return null;
}

function recordScan(taskId, raw, undoMode) {
  const code = norm_(raw);
  if (!code) return { kind: 'empty' };
  const tasks = getTasks_();

  // 課題バーコードを読んだら、その課題に切りかえる
  const asTask = findTask_(tasks, code);
  if (asTask) {
    return { kind: 'task', taskId: asTask.id, message: '課題を切りかえました：' + taskLabel_(asTask), stats: getTaskStats(asTask.id) };
  }

  const task = findTask_(tasks, taskId);
  if (!task) return { kind: 'error', message: '先に課題を選んでください' };
  const students = getStudents_();
  let st = null;
  for (let i = 0; i < students.length; i++) if (students[i].id === code) { st = students[i]; break; }

  if (!st) {
    return { kind: 'error', message: '登録されていないバーコードです（' + code + '）', stats: taskStats_(task, students, readLog_()) };
  }

  const who = st.cls + ' ' + st.no + '番 ' + st.name;
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  let result;
  try {
    const sh = sheet_(SHEET.LOG);
    const logs = readLog_();
    const mine = logs.filter(function (r) { return r.taskId === task.id && r.studentId === st.id; });

    if (undoMode) {
      if (!mine.length) {
        result = { kind: 'error', message: '取り消す記録がありません：' + who };
      } else {
        mine.map(function (r) { return r.row; }).sort(function (a, b) { return b - a; }).forEach(function (row) { sh.deleteRow(row); });
        result = { kind: 'undone', message: '取り消しました：' + who };
      }
    } else if (mine.length) {
      const t = mine[0].time instanceof Date ? Utilities.formatDate(mine[0].time, tz_(), 'M/d HH:mm') : '';
      result = { kind: 'dup', message: 'すでに提出済み：' + who + (t ? '（' + t + '）' : '') };
    } else {
      const now = new Date();
      const late = isLate_(task, now);
      const target = isTarget_(task, st);
      const judge = !target ? '対象外' : (late ? '遅れ' : '期限内');
      sh.appendRow([now, task.id, task.subject, task.name, st.id, st.cls, st.no, st.name, judge]);
      sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(st.id);
      result = {
        kind: !target ? 'notarget' : (late ? 'late' : 'ok'),
        message: (!target ? '対象外のクラスですが記録しました：' : (late ? '締切後の提出：' : '提出：')) + who,
      };
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  result.stats = taskStats_(task, students, readLog_());
  return result;
}

// 選んでいる課題の、いちばん新しい記録を1件消す
function undoLast(taskId) {
  const task = findTask_(getTasks_(), taskId);
  if (!task) return { kind: 'error', message: '課題を選んでください' };
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  let result;
  try {
    const sh = sheet_(SHEET.LOG);
    const logs = readLog_().filter(function (r) { return r.taskId === task.id; });
    if (!logs.length) {
      result = { kind: 'error', message: 'この課題の記録はまだありません' };
    } else {
      const lastRow = logs[logs.length - 1].row;
      const v = sh.getRange(lastRow, 1, 1, LOG_HEADERS.length).getValues()[0];
      sh.deleteRow(lastRow);
      SpreadsheetApp.flush();
      result = { kind: 'undone', message: '取り消しました：' + v[5] + ' ' + v[6] + '番 ' + v[7] };
    }
  } finally {
    lock.releaseLock();
  }
  result.stats = taskStats_(task, getStudents_(), readLog_());
  return result;
}

// ───────── 集計 ─────────

function refreshAll() {
  setup_(false);
  const ss = SpreadsheetApp.getActive();
  const settings = getSettings_();
  const subjectFilter = splitList_(settings[SETTING.SUBJECT]);
  const classFilter = splitList_(settings[SETTING.CLASS]).map(normClass_);
  const allStudents = getStudents_();
  const allTasks = getTasks_();
  const logs = readLog_();
  const today = todayKey_();
  const warnings = [];

  const idCount = {};
  allStudents.forEach(function (s) { idCount[s.id] = (idCount[s.id] || 0) + 1; });
  const dupIds = Object.keys(idCount).filter(function (k) { return idCount[k] > 1; });
  if (dupIds.length) warnings.push('生徒IDが重なっています：' + dupIds.join(', '));

  const active = allStudents.filter(function (s) { return !s.excluded; });
  const subsByTask = {};
  allTasks.forEach(function (t) { subsByTask[t.id] = submissionsOf_(t, logs); });

  // 課題一覧の H〜L 列（全課題）
  const tkSheet = sheet_(SHEET.TASKS);
  allTasks.forEach(function (t) {
    const st = taskStats_(t, active, logs);
    tkSheet.getRange(t.row, 8, 1, 5).setValues([[st.target, st.submitted, st.late, st.target - st.submitted, st.target ? st.rate : '']]);
  });
  tkSheet.getRange('L2:L').setNumberFormat('0%');

  // 設定でしぼりこみ
  const tasks = allTasks.filter(function (t) { return !subjectFilter.length || subjectFilter.indexOf(t.subject.normalize('NFKC').replace(/\s+/g, '')) >= 0; });
  const students = sortStudents_(active.filter(function (s) { return !classFilter.length || classFilter.indexOf(s.cls) >= 0; }));

  writeStatusSheet_(ss, tasks, students, subsByTask, today);
  writeSummarySheet_(ss, tasks, students, subsByTask, today);

  ss.toast('提出状況・集計を更新しました' + (warnings.length ? '\n⚠ ' + warnings.join('\n⚠ ') : ''), '課題バーコード', warnings.length ? 15 : 5);
  return { ok: true, warnings: warnings };
}

function writeStatusSheet_(ss, tasks, students, subsByTask, today) {
  const sh = ss.getSheetByName(SHEET.STATUS);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  const fixed = ['クラス', '番号', '氏名', '提出率', '期限切れの未提出'];
  const header = fixed.concat(tasks.map(function (t) { return (t.subject ? t.subject + '\n' : '') + t.name + (t.due ? '\n〆' + shortDate_(t.due) : ''); }));
  const values = [header];
  const colors = [header.map(function () { return COLOR.header; })];

  students.forEach(function (s) {
    let considered = 0, done = 0, overdueMissing = 0;
    const cells = [], bg = [];
    tasks.forEach(function (t) {
      if (!isTarget_(t, s)) { cells.push('－'); bg.push(COLOR.none); return; }
      const sub = subsByTask[t.id][s.id];
      const overdue = !!t.due && today > t.due;
      if (sub) {
        considered++; done++;
        cells.push(sub.late ? '遅' : '○');
        bg.push(sub.late ? COLOR.late : COLOR.ok);
      } else {
        if (overdue || !t.due) considered++;
        if (overdue) overdueMissing++;
        cells.push('未');
        bg.push(overdue ? COLOR.overdue : COLOR.notYet);
      }
    });
    values.push([s.cls, s.no, s.name, considered ? done / considered : '', overdueMissing || ''].concat(cells));
    colors.push(['#ffffff', '#ffffff', '#ffffff', '#ffffff', overdueMissing ? COLOR.overdue : '#ffffff'].concat(bg));
  });

  ensureSize_(sh, values.length, header.length);
  const rng = sh.getRange(1, 1, values.length, header.length);
  rng.setValues(values).setBackgrounds(colors).setVerticalAlignment('middle');
  sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setWrap(true).setHorizontalAlignment('center');
  if (values.length > 1) {
    sh.getRange(2, 4, values.length - 1, 1).setNumberFormat('0%');
    if (tasks.length) sh.getRange(2, fixed.length + 1, values.length - 1, tasks.length).setHorizontalAlignment('center');
  }
  sh.setRowHeight(1, 70);
  sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 45); sh.setColumnWidth(3, 120); sh.setColumnWidth(4, 60); sh.setColumnWidth(5, 70);
  for (let c = 0; c < tasks.length; c++) sh.setColumnWidth(fixed.length + 1 + c, 90);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(3);
  if (values.length > 1) rng.createFilter();
  sh.getRange(1, 1).setNote('○＝提出　遅＝締切後に提出　未＝未提出（赤は締切を過ぎたもの）　－＝対象外\n提出率は「締切を過ぎた課題・締切のない課題・提出済みの課題」で計算。\n最終更新：' + Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm'));
}

function writeSummarySheet_(ss, tasks, students, subsByTask, today) {
  const sh = ss.getSheetByName(SHEET.SUMMARY);
  sh.clear();
  const classes = classList_(students);
  const header = ['課題ID', '教科', '課題名', '締切', '対象', '提出', 'うち遅れ', '未提出', '提出率'].concat(classes.map(function (c) { return c + '\n提出率'; }));
  const rows = [];
  const bySubject = {};

  tasks.forEach(function (t) {
    const subs = subsByTask[t.id];
    let target = 0, done = 0, late = 0;
    const perClass = {};
    students.forEach(function (s) {
      if (!isTarget_(t, s)) return;
      target++;
      const pc = perClass[s.cls] || (perClass[s.cls] = { target: 0, done: 0 });
      pc.target++;
      if (subs[s.id]) { done++; pc.done++; if (subs[s.id].late) late++; }
    });
    rows.push([t.id, t.subject, t.name, t.due ? shortDate_(t.due) : '', target, done, late, target - done, target ? done / target : '']
      .concat(classes.map(function (c) { return perClass[c] ? perClass[c].done / perClass[c].target : ''; })));
    const key = t.subject || '（教科なし）';
    const b = bySubject[key] || (bySubject[key] = { tasks: 0, target: 0, done: 0, late: 0 });
    b.tasks++; b.target += target; b.done += done; b.late += late;
  });

  ensureSize_(sh, rows.length + Object.keys(bySubject).length + 10, header.length);
  sh.getRange(1, 1).setValue('課題別の提出状況').setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground(COLOR.header).setWrap(true);
  if (rows.length) {
    sh.getRange(3, 1, rows.length, header.length).setValues(rows);
    sh.getRange(3, 9, rows.length, header.length - 8).setNumberFormat('0%');
  }

  const top = 3 + rows.length + 2;
  sh.getRange(top, 1).setValue('教科別の提出率').setFontWeight('bold').setFontSize(13);
  const subjHeader = ['教科', '課題数', '対象（のべ）', '提出（のべ）', 'うち遅れ', '提出率'];
  const subjRows = Object.keys(bySubject).map(function (k) {
    const b = bySubject[k];
    return [k, b.tasks, b.target, b.done, b.late, b.target ? b.done / b.target : ''];
  });
  sh.getRange(top + 1, 1, 1, subjHeader.length).setValues([subjHeader]).setFontWeight('bold').setBackground(COLOR.header);
  if (subjRows.length) {
    sh.getRange(top + 2, 1, subjRows.length, subjHeader.length).setValues(subjRows);
    sh.getRange(top + 2, 6, subjRows.length, 1).setNumberFormat('0%');
  }
  sh.setColumnWidth(3, 220);
  sh.setFrozenRows(2);
  sh.getRange(1, 3).setValue('最終更新：' + Utilities.formatDate(new Date(), tz_(), 'yyyy/MM/dd HH:mm')).setFontColor('#666666');
}

// ───────── 印刷用シート（Select.html から呼ばれる） ─────────

// opts: { type: 'missing' | 'notice', taskIds: [...], classes: [...], overdueOnly: bool }
function makePrintList(opts) {
  const ss = SpreadsheetApp.getActive();
  const settings = getSettings_();
  const title = String(settings[SETTING.TITLE] || '');
  const allTasks = getTasks_();
  const today = todayKey_();
  const tasks = (opts.taskIds || []).map(function (id) { return findTask_(allTasks, id); })
    .filter(function (t) { return t && (!opts.overdueOnly || (t.due && today > t.due)); });
  if (!tasks.length) throw new Error(opts.overdueOnly ? '選んだ課題の中に、締切を過ぎたものがありません' : '課題を1つ以上選んでください');
  tasks.sort(function (a, b) { return a.row - b.row; });

  const classes = (opts.classes || []).map(normClass_);
  const students = sortStudents_(getStudents_().filter(function (s) { return !s.excluded && (!classes.length || classes.indexOf(s.cls) >= 0); }));
  const logs = readLog_();
  const subs = {};
  tasks.forEach(function (t) { subs[t.id] = submissionsOf_(t, logs); });

  const name = opts.type === 'notice' ? SHEET.PRINT_NOTICE : SHEET.PRINT_MISSING;
  const sh = freshSheet_(ss, name);

  const count = opts.type === 'notice'
    ? writeNotices_(sh, tasks, students, subs, title, String(settings[SETTING.NOTICE] || ''))
    : writeMissingList_(sh, tasks, students, subs, title);

  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();
  return { ok: true, count: count, sheetName: name, pdfUrl: pdfUrl_(ss, sh) };
}

function writeMissingList_(sh, tasks, students, subs, title) {
  [90, 50, 180, 130, 230].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  let row = 1;
  sh.getRange(row, 1, 1, 5).merge().setValue('未提出者一覧' + (title ? '　' + title : '') + '　（' + Utilities.formatDate(new Date(), tz_(), 'yyyy/M/d HH:mm') + ' 現在）').setFontSize(14).setFontWeight('bold');
  row += 2;
  let total = 0;
  tasks.forEach(function (t) {
    const missing = students.filter(function (s) { return isTarget_(t, s) && !subs[t.id][s.id]; });
    const target = students.filter(function (s) { return isTarget_(t, s); }).length;
    ensureSize_(sh, row + missing.length + 5, 5);
    sh.getRange(row, 1, 1, 5).merge()
      .setValue(taskLabel_(t) + '　未提出 ' + missing.length + '人 ／ 対象 ' + target + '人')
      .setFontWeight('bold').setBackground('#eeeeee').setBorder(true, false, true, false, false, false);
    row++;
    if (!missing.length) {
      sh.getRange(row, 1).setValue('全員提出しています');
      row += 2;
      return;
    }
    sh.getRange(row, 1, 1, 5).setValues([['クラス', '番号', '氏名', '確認', 'メモ']]).setFontWeight('bold').setFontColor('#555555');
    row++;
    const vals = missing.map(function (s) { return [s.cls, s.no, s.name, '□', '']; });
    sh.getRange(row, 1, vals.length, 5).setValues(vals).setBorder(null, null, true, null, false, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    row += vals.length + 1;
    total += missing.length;
  });
  return total;
}

function writeNotices_(sh, tasks, students, subs, title, message) {
  [70, 220, 260, 100].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  let row = 1;
  let count = 0;
  students.forEach(function (s) {
    const missing = tasks.filter(function (t) { return isTarget_(t, s) && !subs[t.id][s.id]; });
    if (!missing.length) return;
    count++;
    ensureSize_(sh, row + missing.length + 6, 4);
    const start = row;
    sh.getRange(row, 1, 1, 4).merge().setValue('提出のお願い' + (title ? '　　' + title : '')).setFontWeight('bold').setFontSize(13);
    row++;
    sh.getRange(row, 1, 1, 4).merge().setValue(s.cls + '　' + s.no + '番　' + s.name + '　さん').setFontSize(13);
    row++;
    if (message) {
      sh.getRange(row, 1, 1, 4).merge().setValue(message).setWrap(true);
      row++;
    }
    const vals = missing.map(function (t) { return ['□', t.subject, t.name, t.due ? '締切 ' + shortDate_(t.due) : '']; });
    sh.getRange(row, 1, vals.length, 4).setValues(vals);
    sh.getRange(row, 1, vals.length, 1).setHorizontalAlignment('center');
    row += vals.length;
    sh.getRange(start, 1, row - start, 4).setBorder(true, true, true, true, false, false, '#888888', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, 1, 1, 4).merge().setValue('✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -').setFontColor('#999999').setHorizontalAlignment('center');
    row++;
  });
  if (!count) sh.getRange(1, 1).setValue('選んだ課題・クラスでは、未提出の生徒はいません。');
  return count;
}

function pdfUrl_(ss, sh) {
  return 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf&gid=' + sh.getSheetId() +
    '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&fzr=false' +
    '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';
}

// ───────── バーコードをシートに作る（予備：バーコード用フォントを使う） ─────────
// 印刷画面（Print.html）がうまく開けないときのための方法。Code39 という形式で、ほぼすべてのバーコードリーダーで読めます。

function makeBarcodeSheet() {
  setup_(false);
  const ss = SpreadsheetApp.getActive();
  const students = sortStudents_(getStudents_().filter(function (s) { return !s.excluded && s.id; }));
  const bad = students.filter(function (s) { return !/^[0-9A-Z\-. $\/+%]+$/.test(s.id); });
  if (bad.length) {
    SpreadsheetApp.getUi().alert('次の生徒IDはこの方法では印刷できません（数字・英大文字・-のみ可）：\n' + bad.map(function (s) { return s.name + '：' + s.id; }).join('\n'));
    return;
  }
  const sh = freshSheet_(ss, SHEET.PRINT_BARCODE);
  const COLS = 3;
  for (let c = 1; c <= COLS; c++) sh.setColumnWidth(c, 240);
  ensureSize_(sh, Math.ceil(students.length / COLS) * 3, COLS);
  const values = [];
  for (let i = 0; i < students.length; i += COLS) {
    const label = [], code = [], gap = [];
    for (let c = 0; c < COLS; c++) {
      const s = students[i + c];
      label.push(s ? s.cls + ' ' + s.no + '番 ' + s.name : '');
      code.push(s ? '*' + s.id + '*' : '');
      gap.push('');
    }
    values.push(label, code, gap);
  }
  if (!values.length) {
    SpreadsheetApp.getUi().alert('生徒名簿に生徒がいません。');
    return;
  }
  sh.getRange(1, 1, values.length, COLS).setValues(values).setHorizontalAlignment('center').setVerticalAlignment('middle');
  for (let r = 1; r <= values.length; r += 3) {
    sh.getRange(r, 1, 1, COLS).setFontSize(10);
    sh.getRange(r + 1, 1, 1, COLS).setFontFamily('Libre Barcode 39 Text').setFontSize(40);
    sh.setRowHeight(r, 22);
    sh.setRowHeight(r + 1, 70);
    sh.setRowHeight(r + 2, 14);
  }
  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();
  const html = HtmlService.createHtmlOutput(
    '<p style="font-family:sans-serif">「' + SHEET.PRINT_BARCODE + '」シートにバーコードを作りました。</p>' +
    '<p style="font-family:sans-serif"><a href="' + pdfUrl_(ss, sh) + '" target="_blank">PDFで開いて印刷する</a></p>' +
    '<p style="font-family:sans-serif;color:#666;font-size:12px">バーコードが文字のまま見えるときは、少し待ってから開き直してください。</p>'
  ).setWidth(380).setHeight(180);
  SpreadsheetApp.getUi().showModalDialog(html, 'バーコード（シート版）');
}

// ───────── 生徒ID ─────────

function assignStudentIds() {
  setup_(false);
  const list = getStudents_(); // 空欄の人にはここで ID が付く
  SpreadsheetApp.getActive().toast(list.length + '人分の生徒IDがそろっています。', '課題バーコード', 5);
}

// ───────── 共通 ─────────

// HTML の中に別ファイル（JsBarcode.html）を差しこむ
function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// 印刷用シートは毎回作り直す（前回の結合セルや書式が残らないように）
function freshSheet_(ss, name) {
  const old = ss.getSheetByName(name);
  if (old) {
    if (ss.getActiveSheet().getSheetId() === old.getSheetId()) ss.setActiveSheet(ss.getSheets()[0]);
    ss.deleteSheet(old);
  }
  return ss.insertSheet(name, ss.getSheets().length);
}

function ensureSize_(sh, rows, cols) {
  if (sh.getMaxRows() < rows) sh.insertRowsAfter(sh.getMaxRows(), rows - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
}

function toScriptJson_(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}
