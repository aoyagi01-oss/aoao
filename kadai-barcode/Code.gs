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
  PRINT_NOTICE: '印刷_課題提出について',
  PRINT_BARCODE: '印刷_バーコード',
  EXCUSE: '欠席・配慮',
};

// 生徒IDはバーコードの中身（ずっと変えない）。学籍番号は毎年変わってよい（1101＝1年1組1番）
const STUDENT_HEADERS = ['生徒ID（バーコード・変えない）', '学籍番号', 'クラス', '番号', '氏名', 'ふりがな', '除外（転出などは ✓）'];
const TASK_HEADERS = ['課題ID', '教科', '課題名', '対象クラス（空欄＝全員）', '出した日', '締切日', '締切後の扱い', 'メモ', '担当', '対象', '提出', 'うち遅れ', '未提出', '提出率'];
// このプログラムの版（ホーム画面のいちばん下に出ます。貼り直しが反映されたかの確認用）
const VERSION = '10/8-3';
const TASK_STAT_COL = 10; // J列から右は自動で書きこむ

// 締切後に出したときの扱い（課題ごと。課題を登録するときに選ぶ）
//  遅れとして記録／◯日まで期限内（欠席などへの配慮）／区別しない
const LATE_CHOICES = ['遅れとして記録', '1日まで期限内', '2日まで期限内', '3日まで期限内', '5日まで期限内', '7日まで期限内', '区別しない'];
function latePolicy_(text) {
  const t = String(text || '').normalize('NFKC');
  if (/区別しない|しない|なし/.test(t)) return { mode: 'none', days: 0, label: '区別しない' };
  const m = t.match(/(\d+)/);
  if (m && Number(m[1]) > 0) return { mode: 'grace', days: Math.min(30, Number(m[1])), label: Number(m[1]) + '日まで期限内' };
  return { mode: 'strict', days: 0, label: '遅れとして記録' };
}
const LOG_HEADERS = ['読み取り日時', '課題ID', '教科', '課題名', '生徒ID', '学籍番号', 'クラス', '番号', '氏名', '判定'];

const SETTING = {
  TITLE: '学校名・担当（印刷物の見出し）',
  SUBJECT: '提出状況・集計に出す教科（空欄＝すべて／複数は「,」区切り）',
  CLASS: '提出状況・集計に出すクラス（空欄＝すべて／複数は「,」区切り）',
  NOTICE: '「課題提出について」に書く文',
  MARK: '提出状況（教務手帳）の書き方（「日付」または「○」）',
  INCOMPLETE: '未完成の数え方（「未提出」または「提出」）',
  LATE: '締切後の扱い（課題一覧で空欄のとき）',
  TEACHER: '担当の先生の名前（「課題提出について」に出る）',
};
const SETTING_DEFAULTS = [
  [SETTING.TITLE, ''],
  [SETTING.SUBJECT, ''],
  [SETTING.CLASS, ''],
  [SETTING.NOTICE, '次の課題がまだ提出されていません。できるだけ早く提出してください。'],
  [SETTING.MARK, '日付'],
  [SETTING.INCOMPLETE, '未提出'],
  [SETTING.LATE, '遅れとして記録'],
  [SETTING.TEACHER, ''],
];

const COLOR = { header: '#cfe2f3', ok: '#d9ead3', late: '#fff2cc', overdue: '#f4cccc', notYet: '#ffffff', none: '#eeeeee', excused: '#c9daf8', incomplete: '#fce5cd' };

// 欠席・配慮のシート。1行が1件。
//  欠席：日付（から〜まで）に休んだ。→ その期間にかかる課題は、休んだ日数だけその生徒の締切を延ばす（土日はとばす）
//  配慮：その課題は、遅れて出しても期限内あつかい（未提出でも赤にしない）
//  免除：その課題は、その生徒は対象外（提出率の分母から外す）
const EXCUSE_HEADERS = ['種類（欠席／配慮／免除）', '学籍番号', '氏名', '生徒ID', '日付（から）', '日付（まで）', '課題ID（配慮・免除）', 'メモ', '登録日時'];

// ───────── メニュー ─────────

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 課題提出')
    .addItem('🏠 ホーム（はじめはここ）', 'openHome')
    .addSeparator()
    .addItem('📷 読み取る', 'openScanner')
    .addItem('➕ 課題を登録する', 'openTaskForm')
    .addItem('📒 提出状況（教務手帳）を見る', 'showStatusSheet')
    .addSeparator()
    .addSubMenu(ui.createMenu('その他')
      .addItem('🤒 欠席・配慮・免除', 'openExcuseDialog')
      .addItem('👥 名簿を貼り付ける', 'openRosterDialog')
      .addItem('🖨 バーコードを印刷する', 'openBarcodePrint')
      .addItem('🖨 未提出者リスト・課題提出について', 'openListDialog')
      .addItem('📊 成績用の提出率', 'openReportDialog')
      .addItem('🔄 提出状況・集計を更新', 'refreshAll')
      .addSeparator()
      .addItem('バーコードをシートに作る（印刷画面が出ないとき）', 'makeBarcodeSheet')
      .addItem('お試しデータを入れる', 'insertSampleData')
      .addItem('シートを作り直す（消したシートを戻す）', 'setup')
      .addItem('配布用のひな形にする（データを全部消す）', 'makeTemplate'))
    .addToUi();
}

// ───────── ホーム（Home.html） ─────────

function openHome() {
  setup_(false);
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('Home').setTitle('課題提出バーコード'));
}

function openTaskForm() {
  setup_(false);
  SpreadsheetApp.getUi().showSidebar(HtmlService.createHtmlOutputFromFile('Task').setTitle('課題を登録する'));
}

function openRosterDialog() {
  setup_(false);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Roster').setWidth(900).setHeight(660), '名簿を貼り付ける');
}

function docProps_() {
  return PropertiesService.getDocumentProperties();
}

function userKey_(name) {
  return name + ':' + SpreadsheetApp.getActive().getId();
}

function getHomeData() {
  setup_(false);
  // 1人で使っていたころに「設定」シートに入れた担当の名前は、最初にホームを開いた先生の名前にする
  const props = PropertiesService.getUserProperties();
  const shared = String(getSettings_()[SETTING.TEACHER] || '').trim();
  if (props.getProperty(userKey_('teacher')) == null && shared) {
    props.setProperty(userKey_('teacher'), shared);
    putSetting_(SETTING.TEACHER, '');
  }
  const cfg = viewConfig_();
  const students = getStudents_().filter(function (s) { return !s.excluded; });
  const tasks = getTasks_();
  const logs = readLog_();
  const today = todayKey_();
  const mine = tasks.filter(function (t) { return !cfg.mine || t.teacher === cfg.teacher; });
  const myCls = myClasses_();
  const myStudents = myCls.length ? students.filter(function (s) { return myCls.indexOf(s.cls) >= 0; }) : students;
  const recent = mine.slice(-6).reverse().map(function (t) {
    // 担当クラスがあれば、担当クラスの数（その課題が担当クラスに出ていなければ全体の数）
    let st = taskStats_(t, myStudents, logs);
    const scoped = myCls.length && st.target > 0;
    if (!scoped) st = taskStats_(t, students, logs);
    return { id: t.id, label: taskLabelT_(t), submitted: st.submitted, target: st.target, overdue: st.overdue, missing: st.missing.length, scoped: !!scoped };
  });
  return {
    version: VERSION,
    students: students.length,
    noGakuseki: students.filter(function (s) { return !s.gakuseki; }).length,
    classes: classList_(students),
    tasks: tasks.length,
    logs: logs.length,
    todayCount: logs.filter(function (r) { return r.time instanceof Date && dateKey_(r.time) === today; }).length,
    printed: docProps_().getProperty('barcodePrinted') === '1',
    recent: recent,
    subjects: subjectList_(tasks),
    view: {
      teacher: cfg.teacher, personal: cfg.personal || props.getProperty(userKey_('personal')) === '1', subject: cfg.subject,
      cls: cfg.personal ? String(getSettings_()[SETTING.CLASS] || '') : cfg.cls, mine: cfg.mine, myClasses: myCls,
      mark: markMode_() === 'circle' ? '○' : '日付', incomplete: incompleteCounts_() ? '提出' : '未提出',
    },
  };
}

function subjectList_(tasks) {
  const seen = {};
  tasks.forEach(function (t) { if (t.subject) seen[t.subject] = true; });
  return Object.keys(seen);
}

function putSetting_(key, value) {
  const sh = sheet_(SHEET.SETTINGS);
  const vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === key) { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(value); return; }
  sh.appendRow([key, '']);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(value);
}

// ホームの「設定」
//  担当の名前・自分専用の教務手帳・その中のしぼりこみ → この先生だけ
//  ふつうの教務手帳のしぼりこみ・○か日付か・未完成の数え方 → 全員共通（「設定」シート）
function saveViewSettings(v) {
  const p = PropertiesService.getUserProperties();
  if (v.teacher !== undefined) p.setProperty(userKey_('teacher'), String(v.teacher || '').trim());
  if (v.myClasses) p.setProperty(userKey_('myClasses'), v.myClasses.map(normClass_).filter(String).join(','));
  const personal = !!v.personal && !!teacherName_();
  p.setProperty(userKey_('personal'), personal ? '1' : '0');
  if (personal) {
    p.setProperty(userKey_('pSubject'), String(v.subject || ''));
    // 自分専用の教務手帳のクラスは「担当クラス」を使う
    p.setProperty(userKey_('pMine'), v.mine ? '1' : '0');
  } else {
    putSetting_(SETTING.SUBJECT, v.subject || '');
    putSetting_(SETTING.CLASS, v.cls || '');
  }
  putSetting_(SETTING.MARK, v.mark === '○' ? '○' : '日付');
  if (v.incomplete) putSetting_(SETTING.INCOMPLETE, v.incomplete === '提出' ? '提出' : '未提出');
  INC_COUNTS_ = null;
  refreshAll();
  return getHomeData();
}

function showStatusSheet() {
  refreshAll();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(viewConfig_().statusName));
}

function startScanFor(taskId) {
  PropertiesService.getUserProperties().setProperty(userKey_('lastTask'), norm_(taskId));
  openScanner();
}

// ───────── 課題の登録（Task.html） ─────────

function getTaskFormData() {
  setup_(false);
  const students = getStudents_().filter(function (s) { return !s.excluded; });
  const props = PropertiesService.getUserProperties();
  return {
    classes: classList_(students),
    subjects: subjectList_(getTasks_()),
    lastSubject: props.getProperty(userKey_('lastSubject')) || '',
    lastClasses: props.getProperty(userKey_('lastClasses')) || myClasses_().join(','),
    // 締切後の扱い：前に選んだもの（はじめてのときは空 → 画面で「はじめに決めてください」）
    lastLate: props.getProperty(userKey_('lastLate')) || '',
    teacher: teacherName_(),
    defaultLate: latePolicy_(getSettings_()[SETTING.LATE]).label,
    today: todayKey_(),
  };
}

// 'yyyy-MM-dd' を、スプレッドシートのタイムゾーンでその日の 0 時にする（dateKey_ と同じタイムゾーン）
function toDate_(key) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(key || '')) ? Utilities.parseDate(key, tz_(), 'yyyy-MM-dd') : '';
}

// t: { subject, name, classes: [...], given: 'yyyy-MM-dd', due: 'yyyy-MM-dd', memo }
function addTask(t) {
  if (![t.subject, t.name, t.due, t.given, t.memo].some(function (v) { return String(v || '').trim() !== ''; }) && !(t.classes || []).length) {
    throw new Error('教科・課題名・締切日などを、どれか1つは入れてください');
  }
  const sh = sheet_(SHEET.TASKS);
  const classes = (t.classes || []).map(normClass_).filter(String);
  const late = latePolicy_(t.late).label;
  // 担当：画面で入れた名前（入れなければ、この先生の名前）。はじめて入れた名前は、この先生の名前として覚える
  const teacher = String(t.teacher !== undefined ? t.teacher : teacherName_()).trim();
  if (teacher && !teacherName_()) PropertiesService.getUserProperties().setProperty(userKey_('teacher'), teacher);
  sh.appendRow(['', String(t.subject || '').trim(), String(t.name || '').trim(), classes.join(','), toDate_(t.given), toDate_(t.due), late, String(t.memo || ''), teacher]);
  const row = sh.getLastRow();
  sh.getRange(row, 4).setNumberFormat('@').setValue(classes.join(','));
  sh.getRange(row, 5, 1, 2).setNumberFormat('m/d');
  const tasks = getTasks_();
  const task = tasks.filter(function (k) { return k.row === row; })[0];
  if (!task) throw new Error('課題を登録できませんでした。もう一度ためしてください');
  const props = PropertiesService.getUserProperties();
  props.setProperty(userKey_('lastSubject'), String(t.subject || '').trim());
  props.setProperty(userKey_('lastClasses'), classes.join(','));
  props.setProperty(userKey_('lastLate'), late);
  props.setProperty(userKey_('lastTask'), task.id);
  // 自分の教務手帳に、新しい課題の列をすぐ足す
  const view = statusView_(tasks, getStudents_());
  const logs = readLog_();
  const subs = {};
  view.tasks.forEach(function (k) { subs[k.id] = submissionsOf_(k, logs); });
  try { writeStatusSheet_(SpreadsheetApp.getActive(), view.tasks, view.students, subs, todayKey_(), view.cfg.statusName); } catch (e) { console.warn(e); }
  return { id: task.id, label: taskLabel_(task) };
}

// ───────── 名簿の貼り付け（Roster.html） ─────────
// Excel や校務システムからコピーした表を、そのまま貼れるようにする。
// rows: [{ gakuseki, name, kana, line }]（ページ側で1行ずつ読み取ったもの）

function rosterPlan_(rows) {
  const existing = getStudents_();
  const nameKey = function (v) { return String(v || '').normalize('NFKC').replace(/\s+/g, ''); };
  const pastedNames = {};
  rows.forEach(function (r) { pastedNames[nameKey(r.name)] = (pastedNames[nameKey(r.name)] || 0) + 1; });
  const byName = {};
  existing.forEach(function (s) { (byName[nameKey(s.name)] = byName[nameKey(s.name)] || []).push(s); });
  const used = {};
  return rows.map(function (r) {
    const gk = norm_(r.gakuseki);
    const nk = nameKey(r.name);
    let match = null, kind = 'new';
    // 1. 同じ名前の生徒が1人だけいれば、その人（進級して学籍番号が変わった場合）
    if (nk && byName[nk] && byName[nk].length === 1 && pastedNames[nk] === 1 && !used[byName[nk][0].row]) match = byName[nk][0];
    // 2. 同じ学籍番号の生徒（名前の字を直した場合）。ただしその人の名前が貼り付けた中にあるときは別人
    //    別人（前の学年の生徒が残っているところに新入生を貼った）を取りちがえないよう、名字（はじめの2文字）が同じときだけ
    if (!match && gk) {
      const hit = existing.filter(function (s) {
        return s.gakuseki === gk && !used[s.row] && !pastedNames[nameKey(s.name)] && nameKey(s.name).slice(0, 2) === nk.slice(0, 2);
      })[0];
      if (hit) match = hit;
    }
    if (match) {
      used[match.row] = true;
      const same = (!gk || match.gakuseki === gk) && nameKey(match.name) === nk && !match.excluded && (!r.kana || match.kana === String(r.kana).trim());
      kind = same ? 'same' : 'update';
    }
    return {
      gakuseki: gk, name: String(r.name || '').trim(), kana: String(r.kana || '').trim(), kind: kind,
      before: match ? (match.gakuseki || '') + ' ' + match.name : '', row: match ? match.row : 0,
    };
  });
}

function previewRoster(rows) {
  return rosterPlan_(rows);
}

function importRoster(rows) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const plan = rosterPlan_(rows);
    const sh = sheet_(SHEET.STUDENTS);
    let added = 0, updated = 0;
    plan.forEach(function (p) {
      if (p.kind !== 'update') return;
      // 学籍番号が変わったら、クラス・番号は空にして学籍番号から入れ直す
      const cur = sh.getRange(p.row, 2, 1, 6).getValues()[0];
      // 学籍番号のない名簿を貼ったときは、今の学籍番号・クラス・番号をそのまま残す
      const gkChanged = !!p.gakuseki && norm_(cur[0]) !== p.gakuseki;
      if (gkChanged) {
        sh.getRange(p.row, 2).setNumberFormat('@').setValue(p.gakuseki);
        sh.getRange(p.row, 3, 1, 2).setValues([['', '']]);
      }
      sh.getRange(p.row, 5).setValue(p.name);
      if (p.kana) sh.getRange(p.row, 6).setValue(p.kana);
      sh.getRange(p.row, 7).insertCheckboxes().uncheck();
      updated++;
    });
    const fresh = plan.filter(function (p) { return p.kind === 'new'; }).map(function (p) { return ['', p.gakuseki, '', '', p.name, p.kana, false]; });
    if (fresh.length) {
      const start = lastDataRow_(sh, 6) + 1;
      ensureSize_(sh, start + fresh.length, STUDENT_HEADERS.length);
      sh.getRange(start, 1, fresh.length, 2).setNumberFormat('@');
      sh.getRange(start, 1, fresh.length, STUDENT_HEADERS.length).setValues(fresh);
      sh.getRange(start, 7, fresh.length, 1).insertCheckboxes();
      added = fresh.length;
    }
    SpreadsheetApp.flush();
    getStudents_(); // 生徒ID・クラス・番号をここで付ける
    return { added: added, updated: updated, same: plan.length - added - updated };
  } finally {
    lock.releaseLock();
  }
}

// ───────── 欠席・配慮・免除の登録（Excuse.html） ─────────

function openExcuseDialog() {
  setup_(false);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutputFromFile('Excuse').setWidth(820).setHeight(660), '欠席・配慮・免除');
}

function getExcuseData() {
  setup_(false);
  const students = sortStudents_(getStudents_().filter(function (s) { return !s.excluded; }));
  const tasks = getTasks_();
  const taskMap = {};
  tasks.forEach(function (t) { taskMap[t.id] = t; });
  const byId = {};
  students.forEach(function (s) { byId[s.id] = s; });
  const sh = sheet_(SHEET.EXCUSE);
  const entries = [];
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, EXCUSE_HEADERS.length).getValues().forEach(function (r, i) {
      if (String(r[0]).trim() === '') return;
      const st = byId[norm_(r[3])];
      const from = dateKey_(r[4]), to = dateKey_(r[5]);
      const t = taskMap[norm_(r[6])];
      entries.push({
        row: i + 2, type: String(r[0]).trim(), sid: norm_(r[3]),
        who: st ? whoOf_(st) : (String(r[1]) + ' ' + String(r[2])).trim(),
        period: from ? shortDate_(from) + (to && to !== from ? '〜' + shortDate_(to) : '') : '',
        task: t ? taskLabel_(t) : (r[6] ? String(r[6]) : ''),
        memo: String(r[7]),
      });
    });
  }
  return {
    today: todayKey_(),
    students: students.map(function (s) { return { id: s.id, gakuseki: s.gakuseki, name: s.name, cls: s.cls, no: s.no }; }),
    tasks: tasks.slice().reverse().map(function (t) { return { id: t.id, label: taskLabelT_(t) }; }),
    entries: entries.reverse(),
  };
}

// o: { type: '欠席' | '配慮' | '免除', ids: [生徒ID...], from, to, taskId, memo }
function addExcuses(o) {
  const type = o.type === '配慮' || o.type === '免除' ? o.type : '欠席';
  if (type !== '欠席' && !o.taskId) throw new Error('課題を選んでください');
  if (type === '欠席' && !o.from) throw new Error('休んだ日を入れてください');
  if (type === '欠席' && o.to && o.to < o.from) { const tmp = o.from; o.from = o.to; o.to = tmp; }
  const students = getStudents_();
  const byId = {};
  students.forEach(function (s) { byId[s.id] = s; });
  const now = new Date();
  const rows = (o.ids || []).map(norm_).filter(function (id) { return byId[id]; }).map(function (id) {
    const s = byId[id];
    return [type, s.gakuseki, s.name, s.id, type === '欠席' ? toDate_(o.from) : '', type === '欠席' ? toDate_(o.to || o.from) : '',
      type === '欠席' ? '' : norm_(o.taskId), String(o.memo || ''), now];
  });
  if (!rows.length) throw new Error('生徒を選んでください');
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(SHEET.EXCUSE);
    const start = sh.getLastRow() + 1;
    ensureSize_(sh, start + rows.length, EXCUSE_HEADERS.length);
    sh.getRange(start, 2, rows.length, 1).setNumberFormat('@');
    sh.getRange(start, 4, rows.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, rows.length, EXCUSE_HEADERS.length).setValues(rows);
    sh.getRange(start, 5, rows.length, 2).setNumberFormat('m/d');
    sh.getRange(start, 9, rows.length, 1).setNumberFormat('m/d hh:mm');
  } finally {
    lock.releaseLock();
  }
  EXC_ = null;
  refreshAll();
  return getExcuseData();
}

function deleteExcuse(row, sid) {
  const sh = sheet_(SHEET.EXCUSE);
  if (row < 2 || row > sh.getLastRow() || norm_(sh.getRange(row, 4).getValue()) !== norm_(sid)) {
    throw new Error('シートが変わっています。画面を開き直してください。');
  }
  sh.deleteRow(row);
  EXC_ = null;
  refreshAll();
  return getExcuseData();
}

// 読み取り画面で、遅れて出した生徒を「配慮（期限内あつかい）」にする
function excuseLate(taskId, studentId) {
  const tasks = getTasks_();
  const task = findTask_(tasks, taskId);
  const st = getStudents_().filter(function (x) { return x.id === norm_(studentId); })[0];
  if (!task || !st) return { kind: 'error', message: '見つかりませんでした' };
  addExcuseRow_(['配慮', st.gakuseki, st.name, st.id, '', '', task.id, '遅れて提出（欠席など）：読み取り画面から', new Date()]);
  EXC_ = null;
  const logs = readLog_();
  const warn = updateStatusRow_(st, tasks, logs);
  return { kind: 'undone', message: '配慮しました：' + whoOf_(st) + ' → 期限内あつかい', warn: warn, stats: taskStats_(task, getStudents_(), logs) };
}

function addExcuseRow_(r) {
  const sh = sheet_(SHEET.EXCUSE);
  sh.appendRow(r);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(r[1]);
  sh.getRange(sh.getLastRow(), 4).setNumberFormat('@').setValue(r[3]);
}

function showSummarySheet() {
  refreshAll();
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(viewConfig_().summaryName));
}

// ───────── 配布用のひな形 ─────────

function makeTemplate() {
  const ui = SpreadsheetApp.getUi();
  const ok = ui.alert('配布用のひな形にします',
    '生徒名簿・課題一覧・提出記録・欠席・配慮の中身をすべて消します（シートと設定は残ります）。\n元に戻せないので、必ずこのスプレッドシートのコピーで行ってください。\n\n続けますか？', ui.ButtonSet.OK_CANCEL);
  if (ok !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  [SHEET.STUDENTS, SHEET.TASKS, SHEET.LOG, SHEET.EXCUSE].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (!sh || sh.getMaxRows() < 2) return;
    const body = sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns());
    if (n === SHEET.STUDENTS) sh.getRange(2, 7, sh.getMaxRows() - 1, 1).removeCheckboxes();
    body.clearContent();
  });
  [SHEET.STATUS, SHEET.SUMMARY].forEach(function (n) { const sh = ss.getSheetByName(n); if (sh) sh.clear(); });
  ss.getSheets().forEach(function (sh) { // 先生ごとの教務手帳・集計は消す
    const n = sh.getName();
    if ((n.indexOf(SHEET.STATUS + '（') === 0 || n.indexOf(SHEET.SUMMARY + '（') === 0) && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  [SHEET.PRINT_MISSING, SHEET.PRINT_NOTICE, SHEET.PRINT_BARCODE].forEach(function (n) { const sh = ss.getSheetByName(n); if (sh) ss.deleteSheet(sh); });
  docProps_().deleteProperty('barcodePrinted');
  ss.setActiveSheet(ss.getSheetByName(SHEET.HOWTO));
  ui.alert('ひな形になりました。\n\n共有リンクの「/edit…」を「/copy」に変えて配ると、開いた先生が自分用のコピーを作れます。');
}

// 課題一覧に課題名を書いたら、すぐに課題IDを付ける
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() === SHEET.TASKS && e.range.getRow() > 1) getTasks_();
    else if (isStatusSheet_(sh.getName()) && e.range.getLastRow() >= STATUS_TOP) statusEdit_(e.range);
  } catch (err) {
    // 単純トリガーでは失敗しても何もしない（読み取り時や集計時にも ID は付く）
    console.warn('onEdit：' + (err && err.stack || err));
  }
}

// 教務手帳のマスに直接書いたとき（バーコードを読まずに、手で記録する）
//  ○・日付・そのほかの文字 → 提出（日付を書いたらその日の提出）／△ → 未完成／消す → 記録を取り消す
function statusEdit_(range) {
  const sh = range.getSheet();
  const lastCol = sh.getLastColumn();
  if (lastCol <= STATUS_FIXED) return;
  const ids = sh.getRange(1, STATUS_FIXED + 1, 1, lastCol - STATUS_FIXED).getValues()[0].map(norm_);
  const nTasks = ids.indexOf('') < 0 ? ids.length : ids.indexOf('');
  const r0 = Math.max(range.getRow(), STATUS_TOP);
  const r1 = range.getLastRow();
  const c0 = Math.max(range.getColumn(), STATUS_FIXED + 1);
  const c1 = Math.min(range.getLastColumn(), STATUS_FIXED + nTasks);
  if (r0 > r1 || c0 > c1) return; // 名前や合計の列は、書いても次の作り直しで元に戻る
  const ss = SpreadsheetApp.getActive();
  if ((r1 - r0 + 1) * (c1 - c0 + 1) > 40) {
    // まちがえて広い範囲を消した・貼ったときは、記録は変えずに表を元に戻す
    ss.toast('一度に書きかえられるのは40マスまでです。記録は変えずに、表を元に戻しました。', '教務手帳', 8);
    refreshAll();
    return;
  }
  const sids = sh.getRange(r0, 1, r1 - r0 + 1, 1).getValues().map(function (r) { return norm_(r[0]); });
  const vals = sh.getRange(r0, c0, r1 - r0 + 1, c1 - c0 + 1).getValues();
  const touched = {};
  vals.forEach(function (row, i) {
    if (!sids[i]) return;
    row.forEach(function (v, j) {
      const taskId = ids[c0 - STATUS_FIXED - 1 + j];
      const text = String(v).trim();
      if (text === '－' || text === '免') return; // 対象外・免除のマスはそのまま
      touched[sids[i]] = true;
      if (text === '') recordScan(taskId, sids[i], 'undo');
      else if (/△|未完/.test(text)) recordScan(taskId, sids[i], 'incomplete', v instanceof Date ? v : null);
      else recordScan(taskId, sids[i], 'submit', v instanceof Date ? v : null);
    });
  });
  // すでに記録があった・対象外だったなどで変わらなかったマスも、正しい表示（日付・色）に戻す
  const tasks = getTasks_();
  const logs = readLog_();
  getStudents_().forEach(function (st) { if (touched[st.id]) updateStatusRow_(st, tasks, logs); });
  ss.toast('教務手帳に書いた内容を、提出記録に入れました。', '教務手帳', 4);
}

function openScanner() {
  setup_(false);
  const html = HtmlService.createHtmlOutputFromFile('Scan').setTitle('課題の読み取り');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openBarcodePrint() {
  setup_(false);
  docProps_().setProperty('barcodePrinted', '1');
  const students = sortStudents_(getStudents_().filter(function (s) { return !s.excluded && s.id; }));
  const tasks = getTasks_().reverse();
  const t = HtmlService.createTemplateFromFile('Print');
  t.payload = toScriptJson_({
    title: getSettings_()[SETTING.TITLE] || '',
    students: students.map(function (s) { return { id: s.id, gakuseki: s.gakuseki, cls: s.cls, no: s.no, name: s.name }; }),
    tasks: tasks.map(function (k) { return { id: k.id, subject: k.subject, name: displayName_(k), due: k.due }; }),
  });
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(1000).setHeight(720), 'バーコードの印刷');
}

// ───────── 成績用の提出率（Report.html） ─────────

function openReportDialog() {
  setup_(false);
  const tasks = getTasks_();
  const t = HtmlService.createTemplateFromFile('Report');
  t.payload = toScriptJson_({
    today: todayKey_(),
    subjects: subjectList_(tasks).concat(tasks.some(function (k) { return !k.subject; }) ? ['（教科なし）'] : []),
    classes: classList_(getStudents_()),
    me: teacherName_(),
    myClasses: myClasses_(),
    shared: tasks.some(function (k) { return k.teacher && k.teacher !== teacherName_(); }),
  });
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(720).setHeight(640), '成績用の提出率');
}

// opts: { from, to: 'yyyy-MM-dd', subjects: [...]（空＝すべて）, classes: [...], mine: bool, split: bool }
//  数える課題：締切日（なければ出した日）が期間内で、締切が過ぎたもの。免除は数えない。
//  期限内：遅れでない提出（欠席の延長・配慮・猶予の日数をふくむ）。未完成は設定の数え方に合わせる。
function makeGradeReport(opts) {
  const ss = SpreadsheetApp.getActive();
  const today = todayKey_();
  const from = opts.from || '0000-01-01';
  const to = opts.to || '9999-12-31';
  if (from > to) throw new Error('期間の「から」と「まで」が逆です');
  const me = teacherName_();
  const subjSel = opts.subjects || [];
  const subjOf = function (t) { return t.subject || '（教科なし）'; };
  const tasks = getTasks_().filter(function (t) {
    const key = t.due || t.given;
    if (!key || key < from || key > to) return false;
    if (t.due && today <= t.due) return false; // まだ締切が来ていない
    if (subjSel.length && subjSel.indexOf(subjOf(t)) < 0) return false;
    if (opts.mine && t.teacher !== me) return false;
    return true;
  });
  if (!tasks.length) throw new Error('この期間・条件で、締切が過ぎた課題はありません');
  const classes = (opts.classes || []).map(normClass_);
  const students = sortStudents_(getStudents_().filter(function (s) { return !s.excluded && (!classes.length || classes.indexOf(s.cls) >= 0); }));
  const logs = readLog_();
  const subs = {};
  tasks.forEach(function (t) { subs[t.id] = submissionsOf_(t, logs); });

  const subjects = [];
  tasks.forEach(function (t) { if (subjects.indexOf(subjOf(t)) < 0) subjects.push(subjOf(t)); });
  const groups = opts.split && subjects.length > 1 ? subjects.concat(['合計']) : [subjSel.length ? subjects.join('・') : 'すべての教科'];
  const inGroup = function (g, t) { return g === '合計' || g === 'すべての教科' || !opts.split || subjects.length < 2 ? true : subjOf(t) === g; };

  const rows = students.map(function (s) {
    const row = [s.gakuseki, s.cls, s.no, s.name];
    const missingNames = [];
    groups.forEach(function (g) {
      let n = 0, done = 0, onTime = 0;
      tasks.forEach(function (t) {
        if (!inGroup(g, t) || !isTarget_(t, s)) return;
        n++;
        const sub = subs[t.id][s.id];
        if (counts_(sub)) { done++; if (!sub.late && !sub.incomplete) onTime++; }
        else if (g === groups[groups.length - 1]) missingNames.push((t.subject ? t.subject + ' ' : '') + displayName_(t) + (sub ? '（未完成）' : ''));
      });
      row.push(n, done, onTime, n ? done / n : '', n ? onTime / n : '');
    });
    row.push(missingNames.join('、'));
    return row;
  });

  const label = me ? '（' + me.replace(/[\[\]\/\\?*:'']/g, '') + '）' : '';
  const sh = freshSheet_(ss, '成績用_提出率' + label);
  const width = 4 + groups.length * 5 + 1;
  ensureSize_(sh, rows.length + 6, width);
  const period = (opts.from ? shortDate_(from) : '') + '〜' + (opts.to ? shortDate_(to) : '');
  sh.getRange(1, 1).setValue('提出率（成績用）　' + period + (opts.mine && me ? '　担当：' + me + 'の課題' : '') + (me ? '　作成：' + me : '') + '　（' + Utilities.formatDate(new Date(), tz_(), 'M/d HH:mm') + ' 作成）').setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1).setValue('数えた課題：締切日がこの期間にあり、締切が過ぎた ' + tasks.length + ' 件（免除は除く）。期限内には欠席・配慮・猶予の日数のうちの提出をふくむ。未完成は「' + (incompleteCounts_() ? '提出として数える' : '未提出あつかい') + '」。').setFontColor('#555555').setFontSize(9);
  const head1 = ['', '', '', ''], head2 = ['学籍番号', 'クラス', '番号', '氏名'];
  groups.forEach(function (g) { head1.push(g, '', '', '', ''); head2.push('課題数', '提出', '期限内', '提出率', '期限内\n提出率'); });
  head1.push(''); head2.push('未提出の課題');
  sh.getRange(3, 1, 2, width).setValues([head1, head2]).setFontWeight('bold').setBackground(COLOR.header).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  groups.forEach(function (g, i) { sh.getRange(3, 5 + i * 5, 1, 5).merge(); });
  if (rows.length) {
    sh.getRange(5, 1, rows.length, 2).setNumberFormat('@'); // 学籍番号・クラス（「1-1」が日付にならないように）
    sh.getRange(5, 1, rows.length, width).setValues(rows).setVerticalAlignment('middle');
    groups.forEach(function (g, i) {
      sh.getRange(5, 5 + i * 5, rows.length, 3).setNumberFormat('0').setHorizontalAlignment('center');
      sh.getRange(5, 8 + i * 5, rows.length, 2).setNumberFormat('0%').setHorizontalAlignment('center');
    });
    sh.getRange(5, width, rows.length, 1).setWrap(true).setFontSize(9);
    sh.getRange(3, 1, rows.length + 2, width).setBorder(true, true, true, true, true, true, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
    for (let i = 1; i < students.length; i++) {
      if (students[i].cls !== students[i - 1].cls) sh.getRange(5 + i, 1, 1, width).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    }
  }
  [70, 50, 40, 110].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  for (let c = 5; c < width; c++) sh.setColumnWidth(c, 52);
  sh.setColumnWidth(width, 260);
  sh.setFrozenRows(4);
  sh.setFrozenColumns(4);
  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();
  return { sheetName: sh.getName(), tasks: tasks.length, students: rows.length, pdfUrl: pdfUrl_(ss, sh).replace('portrait=true', 'portrait=' + (groups.length > 2 ? 'false' : 'true')) };
}

function openListDialog() {
  setup_(false);
  const tasks = getTasks_().reverse();
  const t = HtmlService.createTemplateFromFile('Select');
  t.payload = toScriptJson_({
    today: todayKey_(),
    tasks: tasks.map(function (k) { return { id: k.id, subject: k.subject, name: displayName_(k), due: k.due, classes: k.classes.join(','), teacher: k.teacher }; }),
    classes: classList_(getStudents_()),
    me: teacherName_(),
    myClasses: myClasses_(),
  });
  SpreadsheetApp.getUi().showModalDialog(t.evaluate().setWidth(720).setHeight(640), '未提出者リスト・課題提出について');
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
  if (howto.sh.getLastRow() > 0 && howto.sh.createTextFinder('業務手帳').findNext()) {
    howto.sh.clear(); // 古い言い方の「使い方」は書き直す
    writeHowTo_(howto.sh);
  }
  if (howto.sh.getLastRow() === 0) {
    writeHowTo_(howto.sh);
    ss.setActiveSheet(howto.sh);
    ss.moveActiveSheet(1);
  }

  const settings = make(SHEET.SETTINGS, ['項目', '値'], [420, 460]);
  if (settings.fresh) {
    settings.sh.getRange(2, 1, SETTING_DEFAULTS.length, 2).setValues(SETTING_DEFAULTS);
  } else {
    // 名前を変えた設定項目は、新しい名前に書きかえる（中身はそのまま）
    const oldNotice = settings.sh.createTextFinder('「提出のお願い」に書く文').matchEntireCell(true).findNext();
    if (oldNotice) oldNotice.setValue(SETTING.NOTICE);
    const oldPrint = ss.getSheetByName('印刷_提出のお願い');
    if (oldPrint && ss.getSheets().length > 1) ss.deleteSheet(oldPrint);
    // あとから増えた設定項目を書き足す
    const have = settings.sh.getLastRow() > 1 ? settings.sh.getRange(2, 1, settings.sh.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
    SETTING_DEFAULTS.forEach(function (r) { if (have.indexOf(r[0]) < 0) settings.sh.appendRow(r); });
  }

  const students = make(SHEET.STUDENTS, STUDENT_HEADERS, [170, 80, 70, 50, 140, 160, 120]);
  if (!students.fresh && String(students.sh.getRange(1, 2).getValue()) === 'クラス') {
    // 旧版の名簿（学籍番号の列なし）には列を差しこむ
    students.sh.insertColumnAfter(1);
    students.sh.getRange(1, 1, 1, STUDENT_HEADERS.length).setValues([STUDENT_HEADERS]);
    created = true;
  }
  if (students.fresh) {
    students.sh.getRange('A2:B').setNumberFormat('@'); // 先頭の 0 が消えないように文字として扱う
    students.sh.getRange('B1').setNote('4けたの学籍番号（例：1101＝1年1組1番）。クラス・番号が空欄なら、ここから自動で入ります。');
  }

  if (students.sh.getRange('C2').getNumberFormat() !== '@') students.sh.getRange('C2:C').setNumberFormat('@'); // 「1-1」が日付にならないように

  const tasks = make(SHEET.TASKS, TASK_HEADERS, [70, 80, 220, 160, 90, 90, 130, 160, 80, 50, 50, 60, 60, 60]);
  if (!tasks.fresh && String(tasks.sh.getRange(1, 9).getValue()) === '対象') {
    // 旧版の課題一覧（「担当」の列なし）には列を差しこむ
    tasks.sh.insertColumnBefore(9);
    tasks.sh.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]);
    tasks.sh.getRange('I2:I').clearDataValidations().setNumberFormat('@');
    // それまでの課題は、更新後に最初に開いた先生（ふつうはこのファイルを作った先生）の課題にする
    const me = teacherName_();
    const lastT = tasks.sh.getLastRow();
    if (me && lastT > 1) {
      const rows = tasks.sh.getRange(2, 1, lastT - 1, 1).getValues();
      tasks.sh.getRange(2, 9, lastT - 1, 1).setValues(rows.map(function (r) { return [r[0] !== '' ? me : '']; }));
    }
  }
  if (!tasks.fresh && String(tasks.sh.getRange(1, 7).getValue()) === 'メモ') {
    // 旧版の課題一覧（「締切後の扱い」の列なし）には列を差しこむ
    tasks.sh.insertColumnBefore(7);
    tasks.sh.getRange(1, 1, 1, TASK_HEADERS.length).setValues([TASK_HEADERS]);
    tasks.sh.getRange('G2:G').clearDataValidations().setNumberFormat('@'); // となりの日付の列の決まりを引きつがない
  }
  if (tasks.sh.getRange('D2').getNumberFormat() !== '@') tasks.sh.getRange('D2:D').setNumberFormat('@'); // 対象クラス「1-1」が日付にならないように
  const gRule = tasks.sh.getRange('G2').getDataValidation();
  if (tasks.fresh || !gRule || gRule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    tasks.sh.getRange('G2:G').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(LATE_CHOICES, true).setAllowInvalid(true).build());
  }
  if (tasks.fresh) {
    const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText('日付を入れてください（例：2026/10/5）').build();
    tasks.sh.getRange('E2:F').setDataValidation(dateRule).setNumberFormat('m/d');
    tasks.sh.getRange('N2:N').setNumberFormat('0%');
    tasks.sh.getRange('J1:N1').setBackground('#e0e0e0');
    tasks.sh.getRange('J1').setNote('J〜N列は「提出状況・集計を更新」で自動で書き込まれます。');
    tasks.sh.getRange('I1').setNote('課題を登録した先生の名前（ホームの設定の「担当の先生の名前」）。共有して使うとき、自分の課題だけを表示するのに使います。');
    tasks.sh.getRange('G1').setNote('締切を過ぎて出したときの扱い。空欄なら「設定」シートの値。\n・遅れとして記録\n・◯日まで期限内（欠席などへの配慮。土日はとばして数える）\n・区別しない');
  }

  const log = make(SHEET.LOG, LOG_HEADERS, [150, 60, 80, 200, 110, 70, 60, 50, 130, 70]);
  if (!log.fresh && String(log.sh.getRange(1, 6).getValue()) === 'クラス') {
    // 旧版の提出記録（学籍番号の列なし）には列を差しこむ
    log.sh.insertColumnAfter(5);
    log.sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
    log.sh.getRange('F2:F').setNumberFormat('@');
  }
  if (log.fresh) {
    log.sh.getRange('A2:A').setNumberFormat('m/d hh:mm');
    log.sh.getRange('E2:F').setNumberFormat('@');
  }

  // 日付は「月/日」で表示する（年は出さない）。前の版のシートもここで切りかえる
  if (tasks.sh.getRange('E2').getNumberFormat() !== 'm/d') tasks.sh.getRange('E2:F').setNumberFormat('m/d');
  if (log.sh.getRange('A2').getNumberFormat() !== 'm/d hh:mm') log.sh.getRange('A2:A').setNumberFormat('m/d hh:mm');

  const exc = make(SHEET.EXCUSE, EXCUSE_HEADERS, [130, 70, 120, 80, 90, 90, 90, 200, 130]);
  if (exc.fresh) {
    exc.sh.getRange('B2:B').setNumberFormat('@');
    exc.sh.getRange('D2:D').setNumberFormat('@');
    exc.sh.getRange('E2:F').setNumberFormat('m/d');
    exc.sh.getRange('A1').setNote('欠席：日付（から〜まで）に休んだ。その期間にかかる課題は、休んだ日数だけこの生徒の締切が延びる（土日はとばす）\n配慮：その課題は遅れても期限内あつかい\n免除：その課題はこの生徒は対象外');
  }

  if (exc.sh.getRange('E2').getNumberFormat() !== 'm/d') {
    exc.sh.getRange('E2:F').setNumberFormat('m/d');
    exc.sh.getRange('I2:I').setNumberFormat('m/d hh:mm');
  }

  make(SHEET.STATUS);
  make(SHEET.SUMMARY);

  // 新しいスプレッドシートに最初からある空の「シート1」は消す
  ['シート1', 'Sheet1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  if (showMessage) {
    ss.setActiveSheet(ss.getSheetByName(SHEET.HOWTO));
    SpreadsheetApp.getUi().alert(created ? 'シートを作りました。メニュー「📋 課題提出」→「🏠 ホーム」から始めてください。' : 'シートはすでにそろっています。');
  }
}

function writeHowTo_(sh) {
  const lines = [
    ['課題提出バーコード管理'],
    ['👉 上のメニュー「📋 課題提出」→「🏠 ホーム」を押してください。右側に出る「ホーム」から全部の操作ができます。'],
    ['　（はじめて押したときは Google の確認画面が出ます。自分のアカウントを選び「許可」を押したあと、もう一度メニューを押してください。「このアプリは確認されていません」と出たら「詳細」→「移動」→「許可」）'],
    [''],
    ['■ はじめの準備（1回だけ）'],
    ['1. 名簿を貼り付ける：Excel や校務システムの名簿をコピーして、ホームの「名簿の貼り付け」に貼るだけ。学籍番号・氏名を自動で見つけます。'],
    ['2. バーコードを印刷する：A4 に1人1本の短冊（左に学籍番号と名前、右にバーコード3つ）。横の点線で裁断機を1回下ろすと1人分。生徒は右から1つずつ切って、ノートなどにはる（シール用紙OK）。'],
    [''],
    ['■ ふだん'],
    ['1. 課題を登録する：ホームの「➕ 課題を登録する」。教科・課題名は空欄でもOK。締切は「明日」「1週間後」などのボタンで。'],
    ['2. 読み取る：「登録して、読み取りへ」を押すか、ホームの課題の「📷 読む」。読み取り欄をクリックしてからピッ・ピッ。提出物の順番はばらばらでOK。'],
    ['3. 見る・配る：「📒 提出状況（教務手帳）」で一覧。「📝 未提出者・課題提出について」で印刷。'],
    ['   遅れて1人だけ出しに来たときは、教務手帳のマスに直接「○」や日付（10/6）を書いてもよい（△＝未完成、消す＝取り消し）。自動で提出記録に入ります。'],
    [''],
    ['■ 遅れ・欠席への配慮'],
    ['・課題を登録するときに「締切を過ぎて出したら？」を1つ選ぶだけ（次からは同じ選び方が最初から入る）：遅れとして記録／少しの遅れは期限内（1〜7日。欠席などへの配慮。1人ずつ登録しなくてよい）／区別しない'],
    ['・遅れて出した提出も「提出」として数え、提出率と期限内提出率の両方を出します。'],
    ['・必要な先生だけ：ホームの下の「1人ずつの欠席・配慮・免除」、読み取り画面の「△ 未完成」モード。使わなくてもかまいません。'],
    [''],
    ['■ 進級したら'],
    ['新しい名簿をそのまま「名簿の貼り付け」に貼るだけ。同じ名前の生徒は学籍番号が新しくなり、バーコードは刷り直さずに使えます。'],
    [''],
    ['■ シートについて（ふつうはさわらなくてOK）'],
    ['生徒名簿・課題一覧・提出記録：データのもと（手で直してもよい）　提出状況：教務手帳の形の一覧（自動）　集計：提出率（自動）　設定：表示の設定'],
    ['生徒名簿の「生徒ID」はバーコードの中身です。変えないでください。'],
  ];
  sh.getRange(1, 1, lines.length, 1).setValues(lines).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 900);
  sh.getRange(1, 1).setFontSize(18).setFontWeight('bold');
  sh.getRange(2, 1).setFontSize(14).setFontWeight('bold').setBackground('#e8f0fe').setFontColor('#1967d2');
  lines.forEach(function (l, i) { if (/^■/.test(l[0])) sh.getRange(i + 1, 1).setFontWeight('bold').setBackground('#eef3fb'); });
}

function insertSampleData() {
  setup_(false);
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('お試しデータ（架空の生徒6人・課題3つ）を入れます。あとで行ごと消してください。よろしいですか？', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName(SHEET.STUDENTS);
  const sample = [
    ['', '1101', '', '', '青木 さくら', 'あおき さくら', false],
    ['', '1102', '', '', '石田 はると', 'いしだ はると', false],
    ['', '1103', '', '', '上野 ゆい', 'うえの ゆい', false],
    ['', '1201', '', '', '江藤 そうた', 'えとう そうた', false],
    ['', '1202', '', '', '小川 めい', 'おがわ めい', false],
    ['', '1203', '', '', '加藤 れん', 'かとう れん', false],
  ];
  const sStart = lastDataRow_(st, 6) + 1;
  ensureSize_(st, sStart + sample.length, STUDENT_HEADERS.length);
  st.getRange(sStart, 1, sample.length, 2).setNumberFormat('@');
  st.getRange(sStart, 1, sample.length, sample[0].length).setValues(sample);
  st.getRange(sStart, 7, sample.length, 1).insertCheckboxes();
  const now = new Date();
  function day(n) { const d = new Date(now); d.setDate(d.getDate() + n); return d; }
  const tk = ss.getSheetByName(SHEET.TASKS);
  const me = teacherName_();
  tk.getRange(tk.getLastRow() + 1, 1, 3, 9).setValues([
    ['', '国語', '漢字ノート 第3回', '', day(-10), day(-3), '遅れとして記録', '', me],
    ['', '数学', 'ワーク p.20〜25', '1-1', day(-5), day(-1), '3日まで期限内', '', me],
    ['', '英語', 'Unit 2 ワークシート', '1-1,1-2', day(-1), day(5), '遅れとして記録', '', me],
  ]);
  getStudents_();
  getTasks_();
  ss.setActiveSheet(st);
  ui.alert('お試しデータを入れました。ホームの課題の「📷 読む」を押し、読み取り欄に学籍番号（例：1101）を打って Enter を押すと試せます。');
}

// ───────── データの読み書き ─────────

function sheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('「' + name + '」シートがありません。メニュー「📋 課題提出」→「その他」→「シートを作り直す」を押してください。');
  return sh;
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
    const name = String(r[4]).trim();
    let id = norm_(r[0]);
    if (!name && !id) return;
    if (!id && name) {
      maxNo += 1;
      id = String(maxNo);
      sh.getRange(i + 2, 1).setNumberFormat('@').setValue(id);
    }
    const gakuseki = norm_(r[1]);
    let cls = normClass_(r[2]);
    let no = r[3];
    if (r[2] instanceof Date) sh.getRange(i + 2, 3).setNumberFormat('@').setValue(cls); // 日付になってしまったクラスを文字に戻す
    // 4けたの学籍番号（1101＝1年1組1番）から、空欄のクラス・番号を入れる
    const m = gakuseki.match(/^(\d)(\d)(\d\d)$/);
    if (m && (cls === '' || no === '')) {
      if (cls === '') cls = m[1] + '-' + m[2];
      if (no === '') no = Number(m[3]);
      sh.getRange(i + 2, 3).setNumberFormat('@').setValue(cls);
      sh.getRange(i + 2, 4).setValue(no);
    }
    const ex = r[6];
    list.push({
      id: id, gakuseki: gakuseki, cls: cls, no: no, name: name, kana: String(r[5]).trim(),
      excluded: ex === true || /^(✓|✔|○|〇|1|TRUE|除外|転出|×)$/i.test(String(ex).trim()),
      row: i + 2,
    });
  });
  return list;
}

// 画面や印刷に出す呼び名（学籍番号があれば「1101 青木 さくら」）
function whoOf_(s) {
  return (s.gakuseki ? s.gakuseki + ' ' : s.cls + ' ' + s.no + '番 ') + s.name;
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
  const vals = sh.getRange(2, 1, last - 1, 9).getValues();
  const defaultLate = getSettings_()[SETTING.LATE];
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
    // 教科・課題名は空欄でもよい（締切日や対象クラスだけでも1つの課題として扱う）
    const filled = [r[1], r[2], r[3], r[4], r[5], r[7]].some(function (v) { return String(v).trim() !== ''; });
    if (!filled && !id) return;
    if (!id || seen[id]) { // 空欄、または行をコピーして ID が重なったときは新しい ID
      maxNo += 1;
      id = 'K' + ('00' + maxNo).slice(-3);
      sh.getRange(i + 2, 1).setValue(id);
    }
    seen[id] = true;
    if (r[3] instanceof Date) sh.getRange(i + 2, 4).setNumberFormat('@').setValue(normClass_(r[3])); // 日付になってしまった対象クラスを文字に戻す
    tasks.push({
      id: id, subject: String(r[1]).trim(), name: name,
      classes: splitList_(r[3]), given: dateKey_(r[4]), due: dateKey_(r[5]), late: latePolicy_(String(r[6]).trim() || defaultLate),
      memo: String(r[7]), teacher: String(r[8]).trim(), row: i + 2,
    });
  });
  return tasks;
}

function taskLabel_(t) {
  return (t.subject ? '【' + t.subject + '】' : '') + displayName_(t) + (t.due ? '（締切 ' + shortDate_(t.due) + '）' : '');
}

// 共有して使うとき、だれの課題か分かるように［担当］を付ける
function taskLabelT_(t) {
  return taskLabel_(t) + (t.teacher ? '［' + t.teacher + '］' : '');
}

// 課題名が空欄のときの呼び名（例：「10/2の課題」「課題K005」）
function displayName_(t) {
  if (t.name) return t.name;
  if (t.given) return shortDate_(t.given) + 'に出した課題';
  return '課題' + t.id;
}

function isTarget_(task, student) {
  if (student.excluded) return false;
  if (isExempt_(task, student.id)) return false;
  return !task.classes.length || task.classes.indexOf(student.cls) >= 0;
}

// 学級としては対象だが、免除にした生徒
function isExemptTarget_(task, student) {
  return !student.excluded && isExempt_(task, student.id) && (!task.classes.length || task.classes.indexOf(student.cls) >= 0);
}

// ───────── 欠席・配慮・免除 ─────────

let EXC_ = null; // 1回の実行の中だけ覚えておく

function excuses_() {
  if (EXC_) return EXC_;
  const out = { absent: {}, task: {} };
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.EXCUSE);
  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, EXCUSE_HEADERS.length).getValues();
    let byGk = null;
    vals.forEach(function (r) {
      const type = String(r[0]).normalize('NFKC').trim();
      let sid = norm_(r[3]);
      if (!sid && r[1] !== '') { // 手で学籍番号だけ書いた行
        if (!byGk) { byGk = {}; getStudents_().forEach(function (st) { if (st.gakuseki && !st.excluded) byGk[st.gakuseki] = st.id; }); }
        sid = byGk[norm_(r[1])] || '';
      }
      if (!sid) return;
      if (/欠/.test(type)) {
        const from = dateKey_(r[4]);
        const to = dateKey_(r[5]) || from;
        if (!from) return;
        const days = out.absent[sid] || (out.absent[sid] = {});
        for (let k = from, n = 0; k <= to && n < 120; k = addDaysKey_(k, 1), n++) days[k] = true;
      } else {
        const tid = norm_(r[6]);
        if (!tid) return;
        if (/免/.test(type)) out.task[tid + '|' + sid] = '免除';
        else if (/配|慮|延/.test(type) && out.task[tid + '|' + sid] !== '免除') out.task[tid + '|' + sid] = '配慮';
      }
    });
  }
  EXC_ = out;
  return out;
}

function keyToUtc_(key) {
  const p = key.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2]));
}

function utcToKey_(d) {
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

function addDaysKey_(key, n) {
  const d = keyToUtc_(key);
  d.setUTCDate(d.getUTCDate() + n);
  return utcToKey_(d);
}

function isWeekday_(key) {
  const w = keyToUtc_(key).getUTCDay();
  return w !== 0 && w !== 6;
}

// 土日をとばして n 日あとの日付
function addSchoolDays_(key, n) {
  let k = key;
  while (n > 0) { k = addDaysKey_(k, 1); if (isWeekday_(k)) n--; }
  return k;
}

function isExempt_(task, sid) {
  return excuses_().task[task.id + '|' + sid] === '免除';
}

// その生徒にとっての締切
//  due：これを過ぎて出したら「遅れ」（空＝遅れにしない）
//  deadline：これを過ぎて出していなければ「未提出（赤）」（空＝赤にしない）
//  shown：生徒に見せる締切（欠席で延ばしたときだけ変わる）
//  ext：欠席で延ばした日数　grace：課題の「◯日まで期限内」　excused：配慮
function dueFor_(task, sid) {
  if (!task.due) return { due: '', deadline: '', shown: '', ext: 0, grace: 0, excused: false };
  const ex = excuses_();
  if (ex.task[task.id + '|' + sid] === '配慮') return { due: '', deadline: '', shown: task.due, ext: 0, grace: 0, excused: true };
  let ext = 0;
  const days = ex.absent[sid];
  if (days) {
    // 課題を出した日から締切日までに休んだ日（平日）を数える。出した日が空欄なら締切の1週間前から
    const start = task.given && task.given <= task.due ? task.given : addDaysKey_(task.due, -6);
    Object.keys(days).forEach(function (k) { if (k >= start && k <= task.due && isWeekday_(k)) ext++; });
  }
  const pol = task.late || latePolicy_('');
  const grace = pol.mode === 'grace' ? pol.days : 0;
  const deadline = ext + grace ? addSchoolDays_(task.due, ext + grace) : task.due;
  return {
    due: pol.mode === 'none' ? '' : deadline, deadline: deadline,
    shown: ext ? addSchoolDays_(task.due, ext) : task.due, ext: ext, grace: grace, excused: false,
  };
}

function isOverdue_(task, sid, today) {
  const d = dueFor_(task, sid);
  return !!d.deadline && today > d.deadline;
}

function readLog_() {
  const sh = sheet_(SHEET.LOG);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues().map(function (r, i) {
    return { time: r[0], taskId: norm_(r[1]), studentId: norm_(r[4]), incomplete: /未完成/.test(String(r[9])), row: i + 2 };
  }).filter(function (r) { return r.taskId && r.studentId; });
}

// 未完成を「提出」として数えるか（設定。ふつうは数えない＝未提出あつかい）
let INC_COUNTS_ = null;
function incompleteCounts_() {
  if (INC_COUNTS_ === null) INC_COUNTS_ = /^提出/.test(String(getSettings_()[SETTING.INCOMPLETE] || '').trim());
  return INC_COUNTS_;
}

// 提出として数えるか（出していない・未完成で未提出あつかい → false）
function counts_(sub) {
  return !!sub && (!sub.incomplete || incompleteCounts_());
}

// 遅れ＝その生徒の締切（欠席で延ばした分・配慮をふくむ）を過ぎてから出した
function isLate_(task, time, sid) {
  const d = dueFor_(task, sid);
  return !!d.due && time instanceof Date && dateKey_(time) > d.due;
}

// 課題ごとに「だれが・いつ」出したか（同じ生徒が2回あれば最初の記録）
//  late：遅れ　excused：もとの締切は過ぎたが、欠席・配慮で期限内あつかい
//  incomplete：未完成（出し直していない）。未完成のあとに出し直した記録があれば、そちらを使う
function submissionsOf_(task, logs) {
  const map = {};
  logs.forEach(function (r) {
    if (r.taskId !== task.id) return;
    const cur = map[r.studentId];
    if (cur && !(cur.incomplete && !r.incomplete)) return;
    const late = isLate_(task, r.time, r.studentId);
    const afterDue = !!task.due && r.time instanceof Date && dateKey_(r.time) > task.due;
    const none = task.late && task.late.mode === 'none';
    map[r.studentId] = { time: r.time, late: late, excused: afterDue && !late && !none, incomplete: r.incomplete, resubmitted: !!cur };
  });
  return map;
}

function taskStats_(task, students, logs) {
  const subs = submissionsOf_(task, logs);
  let target = 0, submitted = 0, late = 0, incomplete = 0;
  const missing = [];
  const byClass = {}; // クラスごとの数（読み取り画面で「担当クラスだけ」を出すため）
  sortStudents_(students).forEach(function (s) {
    if (!isTarget_(task, s)) return;
    target++;
    const bc = byClass[s.cls] || (byClass[s.cls] = { target: 0, submitted: 0, late: 0 });
    bc.target++;
    const sub = subs[s.id];
    if (sub && sub.incomplete) incomplete++;
    if (counts_(sub)) {
      submitted++;
      bc.submitted++;
      if (sub.late) { late++; bc.late++; }
    } else {
      const d = dueFor_(task, s.id);
      missing.push({ gakuseki: s.gakuseki, cls: s.cls, no: s.no, name: s.name, note: sub ? '未完成' : (d.excused ? '配慮' : (d.ext ? '〆' + shortDate_(d.shown) : '')) });
    }
  });
  return {
    taskId: task.id, label: taskLabel_(task), target: target, submitted: submitted, late: late, incomplete: incomplete,
    rate: target ? submitted / target : 0, missing: missing, byClass: byClass,
    overdue: !!task.due && todayKey_() > task.due,
  };
}

// ───────── 読み取り画面（Scan.html）から呼ばれる ─────────

function getScanInit() {
  setup_(false);
  const tasks = getTasks_().reverse(); // 新しい課題を上に
  const last = PropertiesService.getUserProperties().getProperty(userKey_('lastTask'));
  const lastTask = tasks.some(function (t) { return t.id === last; }) ? last : (tasks[0] ? tasks[0].id : '');
  return {
    tasks: tasks.map(function (t) { return { id: t.id, label: t.id + ' ' + taskLabelT_(t) }; }),
    taskId: lastTask,
    stats: lastTask ? getTaskStats(lastTask) : null,
    myClasses: myClasses_(),
  };
}

function getTaskStats(taskId) {
  const task = findTask_(getTasks_(), taskId);
  if (!task) return null;
  PropertiesService.getUserProperties().setProperty(userKey_('lastTask'), task.id);
  return taskStats_(task, getStudents_(), readLog_());
}

function findTask_(tasks, id) {
  const key = norm_(id);
  for (let i = 0; i < tasks.length; i++) if (tasks[i].id === key) return tasks[i];
  return null;
}

// mode：'submit'（提出）／'incomplete'（未完成）／'undo'（取り消し）。古い画面からの true は取り消し
// when：提出した日時（教務手帳に日付を書いたとき）。ふだんは今
function recordScan(taskId, raw, mode, when) {
  if (mode === true) mode = 'undo';
  if (mode !== 'undo' && mode !== 'incomplete') mode = 'submit';
  const code = norm_(raw);
  if (!code) return { kind: 'empty' };
  const tasks = getTasks_();

  // 課題バーコードを読んだら、その課題に切りかえる
  const asTask = findTask_(tasks, code);
  if (asTask) {
    return { kind: 'task', taskId: asTask.id, label: asTask.id + ' ' + taskLabel_(asTask), message: '課題を切りかえました：' + taskLabel_(asTask), stats: getTaskStats(asTask.id) };
  }

  const task = findTask_(tasks, taskId);
  if (!task) return { kind: 'error', message: '先に課題を選んでください' };
  const students = getStudents_();
  let st = null;
  for (let i = 0; i < students.length; i++) if (students[i].id === code) { st = students[i]; break; }
  if (!st) {
    // 手で学籍番号を打ったとき（バーコードが読めない・持っていない生徒）
    const hits = students.filter(function (x) { return x.gakuseki && x.gakuseki === code && !x.excluded; });
    if (hits.length === 1) st = hits[0];
  }

  if (!st) {
    return { kind: 'error', message: '登録されていないバーコードです（' + code + '）', stats: taskStats_(task, students, readLog_()) };
  }

  const who = whoOf_(st);
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  let result;
  try {
    const sh = sheet_(SHEET.LOG);
    const logs = readLog_();
    const mine = logs.filter(function (r) { return r.taskId === task.id && r.studentId === st.id; });

    if (mode === 'incomplete') {
      // 出した記録を「未完成」にする（まだ記録がなければ、未完成として記録する）
      if (mine.length) {
        mine.forEach(function (r) { sh.getRange(r.row, 10).setValue('未完成'); });
      } else {
        sh.appendRow([when instanceof Date ? when : new Date(), task.id, task.subject, task.name, st.id, st.gakuseki, st.cls, st.no, st.name, '未完成']);
        sh.getRange(sh.getLastRow(), 5, 1, 3).setNumberFormat('@').setValues([[st.id, st.gakuseki, st.cls]]);
      }
      result = { kind: 'incomplete', message: '未完成：' + who + (incompleteCounts_() ? '（提出として数えます）' : '（未提出あつかい。直して出したら、ふつうに読む）'), studentId: st.id };
    } else if (mode === 'undo') {
      if (!mine.length) {
        result = { kind: 'error', message: '取り消す記録がありません：' + who };
      } else {
        mine.map(function (r) { return r.row; }).sort(function (a, b) { return b - a; }).forEach(function (row) { sh.deleteRow(row); });
        result = { kind: 'undone', message: '取り消しました：' + who };
      }
    } else if (mine.some(function (r) { return !r.incomplete; })) {
      const done = mine.filter(function (r) { return !r.incomplete; })[0];
      const t = done.time instanceof Date ? Utilities.formatDate(done.time, tz_(), 'M/d HH:mm') : '';
      result = { kind: 'dup', message: 'すでに提出済み：' + who + (t ? '（' + t + '）' : '') };
    } else {
      const resubmit = mine.length > 0; // 未完成だった生徒が出し直した
      const now = when instanceof Date ? when : new Date();
      const late = isLate_(task, now, st.id);
      const target = isTarget_(task, st) || isExemptTarget_(task, st);
      const d = dueFor_(task, st.id);
      const saved = !late && !!task.due && dateKey_(now) > task.due;
      const judge = (!target ? '対象外' : (late ? '遅れ' : '期限内')) + (resubmit ? '（再提出）' : '');
      sh.appendRow([now, task.id, task.subject, task.name, st.id, st.gakuseki, st.cls, st.no, st.name, judge]);
      sh.getRange(sh.getLastRow(), 5, 1, 3).setNumberFormat('@').setValues([[st.id, st.gakuseki, st.cls]]);
      const exempt = isExemptTarget_(task, st);
      result = {
        kind: !target ? 'notarget' : (exempt ? 'ok' : (late ? 'late' : 'ok')),
        message: (resubmit ? '再提出（完成）・' : '') + (!target ? '対象外のクラスですが記録しました：' : (exempt ? '提出（この課題は免除の生徒）：' : (late ? '締切後の提出：' : '提出：'))) + who +
          (saved && target && !(task.late && task.late.mode === 'none') ? (d.excused ? '（配慮：期限内あつかい）' : (d.ext ? '（欠席のため締切を延長：期限内）' : '（' + d.grace + '日の猶予：期限内）')) : ''),
        studentId: st.id,
      };
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  const logsAfter = readLog_();
  if (result.kind !== 'dup' && result.kind !== 'error') result.warn = updateStatusRow_(st, tasks, logsAfter);
  result.stats = taskStats_(task, students, logsAfter);
  result.taskId = task.id;
  return result;
}

// 選んでいる課題の、いちばん新しい記録を1件消す
function undoLast(taskId) {
  const task = findTask_(getTasks_(), taskId);
  if (!task) return { kind: 'error', message: '課題を選んでください' };
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  let result, undoneId = '';
  try {
    const sh = sheet_(SHEET.LOG);
    const logs = readLog_().filter(function (r) { return r.taskId === task.id; });
    if (!logs.length) {
      result = { kind: 'error', message: 'この課題の記録はまだありません' };
    } else {
      const lastRow = logs[logs.length - 1];
      sh.deleteRow(lastRow.row);
      SpreadsheetApp.flush();
      undoneId = lastRow.studentId;
      result = { kind: 'undone', message: '取り消しました：' + lastRow.studentId };
    }
  } finally {
    lock.releaseLock();
  }
  const students = getStudents_();
  const logsAfter = readLog_();
  const st = students.filter(function (x) { return x.id === undoneId; })[0];
  if (st) {
    result.message = '取り消しました：' + whoOf_(st);
    result.warn = updateStatusRow_(st, getTasks_(), logsAfter);
  }
  result.stats = taskStats_(task, students, logsAfter);
  return result;
}

// ───────── 集計 ─────────

function refreshAll() {
  setup_(false);
  const ss = SpreadsheetApp.getActive();
  const allStudents = getStudents_();
  const allTasks = getTasks_();
  const logs = readLog_();
  const today = todayKey_();
  const warnings = [];

  const idCount = {};
  allStudents.forEach(function (s) { idCount[s.id] = (idCount[s.id] || 0) + 1; });
  const dupIds = Object.keys(idCount).filter(function (k) { return idCount[k] > 1; });
  if (dupIds.length) warnings.push('生徒IDが重なっています：' + dupIds.join(', '));
  const gkCount = {};
  allStudents.forEach(function (s) { if (s.gakuseki && !s.excluded) gkCount[s.gakuseki] = (gkCount[s.gakuseki] || 0) + 1; });
  const dupGk = Object.keys(gkCount).filter(function (k) { return gkCount[k] > 1; });
  if (dupGk.length) warnings.push('学籍番号が重なっています：' + dupGk.join(', '));

  const active = allStudents.filter(function (s) { return !s.excluded; });
  const subsByTask = {};
  allTasks.forEach(function (t) { subsByTask[t.id] = submissionsOf_(t, logs); });

  // 課題一覧の J〜N 列（全課題・全生徒。設定のしぼりこみはかけない）
  const tkSheet = sheet_(SHEET.TASKS);
  allTasks.forEach(function (t) {
    const st = taskStats_(t, active, logs);
    tkSheet.getRange(t.row, TASK_STAT_COL, 1, 5).setValues([[st.target, st.submitted, st.late, st.target - st.submitted, st.target ? st.rate : '']]);
  });
  tkSheet.getRange('J2:M').setNumberFormat('0');
  tkSheet.getRange('N2:N').setNumberFormat('0%');

  const view = statusView_(allTasks, allStudents);
  writeStatusSheet_(ss, view.tasks, view.students, subsByTask, today, view.cfg.statusName);
  writeSummarySheet_(ss, view.tasks, view.students, subsByTask, today, view.cfg);
  if (view.cfg.personal) {
    // 自分専用の表を使っている先生が押しても、全員共通の「提出状況」も新しくする（ほかのクラス・課題も見られるように）
    const common = statusView_(allTasks, allStudents, sharedConfig_());
    writeStatusSheet_(ss, common.tasks, common.students, subsByTask, today, common.cfg.statusName);
  }

  ss.toast('提出状況・集計を更新しました' + (warnings.length ? '\n⚠ ' + warnings.join('\n⚠ ') : ''), '課題バーコード', warnings.length ? 15 : 5);
  return { ok: true, warnings: warnings };
}

// 「設定」の教科・クラスのしぼりこみをかけた、表に出す課題と生徒
function statusView_(allTasks, allStudents, cfg) {
  cfg = cfg || viewConfig_();
  const subjectFilter = splitList_(cfg.subject);
  const classFilter = splitList_(cfg.cls).map(normClass_);
  return {
    cfg: cfg,
    tasks: allTasks.filter(function (t) {
      if (cfg.mine && t.teacher !== cfg.teacher) return false;
      return !subjectFilter.length || subjectFilter.indexOf(t.subject.normalize('NFKC').replace(/\s+/g, '')) >= 0;
    }),
    students: sortStudents_(allStudents.filter(function (s) { return !s.excluded && (!classFilter.length || classFilter.indexOf(s.cls) >= 0); })),
    mark: markMode_(),
  };
}

// ───────── 提出状況（教務手帳の形） ─────────
//  1行目（かくし）：課題ID　　A列（かくし）：生徒ID
//  2〜4行目：教科／課題名／締切日　　5行目から：1人1行（学籍番号・氏名・課題ごとの提出日・提出数・提出率・未提出）
const STATUS_TOP = 5;
const STATUS_FIXED = 3;

function statusCell_(s, t, subsByTask, today, mark) {
  if (isExemptTarget_(t, s)) return { v: '免', bg: COLOR.none, considered: false };
  if (!isTarget_(t, s)) return { v: '－', bg: COLOR.none, considered: false };
  const sub = subsByTask[t.id][s.id];
  if (sub && sub.incomplete) {
    // 未完成：△。数え方は設定しだい（ふつうは未提出あつかい）
    const counted = incompleteCounts_();
    const dd = dueFor_(t, s.id);
    const over = !!dd.deadline && today > dd.deadline;
    return { v: '△', bg: COLOR.incomplete, considered: counted || over || (!t.due && !dd.excused), done: counted, onTime: false, overdue: !counted && over };
  }
  if (sub) {
    const v = mark === 'circle' ? (sub.late ? '遅' : '○') : (sub.time instanceof Date ? sub.time : '○');
    return { v: v, bg: sub.late ? COLOR.late : (sub.excused ? COLOR.excused : COLOR.ok), considered: true, done: true, onTime: !sub.late };
  }
  const d = dueFor_(t, s.id);
  const overdue = !!d.deadline && today > d.deadline;
  return { v: '', bg: overdue ? COLOR.overdue : COLOR.notYet, considered: overdue || (!t.due && !d.excused), overdue: overdue };
}

function statusRow_(s, tasks, subsByTask, today, mark) {
  let considered = 0, done = 0, onTime = 0, overdueMissing = 0;
  const vals = [s.id, s.gakuseki || (s.cls + '-' + s.no), s.name];
  const bgs = ['#ffffff', '#ffffff', '#ffffff'];
  tasks.forEach(function (t) {
    const c = statusCell_(s, t, subsByTask, today, mark);
    if (c.considered) considered++;
    if (c.done) done++;
    if (c.onTime) onTime++;
    if (c.overdue) overdueMissing++;
    vals.push(c.v);
    bgs.push(c.bg);
  });
  vals.push(done, considered ? done / considered : '', considered ? onTime / considered : '', overdueMissing || '');
  bgs.push('#ffffff', '#ffffff', '#ffffff', overdueMissing ? COLOR.overdue : '#ffffff');
  return { vals: vals, bgs: bgs };
}

// 教務手帳の右はし4列の表示：提出数（数）・提出率・期限内提出率（％）・未提出（数）
const TAIL_FORMATS = ['0', '0%', '0%', '0'];

function writeStatusSheet_(ss, tasks, students, subsByTask, today, sheetName) {
  const mark = markMode_();
  const sh = sheetOrNew_(ss, sheetName || SHEET.STATUS);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('General'); // 前の表の「％」などが残らないように
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  if (sh.getMaxColumns() > 1) sh.showColumns(1, sh.getMaxColumns());
  if (sh.getMaxRows() > 1) sh.showRows(1, sh.getMaxRows());

  const n = tasks.length;
  const width = STATUS_FIXED + n + 4;
  const head = [
    ['（この行とA列は消さないでください）', '', ''].concat(tasks.map(function (t) { return t.id; }), ['', '', '', '']),
    ['', '', '教科'].concat(tasks.map(function (t) { return t.subject; }), ['', '', '', '']),
    ['', '', '課題名'].concat(tasks.map(function (t) { return t.name; }), ['提出数', '提出率', '期限内\n提出率', '未提出\n（締切後）']),
    ['生徒ID', '学籍番号', '氏名'].concat(tasks.map(function (t) {
      return t.due ? '〆' + shortDate_(t.due) : (t.given ? '出' + shortDate_(t.given) : '');
    }), ['', '', '', '']),
  ];
  const values = head.slice();
  const colors = head.map(function () { const r = []; for (let i = 0; i < width; i++) r.push(COLOR.header); return r; });
  students.forEach(function (s) {
    const r = statusRow_(s, tasks, subsByTask, today, mark);
    values.push(r.vals);
    colors.push(r.bgs);
  });

  ensureSize_(sh, values.length, width);
  sh.getRange(1, 1, values.length, 2).setNumberFormat('@'); // 生徒ID・学籍番号（「1-1-3」などが日付にならないように）
  const rng = sh.getRange(1, 1, values.length, width);
  rng.setValues(values).setBackgrounds(colors).setVerticalAlignment('middle').setFontSize(10);
  sh.getRange(2, 1, 3, width).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true);
  sh.getRange(2, 3, 2, 1).setHorizontalAlignment('right').setFontColor('#555555');
  sh.getRange(3, 1, 1, width).setVerticalAlignment('top');
  if (students.length) {
    const body = sh.getRange(STATUS_TOP, 1, students.length, width);
    body.setBorder(true, true, true, true, true, true, '#bbbbbb', SpreadsheetApp.BorderStyle.SOLID);
    if (n) sh.getRange(STATUS_TOP, STATUS_FIXED + 1, students.length, n).setNumberFormat('m/d').setHorizontalAlignment('center');
    sh.getRange(STATUS_TOP, STATUS_FIXED + n + 1, students.length, 4).setHorizontalAlignment('center');
    sh.getRange(STATUS_TOP, STATUS_FIXED + n + 1, students.length, 4).setNumberFormats(students.map(function () { return TAIL_FORMATS; }));
    // クラスの切れ目に太線（教務手帳のページの区切りのように）
    for (let i = 1; i < students.length; i++) {
      if (students[i].cls !== students[i - 1].cls) {
        sh.getRange(STATUS_TOP + i, 1, 1, width).setBorder(true, null, null, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }
    }
  }
  sh.getRange(2, 1, 3, width).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  sh.setRowHeight(2, 22);
  sh.setRowHeight(3, 110);
  sh.setRowHeight(4, 22);
  sh.setColumnWidth(2, 70);
  sh.setColumnWidth(3, 120);
  for (let c = 0; c < n; c++) sh.setColumnWidth(STATUS_FIXED + 1 + c, 52);
  [50, 55, 55, 60].forEach(function (w, i) { sh.setColumnWidth(STATUS_FIXED + n + 1 + i, w); });
  sh.setFrozenRows(4);
  sh.setFrozenColumns(STATUS_FIXED);
  sh.hideRows(1);
  sh.hideColumns(1);
  sh.getRange(2, 2).setValue('教務手帳').setFontSize(12);
  sh.getRange(3, 2).setValue((mark === 'circle' ? '○＝提出\n遅＝遅れ\n' : '日付＝提出日\n黄＝遅れ\n') + '青＝欠席・配慮で期限内\n△＝未完成\n赤＝未提出（締切後）\n免＝免除　－＝対象外')
    .setFontWeight('normal').setFontSize(8).setHorizontalAlignment('left');
  sh.getRange(4, 2).setNote('提出率は「締切を過ぎた課題・締切のない課題・提出済みの課題」で計算。\n最終更新：' + Utilities.formatDate(new Date(), tz_(), 'M/d HH:mm'));
}

// 読み取るたびに、その生徒の行だけを書きかえる
//  自分の教務手帳：課題・生徒が足りなければ全体を作り直す
//  ほかの先生の教務手帳：その課題の列とその生徒の行があれば、その行だけ書きかえる
// 失敗しても読み取りは止めず、理由を返す（読み取り画面に出す）
function updateStatusRow_(st, allTasks, logs) {
  try {
    const ss = SpreadsheetApp.getActive();
    const view = statusView_(allTasks, getStudents_());
    const today = todayKey_();
    const taskMap = {};
    allTasks.forEach(function (t) { taskMap[t.id] = t; });
    const subsCache = {};
    const subsFor = function (ts) {
      const out = {};
      ts.forEach(function (t) { out[t.id] = subsCache[t.id] || (subsCache[t.id] = submissionsOf_(t, logs)); });
      return out;
    };
    ss.getSheets().forEach(function (sh) {
      if (!isStatusSheet_(sh.getName())) return;
      const own = sh.getName() === view.cfg.statusName;
      const built = sh.getLastRow() >= 4 && String(sh.getRange(4, 3).getValue()) === '氏名';
      const lastCol = built ? sh.getLastColumn() : 0;
      const ids = lastCol > STATUS_FIXED ? sh.getRange(1, STATUS_FIXED + 1, 1, lastCol - STATUS_FIXED).getValues()[0].map(norm_) : [];
      const sheetIds = ids.slice(0, ids.indexOf('') < 0 ? ids.length : ids.indexOf(''));
      if (own && (!built || sheetIds.join(',') !== view.tasks.map(function (t) { return t.id; }).join(','))) {
        writeStatusSheet_(ss, view.tasks, view.students, subsFor(view.tasks), today, sh.getName());
        return;
      }
      if (!built) return;
      const tasks = sheetIds.map(function (id) { return taskMap[id]; });
      if (tasks.some(function (t) { return !t; })) return; // 課題が消えた表は、その先生が開いたときに作り直す
      const nRows = sh.getLastRow() - STATUS_TOP + 1;
      const col = nRows > 0 ? sh.getRange(STATUS_TOP, 1, nRows, 1).getValues().map(function (r) { return norm_(r[0]); }) : [];
      const idx = col.indexOf(st.id);
      if (idx < 0) {
        if (own && view.students.some(function (s) { return s.id === st.id; })) writeStatusSheet_(ss, view.tasks, view.students, subsFor(view.tasks), today, sh.getName());
        return; // しぼりこみで表に出ていない生徒
      }
      const r = statusRow_(st, tasks, subsFor(tasks), today, view.mark);
      sh.getRange(STATUS_TOP + idx, 1, 1, r.vals.length).setValues([r.vals]).setBackgrounds([r.bgs]);
      // マスに日付を打つと表示が「2026/10/06」に変わってしまうので、月/日に戻す
      if (tasks.length) sh.getRange(STATUS_TOP + idx, STATUS_FIXED + 1, 1, tasks.length).setNumberFormat('m/d');
      sh.getRange(STATUS_TOP + idx, STATUS_FIXED + tasks.length + 1, 1, 4).setNumberFormats([TAIL_FORMATS]); // 提出数が％にならないように
    });
  } catch (err) {
    console.warn('提出状況の書きかえに失敗：' + (err && err.stack || err));
    return '記録はできましたが、提出状況（教務手帳）の書きかえに失敗しました：' + (err && err.message || err);
  }
  return '';
}

function writeSummarySheet_(ss, tasks, students, subsByTask, today, cfg) {
  cfg = cfg || viewConfig_();
  const sh = sheetOrNew_(ss, cfg.summaryName);
  sh.clear();
  const classes = classList_(students);
  const header = ['課題ID', '教科', '課題名', '締切', '対象', '提出', 'うち遅れ', '未提出', '未完成', '免除', '提出率', '期限内\n提出率'].concat(classes.map(function (c) { return c + '\n提出率'; }));
  const rows = [];
  const bySubject = {};

  tasks.forEach(function (t) {
    const subs = subsByTask[t.id];
    let target = 0, done = 0, late = 0, exempt = 0, inc = 0, onTime = 0;
    const perClass = {};
    students.forEach(function (s) {
      if (isExemptTarget_(t, s)) exempt++;
      if (!isTarget_(t, s)) return;
      target++;
      const pc = perClass[s.cls] || (perClass[s.cls] = { target: 0, done: 0 });
      pc.target++;
      const sub = subs[s.id];
      if (sub && sub.incomplete) inc++;
      if (counts_(sub)) { done++; pc.done++; if (sub.late) late++; if (!sub.late && !sub.incomplete) onTime++; }
    });
    rows.push([t.id, t.subject, displayName_(t), t.due ? shortDate_(t.due) : '', target, done, late, target - done, inc || '', exempt || '', target ? done / target : '', target ? onTime / target : '']
      .concat(classes.map(function (c) { return perClass[c] ? perClass[c].done / perClass[c].target : ''; })));
    const key = t.subject || '（教科なし）';
    const b = bySubject[key] || (bySubject[key] = { tasks: 0, target: 0, done: 0, late: 0, onTime: 0 });
    b.tasks++; b.target += target; b.done += done; b.late += late; b.onTime += onTime;
  });

  ensureSize_(sh, rows.length + Object.keys(bySubject).length + 10, header.length);
  const filt = [cfg.mine ? cfg.teacher + 'の課題' : '', String(cfg.subject || '').trim(), String(cfg.cls || '').trim()].filter(String);
  sh.getRange(1, 1).setValue('課題別の提出状況' + (filt.length ? '（しぼりこみ中：' + filt.join('／') + '）' : '')).setFontWeight('bold').setFontSize(13);
  sh.getRange(2, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground(COLOR.header).setWrap(true);
  if (rows.length) {
    sh.getRange(3, 1, rows.length, header.length).setValues(rows);
    sh.getRange(3, 5, rows.length, 6).setNumberFormat('0');
    sh.getRange(3, 11, rows.length, header.length - 10).setNumberFormat('0%');
  }

  const top = 3 + rows.length + 2;
  sh.getRange(top, 1).setValue('教科別の提出率').setFontWeight('bold').setFontSize(13);
  const subjHeader = ['教科', '課題数', '対象（のべ）', '提出（のべ）', 'うち遅れ', '提出率', '期限内提出率'];
  const subjRows = Object.keys(bySubject).map(function (k) {
    const b = bySubject[k];
    return [k, b.tasks, b.target, b.done, b.late, b.target ? b.done / b.target : '', b.target ? b.onTime / b.target : ''];
  });
  sh.getRange(top + 1, 1, 1, subjHeader.length).setValues([subjHeader]).setFontWeight('bold').setBackground(COLOR.header);
  if (subjRows.length) {
    sh.getRange(top + 2, 1, subjRows.length, subjHeader.length).setValues(subjRows);
    sh.getRange(top + 2, 2, subjRows.length, 4).setNumberFormat('0');
    sh.getRange(top + 2, 6, subjRows.length, 2).setNumberFormat('0%');
  }
  sh.setColumnWidth(3, 220);
  sh.setFrozenRows(2);
  sh.getRange(1, 3).setValue('最終更新：' + Utilities.formatDate(new Date(), tz_(), 'M/d HH:mm')).setFontColor('#666666');
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

  // まだ出していない、かどうか（「締切を過ぎた課題だけ」のときは、その生徒の締切＝欠席で延ばした締切で見る）
  const isMissing = function (t, s) {
    return isTarget_(t, s) && !counts_(subs[t.id][s.id]) && (!opts.overdueOnly || isOverdue_(t, s.id, today));
  };
  const incompleteOf = function (t, s) { return !!subs[t.id][s.id]; }; // 未提出の中で、未完成で出したもの
  const slips = [];
  const count = opts.type === 'notice'
    ? writeNotices_(sh, tasks, students, isMissing, incompleteOf, title, String(settings[SETTING.NOTICE] || ''), teacherLabel_(), slips)
    : writeMissingList_(sh, tasks, students, isMissing, incompleteOf, title + (teacherLabel_() ? '　' + teacherLabel_() : ''));

  ss.setActiveSheet(sh);
  SpreadsheetApp.flush();
  // 「課題提出について」は、短冊がページの境目で切れないように、印刷画面（ブラウザ）で印刷する
  return { ok: true, count: count, sheetName: name, pdfUrl: pdfUrl_(ss, sh), slips: opts.type === 'notice' ? slips : null };
}

function writeMissingList_(sh, tasks, students, isMissing, incompleteOf, title) {
  [70, 60, 45, 160, 60, 230].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  let row = 1;
  sh.getRange(row, 1, 1, 6).merge().setValue('未提出者一覧' + (title ? '　' + title : '') + '　（' + Utilities.formatDate(new Date(), tz_(), 'M/d HH:mm') + ' 現在）').setFontSize(14).setFontWeight('bold');
  row += 2;
  let total = 0;
  tasks.forEach(function (t) {
    const missing = students.filter(function (s) { return isMissing(t, s); });
    const target = students.filter(function (s) { return isTarget_(t, s); }).length;
    ensureSize_(sh, row + missing.length + 5, 6);
    sh.getRange(row, 1, 1, 6).merge()
      .setValue(taskLabel_(t) + '　未提出 ' + missing.length + '人 ／ 対象 ' + target + '人')
      .setFontWeight('bold').setBackground('#eeeeee').setBorder(true, false, true, false, false, false);
    row++;
    if (!missing.length) {
      sh.getRange(row, 1).setValue('全員提出しています');
      row += 2;
      return;
    }
    sh.getRange(row, 1, 1, 6).setValues([['学籍番号', 'クラス', '番号', '氏名', '確認', 'メモ']]).setFontWeight('bold').setFontColor('#555555');
    row++;
    const vals = missing.map(function (s) { return [s.gakuseki, s.cls, s.no, s.name, '□', [incompleteOf(t, s) ? '未完成（再提出）' : '', dueNote_(t, s.id)].filter(String).join('／')]; });
    sh.getRange(row, 1, vals.length, 2).setNumberFormat('@'); // 学籍番号・クラス（「1-1」が日付にならないように）
    sh.getRange(row, 1, vals.length, 6).setValues(vals).setBorder(null, null, true, null, false, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
    row += vals.length + 1;
    total += missing.length;
  });
  return total;
}

// 「担当：青柳」（この先生の名前をそのまま出す）
function teacherLabel_() {
  const name = teacherName_();
  return name ? '担当：' + name : '';
}

// この先生の名前（先生ごとに覚える。1人で使っていたころの「設定」シートの名前も使う）
function teacherName_() {
  const own = PropertiesService.getUserProperties().getProperty(userKey_('teacher'));
  if (own != null) return String(own).trim();
  return String(getSettings_()[SETTING.TEACHER] || '').trim();
}

// 教務手帳・集計の見せ方
//  ふつう（1人で使う・みんなで同じ表を見る）：「提出状況」「集計」シート、しぼりこみは「設定」シート
//  自分専用：「提出状況（青柳）」「集計（青柳）」シート、しぼりこみ・自分の課題だけはこの先生だけのもの
function viewConfig_() {
  const p = PropertiesService.getUserProperties();
  const settings = getSettings_();
  const name = teacherName_();
  if (p.getProperty(userKey_('personal')) === '1' && name) {
    const safe = name.replace(/[\[\]\/\\?*:'']/g, '');
    return {
      personal: true, teacher: name,
      statusName: SHEET.STATUS + '（' + safe + '）', summaryName: SHEET.SUMMARY + '（' + safe + '）',
      subject: p.getProperty(userKey_('pSubject')) || '', cls: myClasses_().join(','),
      mine: p.getProperty(userKey_('pMine')) === '1',
    };
  }
  return sharedConfig_(name);
}

// 全員共通の「提出状況」「集計」シートの見せ方（「設定」シートのしぼりこみ）
function sharedConfig_(name) {
  const settings = getSettings_();
  return {
    personal: false, teacher: name || teacherName_(), statusName: SHEET.STATUS, summaryName: SHEET.SUMMARY,
    subject: String(settings[SETTING.SUBJECT] || ''), cls: String(settings[SETTING.CLASS] || ''), mine: false,
  };
}

// この先生の担当クラス（先生ごと。空＝全クラス）
function myClasses_() {
  return splitList_(PropertiesService.getUserProperties().getProperty(userKey_('myClasses')) || '').map(normClass_);
}

function isStatusSheet_(name) {
  return name === SHEET.STATUS || name.indexOf(SHEET.STATUS + '（') === 0;
}

function markMode_() {
  return /○|〇|丸/.test(String(getSettings_()[SETTING.MARK] || '')) ? 'circle' : 'date';
}

function sheetOrNew_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name, ss.getSheets().length);
}

function writeNotices_(sh, tasks, students, isMissing, incompleteOf, title, message, teacher, slips) {
  // A4 の横いっぱい（PDF の余白をのぞいて約 700px）に広げ、文字も大きくする
  [56, 140, 360, 150].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  let row = 1;
  let count = 0;
  students.forEach(function (s) {
    const missing = tasks.filter(function (t) { return isMissing(t, s); });
    if (!missing.length) return;
    count++;
    ensureSize_(sh, row + missing.length + 6, 4);
    const start = row;
    sh.getRange(row, 1, 1, 3).merge().setValue('課題提出について' + (title ? '　　' + title : '')).setFontWeight('bold').setFontSize(16).setVerticalAlignment('middle');
    sh.getRange(row, 4).setValue(teacher || '').setFontSize(13).setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.setRowHeight(row, 36);
    row++;
    sh.getRange(row, 1, 1, 4).merge().setValue((s.gakuseki ? s.gakuseki + '　' : '') + s.cls + '　' + s.no + '番　' + s.name + '　さん').setFontSize(20).setFontWeight('bold').setVerticalAlignment('middle');
    sh.setRowHeight(row, 42);
    row++;
    if (message) {
      sh.getRange(row, 1, 1, 4).merge().setValue(message).setWrap(true).setFontSize(13).setVerticalAlignment('middle');
      sh.setRowHeight(row, 40);
      row++;
    }
    const vals = missing.map(function (t) {
      const d = dueFor_(t, s.id);
      return ['□', t.subject, displayName_(t) + (incompleteOf(t, s) ? '（未完成・出し直し）' : ''), d.shown ? '締切 ' + shortDate_(d.shown) + (d.ext ? '（延長）' : '') : ''];
    });
    if (slips) {
      slips.push({
        head: '課題提出について' + (title ? '　　' + title : ''), teacher: teacher || '',
        pre: (s.gakuseki ? s.gakuseki + '　' : '') + s.cls + '　' + s.no + '番', name: s.name,
        message: message || '', items: vals.map(function (v) { return v.slice(1); }),
      });
    }
    sh.getRange(row, 1, vals.length, 4).setValues(vals).setFontSize(15).setVerticalAlignment('middle').setWrap(true);
    sh.getRange(row, 1, vals.length, 1).setHorizontalAlignment('center').setFontSize(18);
    for (let k = 0; k < vals.length; k++) sh.setRowHeight(row + k, 34);
    row += vals.length;
    sh.getRange(start, 1, row - start, 4).setBorder(true, true, true, true, false, false, '#888888', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(row, 1, 1, 4).merge().setValue('✂ - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -').setFontColor('#999999').setHorizontalAlignment('center').setFontSize(11);
    sh.setRowHeight(row, 30);
    row++;
  });
  if (!count) sh.getRange(1, 1).setValue('選んだ課題・クラスでは、未提出の生徒はいません。');
  return count;
}

// 未提出者リストのメモ欄：欠席で締切を延ばした・配慮している
function dueNote_(t, sid) {
  const d = dueFor_(t, sid);
  if (d.excused) return '配慮（期限内あつかい）';
  if (d.ext) return '欠席' + d.ext + '日 → 締切 ' + shortDate_(d.shown) + ' に延長';
  return '';
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
      label.push(s ? whoOf_(s) : '');
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

// 実際に文字が入っている最後の行（チェックボックスの FALSE は数えない）
function lastDataRow_(sh, ncols) {
  const last = sh.getLastRow();
  if (last < 2) return last;
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

function toScriptJson_(obj) {
  return JSON.stringify(obj).replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');
}
