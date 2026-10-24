/**
 * 資質・能力21 ふりかえりアプリ - 教師用データ集約・レポートスクリプト
 *
 * ▼セットアップ手順（Webアプリとして生徒に配布する場合）
 * 1. Googleスプレッドシートを新規作成する（例:「探究アセスメント集約」）
 * 2. メニュー「拡張機能」→「Apps Script」を開く
 * 3. デフォルトのコードを全て削除し、このファイル(.gs)の内容を貼り付けて保存する
 * 4. プロジェクトにHTMLファイルを2つ追加する（ファイル → 新規 → HTML）。
 *    - 「Index」：生徒用の画面（Index.html の中身を丸ごと貼り付ける）
 *    - 「Report」：先生用のレポート印刷画面（Report.html の中身を丸ごと貼り付ける）
 * 5. 画面右上「デプロイ」→「新しいデプロイ」をクリック
 *    - 種類の選択: 「ウェブアプリ」
 *    - 説明: 任意（例: 生徒用アプリ配布）
 *    - 実行ユーザー: 「自分」（＝このスプレッドシートの持ち主の権限で書き込む）
 *    - アクセスできるユーザー: 「〇〇（学校のドメイン名）内の全員」
 *      ※これにより、学校のGoogleアカウントでログインしないとアプリを開けなくなる
 *    ※スプレッドシート自体は生徒と共有しない（先生だけが開ける状態にしておく）。
 *      生徒はアプリ経由で自分の回答を送るだけで、他の生徒の回答は見られない。
 * 6. 「デプロイ」をクリックし、発行された「ウェブアプリのURL」を生徒に共有する
 *    （このURLをブックマーク・Google Classroom等で配れば、以後ファイル配布は不要になる）
 * 7. スプレッドシートを再読み込みすると、メニューバーに「探究アセスメント」が追加される
 * 8. メニュー「探究アセスメント」→「生徒名簿を設定する」で、今年度の名簿（出席番号・氏名）を貼り付ける
 * 9. メニュー「探究アセスメント」→「実施回を設定する」で、生徒が選ぶ実施回の名前を登録する
 *
 * ▼コードを貼り直したとき
 * 「デプロイ」→「デプロイを管理」→ 鉛筆マーク →「バージョン：新バージョン」→「デプロイ」を押すと、
 * 生徒用URLを変えずに新しいコードに切り替わる。
 * メニュー「探究アセスメント」→「版を確認」と、生徒画面のいちばん下に出る版が同じなら貼り直し完了。
 *
 * ▼出席番号のルール
 * 半角4桁のうち、1桁目=学年、2桁目=組、3〜4桁目=出席番号　という前提で集計します。
 * （例: 1215 → 1年2組15番）　※このルールと異なる場合は parseStudentNumber_() を調整してください。
 *
 * ▼シート構成
 * - 「回答」　　　: 生徒から送信された生データが1件1行で自動追記される
 * - 「生徒名簿_2026」など: 年度ごとの名簿（出席番号・氏名・ふりがな・Googleアカウント）。
 *     生徒は出席番号を入れるだけで名前が名簿から入り、レポートは名簿の番号順に並ぶ
 * - 「実施回」　　: 生徒が選ぶ実施回の一覧（B列にチェックが入っている回だけ生徒画面に出る）
 *
 * ▼レポート印刷（メニュー「レポートを作る・印刷する」）
 * クラス集計（クラスを選ぶ）／学年・クラスの推移／クラスの生徒の変容（名簿順の一覧表）／生徒個人の成長を
 * 画面で選んでプレビューし、そのままA4で印刷・PDF保存する。年度ごと（4月〜3月）に集計する。
 *
 * ▼安全のための約束
 * 「実行ユーザー: 自分」で公開しているため、名前の末尾が「_」でない関数は生徒の画面からも呼べてしまう。
 * 先生用の関数は最初に assertTeacher_() を呼び、スプレッドシートの編集者以外は止める。
 * 内部だけで使う関数は、名前の末尾を「_」にする（外から呼べなくなる）。
 *
 * ▼集計に使う回答
 * 「5項目選択」「自己×他己比較」モードは生徒が自分で選んだ5項目だけを答えるため、平均が高めに出やすい。
 * そのためレポートは「全21項目」モードの回答だけで計算する（Report.html の prepare()）。
 * 全21項目の回答が1件もないときに限り、5項目モードの回答で計算し、その旨をレポートに書く。
 */

// 生徒画面（Index）の APP_VERSION と同じ値にそろえる。貼り直しの確認に使う
const APP_VERSION = '2026-10-24.3';

const SHEET_NAME = '回答';

// 21の資質・能力（ID・名称・カテゴリー）※生徒用アプリのCOMPETENCIESと対応
const COMPETENCIES = [
  { id: 1,  category: 'Thinking',      name: '論理的思考力' },
  { id: 2,  category: 'Thinking',      name: '課題発見力' },
  { id: 3,  category: 'Thinking',      name: '課題解決力' },
  { id: 4,  category: 'Thinking',      name: '創造力' },
  { id: 5,  category: 'Thinking',      name: '学び続ける力' },
  { id: 6,  category: 'Thinking',      name: '知的好奇心' },
  { id: 7,  category: 'Thinking',      name: '柔軟性' },
  { id: 8,  category: 'Communication', name: '協働力' },
  { id: 9,  category: 'Communication', name: '表現力' },
  { id: 10, category: 'Communication', name: '発信力' },
  { id: 11, category: 'Communication', name: '傾聴力・共感力' },
  { id: 12, category: 'Communication', name: '対話力・質問力' },
  { id: 13, category: 'Communication', name: '多角的・多面的視点' },
  { id: 14, category: 'Communication', name: '深く掘り下げる力' },
  { id: 15, category: 'Leadership',    name: '目標への段取り力' },
  { id: 16, category: 'Leadership',    name: '主体的行動力' },
  { id: 17, category: 'Leadership',    name: '向上心' },
  { id: 18, category: 'Leadership',    name: '自己管理能力' },
  { id: 19, category: 'Leadership',    name: '自己有用感' },
  { id: 20, category: 'Leadership',    name: 'レジリエンス（挫折回復力）' },
  { id: 21, category: 'Leadership',    name: 'リーダーシップ' }
];
const CATEGORY_LABEL = { Thinking: 'Thinking（思考・探究）', Communication: 'Communication（発信・表現・協働）', Leadership: 'Leadership（主体性・実行力）' };
// 生徒用アプリと同じ配色（Thinking=青, Communication=緑, Leadership=紫）でレポートの見た目を揃える
const CATEGORY_COLOR = { Thinking: '#2c5f8a', Communication: '#3d7a5c', Leadership: '#6b4c93' };

const RESPONSE_BASE_HEADER = [
  '受信日時', '実施日', '出席番号', '氏名', 'タイトル', 'モード',
  '各項目スコア(JSON)', '振り返り(原文)', '振り返り(整文後)', '記録ID', '送信者Googleアカウント'
];
const RESPONSE_EXTRA_HEADER = [
  '項目別メモ(JSON)', '項目別メモ(一覧)', '他己評価人数',
  '他己評価平均(JSON)', '他己評価詳細(JSON)', '友人メッセージ(一覧)'
];
const SCORE_HEADER = COMPETENCIES.map(c => c.id + '.' + c.name);
const RESPONSE_HEADER = RESPONSE_BASE_HEADER.concat(RESPONSE_EXTRA_HEADER, SCORE_HEADER);
const SCORE_COL_START = RESPONSE_BASE_HEADER.length + RESPONSE_EXTRA_HEADER.length + 1; // 18列目(R)
const JSON_COLUMNS = [7, 12, 15, 16]; // 折りたたむJSON列

/**
 * 「回答」シートに1件追記する共通処理（doPost/submitAssessmentToServer両方から呼ばれる）
 * @param {Object} data 生徒アプリから送られてきた評価データ
 * @param {string} senderEmail 送信者のGoogleアカウント（不明な場合は空文字）
 */
function appendResponseRow_(data, senderEmail) {
  // 同時送信でも行が混ざったり、重複判定をすり抜けたりしないよう、書き込み中はロックする
  // クラス全員がほぼ同時に送信しても待ちきれずに失敗しないよう、待ち時間は長めにとる
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      ensureResponseColumns_(sheet);
      sheet.getRange(1, 1, 1, RESPONSE_HEADER.length).setValues([RESPONSE_HEADER]);
      formatResponseSheet_(sheet);
    }
    // 「サーバーには届いたが応答が返らなかった」記録を再送信しても、同じ記録の行は追加しない
    if (isDuplicateRecord_(sheet, data)) return;
    sheet.appendRow(buildResponseRow_(data, senderEmail));
  } finally {
    lock.releaseLock();
  }
}

/**
 * 「回答」シートに同じ記録が既にあるかを調べる。
 * 記録IDは送信した瞬間の時刻（ミリ秒）なので、別の生徒と偶然重なることがある。
 * 記録ID（J列）と出席番号（C列）の両方が一致したときだけ重複とみなす。
 */
function isDuplicateRecord_(sheet, data) {
  const recordId = data.id;
  if (recordId === undefined || recordId === null || recordId === '') return false;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  const rows = sheet.getRange(2, 3, lastRow - 1, 8).getValues(); // C列〜J列
  const targetId = String(recordId);
  const targetNumber = String(data.studentNumber || '');
  return rows.some(r => String(r[7]) === targetId && String(r[0]) === targetNumber);
}

/**
 * 1件の評価データを「回答」シートの1行分（RESPONSE_HEADERと同じ並び）に変換する
 */
function buildResponseRow_(data, senderEmail) {
  const scores = data.scores || {};
  const notes = data.notes || {};
  const peerEvals = Array.isArray(data.peerEvals) ? data.peerEvals : [];

  const noteLines = COMPETENCIES.filter(c => notes[c.id]).map(c => c.id + '.' + c.name + ': ' + notes[c.id]);
  const messageLines = peerEvals.filter(ev => ev && ev.message).map(ev => (ev.name || '友人') + ': ' + ev.message);

  const row = [
    new Date(),
    data.date || '',
    data.studentNumber || '',
    data.studentName || '',
    data.title || '',
    data.mode || '',
    JSON.stringify(scores),
    data.reflection || '',
    data.refinedReflection || '',
    data.id || '',
    senderEmail || '',
    Object.keys(notes).length ? JSON.stringify(notes) : '',
    noteLines.join('\n'),
    peerEvals.length || '',
    data.peerScores ? JSON.stringify(data.peerScores) : '',
    peerEvals.length
      ? JSON.stringify(peerEvals.map(ev => ({ name: ev.name || '', message: ev.message || '', scores: ev.scores || {} })))
      : '',
    messageLines.join('\n')
  ];
  COMPETENCIES.forEach(c => {
    const v = Number(scores[c.id]);
    row.push(v > 0 ? v : '');
  });
  // スプレッドシートの数式インジェクション対策：生徒が入力した文字列が「=」「+」「-」「@」で始まる場合、
  // 先頭にアポストロフィを付けて数式ではなく文字列として保存する（アポストロフィ自体はセルには表示されない）
  return row.map(v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v);
}

/** 「回答」シートの列数がRESPONSE_HEADER分に足りない場合、列を追加する */
function ensureResponseColumns_(sheet) {
  const need = RESPONSE_HEADER.length - sheet.getMaxColumns();
  if (need > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), need);
}

function parseJsonSafe_(text) {
  try { return JSON.parse(text || '{}') || {}; } catch (e) { return {}; }
}

/** 「回答」シートの見た目を整える（見出し色・幅・ヒートマップ・JSON列の折りたたみ） */
function formatResponseSheet_(sheet) {
  const lastCol = RESPONSE_HEADER.length;
  const bodyRows = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#f0ece3')
    .setVerticalAlignment('middle').setWrap(true);
  COMPETENCIES.forEach((c, i) => {
    sheet.getRange(1, SCORE_COL_START + i).setBackground(CATEGORY_COLOR[c.category]).setFontColor('#ffffff');
  });

  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(4);
  sheet.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.getRange('C:C').setNumberFormat('@'); // 出席番号は文字列として保存

  // 読みやすさ：スコア列は狭く中央寄せ、文章列は広めに折り返す
  sheet.setColumnWidths(SCORE_COL_START, COMPETENCIES.length, 78);
  [8, 9, 13, 17].forEach(col => {
    sheet.setColumnWidth(col, 280);
    sheet.getRange(2, col, bodyRows, 1).setWrap(true);
  });

  // スコア列のヒートマップ（1=オレンジ、2.5=黄、4=緑）
  const scoreRange = sheet.getRange(2, SCORE_COL_START, bodyRows, COMPETENCIES.length);
  scoreRange.setHorizontalAlignment('center');
  const heat = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#f4b183', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setGradientMidpointWithValue('#ffe699', SpreadsheetApp.InterpolationType.NUMBER, '2.5')
    .setGradientMaxpointWithValue('#a9d18e', SpreadsheetApp.InterpolationType.NUMBER, '4')
    .setRanges([scoreRange]).build();
  sheet.setConditionalFormatRules([heat]);

  // JSON列は列グループにして折りたたむ（見た目の整理のみ。失敗してもデータには影響しない）
  try {
    JSON_COLUMNS.forEach(col => {
      if (!sheet.getColumnGroup(col, 1)) sheet.getRange(1, col).shiftColumnGroupDepth(1);
      const g = sheet.getColumnGroup(col, 1);
      if (g) g.collapse();
    });
  } catch (e) {
    console.error('列グループの設定に失敗:', e);
  }
}

/**
 * メニュー用：既存の「回答」シートを新形式に更新し、過去の行の21項目スコア列も自動で埋める
 */
function upgradeResponseSheet() {
  assertTeacher_();
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert('「回答」シートがまだありません。最初の回答が届くと、新形式で自動作成されます。');
    return;
  }
  ensureResponseColumns_(sheet);
  sheet.getRange(1, 1, 1, RESPONSE_HEADER.length).setValues([RESPONSE_HEADER]);

  const rowCount = sheet.getLastRow() - 1;
  if (rowCount >= 1) {
    const jsonVals = sheet.getRange(2, 7, rowCount, 1).getValues();
    const existing = sheet.getRange(2, SCORE_COL_START, rowCount, COMPETENCIES.length).getValues();
    const out = jsonVals.map((r, i) => {
      const scores = parseJsonSafe_(r[0]);
      return COMPETENCIES.map((c, j) => {
        if (existing[i][j] !== '') return existing[i][j];
        const v = Number(scores[c.id]);
        return v > 0 ? v : '';
      });
    });
    sheet.getRange(2, SCORE_COL_START, rowCount, COMPETENCIES.length).setValues(out);
  }
  formatResponseSheet_(sheet);
  ui.alert('「回答」シートを新形式に更新しました（' + Math.max(rowCount, 0) + '行）。');
}

/**
 * アプリを開いている生徒のGoogleアカウントを返す。
 * 「実行ユーザー: 自分」でも、生徒が同じ学校ドメインのアカウントでログインしていれば取得できる
 * （ドメイン外のアカウントなどで取得できないときは空文字）。
 */
function getViewerEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

/**
 * 先生用の関数の最初に呼ぶ。スプレッドシートの持ち主・編集者でなければエラーにして止める。
 * （生徒の画面から先生用の関数を呼ばれても、名前や点数が見えないようにするため）
 */
function assertTeacher_() {
  const email = getViewerEmail_().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owner = ss.getOwner();
  const allowed = ss.getEditors().map(u => u.getEmail())
    .concat(owner ? [owner.getEmail()] : [])
    .map(e => String(e || '').toLowerCase());
  if (!email || allowed.indexOf(email) === -1) {
    throw new Error('この操作は、スプレッドシートの編集者（先生）だけが使えます。');
  }
}

/** 出席番号をそろえる（全角→半角、英字は大文字、前後の空白を取る）。例: '１ｂ１５' → '1B15' */
function normalizeStudentNumber_(value) {
  return String(value === undefined || value === null ? '' : value).normalize('NFKC').replace(/\s+/g, '').toUpperCase();
}

/** 日付が何年度か（4月始まり）。例: 2027年3月 → 2026 */
function fiscalYearOf_(dateValue) {
  const d = dateValue instanceof Date ? dateValue : new Date(dateValue);
  if (isNaN(d.getTime())) return null;
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const y = Number(Utilities.formatDate(d, tz, 'yyyy'));
  const m = Number(Utilities.formatDate(d, tz, 'M'));
  return m >= 4 ? y : y - 1;
}

/* =====================================================================
 * 生徒名簿（年度ごとに「生徒名簿_2026」のようなシートを1枚）
 * A列=出席番号（4桁）、B列=氏名、C列=ふりがな（任意）、D列=Googleアカウント（任意）、E列=メモ
 * ===================================================================== */

const ROSTER_SHEET_PREFIX = '生徒名簿_';
const ROSTER_HEADER = ['出席番号', '氏名', 'ふりがな（任意）', 'Googleアカウント（任意）', 'メモ'];

/**
 * 指定した年度の名簿を読む。
 * @return {{exists: boolean, sheetName: string, students: Array, problems: string[]}}
 *   students は出席番号順。problems は番号の書き方の誤り・重複など、先生に知らせたいこと
 */
function loadRoster_(year) {
  const sheetName = ROSTER_SHEET_PREFIX + year;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const result = { exists: !!sheet, sheetName: sheetName, students: [], problems: [] };
  if (!sheet || sheet.getLastRow() < 2) return result;

  const seen = {};
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues().forEach((r, i) => {
    const number = normalizeStudentNumber_(r[0]);
    const name = String(r[1] || '').trim();
    if (!number && !name) return;
    const rowNo = i + 2;
    const parsed = parseStudentNumber_(number);
    if (!parsed) {
      result.problems.push(`${rowNo}行目：出席番号「${r[0]}」が4桁の形になっていません（例：1215、1B15）`);
      return;
    }
    if (!name) {
      result.problems.push(`${rowNo}行目：${number} の氏名が空です`);
      return;
    }
    if (seen[number]) {
      result.problems.push(`${rowNo}行目：出席番号 ${number} が ${seen[number]}行目と重なっています`);
      return;
    }
    seen[number] = rowNo;
    result.students.push({
      number: number,
      name: name,
      kana: String(r[2] || '').trim(),
      email: String(r[3] || '').trim().toLowerCase(),
      grade: parsed.grade,
      klass: parsed.klass,
      no: parsed.number
    });
  });
  result.students.sort((a, b) => a.number < b.number ? -1 : a.number > b.number ? 1 : 0);
  return result;
}

function currentFiscalYear_() {
  return fiscalYearOf_(new Date());
}

/** メニュー用：今年度の名簿シートを開く（なければ作る）。入力の誤りがあれば知らせる */
function showRosterSheet() {
  assertTeacher_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const year = currentFiscalYear_();
  const sheetName = ROSTER_SHEET_PREFIX + year;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, ROSTER_HEADER.length).setValues([ROSTER_HEADER])
      .setFontWeight('bold').setBackground('#f0ece3');
    sheet.getRange('A:A').setNumberFormat('@'); // 「1B15」や先頭の0が数値・日付に化けないように
    sheet.setFrozenRows(1);
    [90, 160, 160, 260, 200].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
    sheet.activate();
    ui.alert(
      `「${sheetName}」シートを作りました（${year}年度の名簿）。\n\n` +
      '・A列に出席番号（4桁）、B列に氏名を貼り付けてください。Excel や校務システムの表からコピーできます。\n' +
      '　出席番号の書き方：高校1年2組15番 → 1215、中学1年B組15番 → 1B15\n' +
      '・D列に生徒のGoogleアカウント（メールアドレス）を入れておくと、生徒が開いたときに番号と名前が自動で入ります（空でもかまいません）。\n' +
      '・生徒は出席番号を入れるだけで、名前は名簿から入ります。レポートもこの名簿の番号順に並びます。\n' +
      '・貼り付けたら、もう一度このメニューを押すと、書き方の誤りがないか確かめられます。\n' +
      '・4月になると、新しい年度の名簿シートが作られます（前の年度のシートはそのまま残ります）。');
    return;
  }
  sheet.activate();
  const roster = loadRoster_(year);
  ui.alert(
    `「${sheetName}」：${roster.students.length}人が登録されています。\n\n` +
    (roster.problems.length
      ? '次の行を直してください（直すまで、その生徒は名簿にいないものとして扱われます）：\n・' +
        roster.problems.slice(0, 15).join('\n・') + (roster.problems.length > 15 ? `\n…ほか${roster.problems.length - 15}件` : '')
      : '書き方の誤りや番号の重なりはありませんでした。'));
}

/**
 * 生徒画面用：開いている生徒のGoogleアカウントが今年度の名簿にあれば、出席番号と氏名を返す。
 * 名簿が作られていなければ rosterActive=false（これまでどおり名前も手で入力する）。
 */
function whoAmI() {
  const roster = loadRoster_(currentFiscalYear_());
  const active = roster.students.length > 0;
  const email = getViewerEmail_().toLowerCase();
  const me = active && email ? roster.students.find(s => s.email && s.email === email) : null;
  return me
    ? { rosterActive: true, number: me.number, name: me.name }
    : { rosterActive: active };
}

/**
 * 生徒画面用：出席番号から名簿の氏名を返す。
 * 名簿にGoogleアカウントが書いてある生徒は、そのアカウントで開いたときだけ名前を返す
 * （他の生徒の番号を入れて名前を調べたり、なりすましたりできないように）。
 */
function lookupStudent(studentNumber) {
  const roster = loadRoster_(currentFiscalYear_());
  if (roster.students.length === 0) return { rosterActive: false };
  const number = normalizeStudentNumber_(studentNumber);
  const student = roster.students.find(s => s.number === number);
  if (!student) {
    return { rosterActive: true, ok: false, message: `出席番号 ${number} は名簿にありません。番号を確かめてください。` };
  }
  if (student.email && student.email !== getViewerEmail_().toLowerCase()) {
    return { rosterActive: true, ok: false, message: `出席番号 ${number} は、別のGoogleアカウントの生徒として登録されています。番号を確かめてください。` };
  }
  return { rosterActive: true, ok: true, number: number, name: student.name };
}

/**
 * Webアプリ版の生徒アプリから google.script.run 経由で呼ばれる送信用エンドポイント。
 * 名簿がある年度は、氏名を名簿の表記にそろえて保存する（入力ミスや表記ゆれで別人扱いにならないように）。
 */
function submitAssessmentToServer(data) {
  data.studentNumber = normalizeStudentNumber_(data.studentNumber);
  const year = fiscalYearOf_(data.date) || currentFiscalYear_();
  const student = loadRoster_(year).students.find(s => s.number === data.studentNumber);
  const email = getViewerEmail_();
  if (student && student.email && student.email !== email.toLowerCase()) {
    throw new Error(`出席番号 ${data.studentNumber} は、別のGoogleアカウントの生徒として名簿に登録されています。`);
  }
  if (student) data.studentName = student.name;
  appendResponseRow_(data, email);
  return { ok: true };
}

/**
 * 生徒用：自分のアカウントから送った記録だけを返す（生徒画面の「成長の経過」で、
 * 端末に残っていない過去の記録も比べられるようにするため）。
 * 他の生徒の記録は返さない。アカウントが分からないときは何も返さない。
 * ※google.script.run は日付(Date)を返せないため、日付は文字列にして返す。
 */
function getMyHistory() {
  const email = getViewerEmail_().toLowerCase();
  if (!email) return [];
  return loadResponses_()
    .filter(r => String(r.senderEmail).toLowerCase() === email)
    .map(r => ({
      id: Number(r.recordId) || String(r.recordId),
      date: r.date instanceof Date ? r.date.toISOString() : String(r.date || ''),
      studentNumber: r.studentNumber,
      title: String(r.title || ''),
      mode: String(r.mode || ''),
      scores: r.scores
    }))
    .filter(r => r.id && r.date);
}

/* =====================================================================
 * 実施回の一覧（生徒は自由入力ではなく、ここに登録した実施回から選ぶ）
 * 「実施回」シート：A列=実施回の名前、B列=受付中（チェックが入っている回だけ生徒画面に出る）
 * ===================================================================== */

const OCCASION_LIST_SHEET = '実施回';
const OCCASION_CACHE_KEY = 'openOccasions';

/** 生徒画面用：受付中の実施回の名前を、シートの上から順に返す（1件もなければ空配列＝自由入力） */
function getOpenOccasions() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(OCCASION_CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OCCASION_LIST_SHEET);
  let names = [];
  if (sheet && sheet.getLastRow() >= 2) {
    names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
      .filter(r => r[1] === true && String(r[0]).trim() !== '')
      .map(r => normalizeTitle_(r[0]));
  }
  // 生徒が一斉に開いてもシートを毎回読まないよう、1分だけ覚えておく
  cache.put(OCCASION_CACHE_KEY, JSON.stringify(names), 60);
  return names;
}

/** 「実施回」シートに登録されているすべての実施回の名前（受付中かどうかは問わない） */
function getRegisteredOccasionNames_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OCCASION_LIST_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => normalizeTitle_(r[0]));
}

/** 「実施回」シートが書きかえられたら、生徒画面にすぐ反映されるよう覚えておいた一覧を捨てる */
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === OCCASION_LIST_SHEET) {
    CacheService.getScriptCache().remove(OCCASION_CACHE_KEY);
  }
}

/** メニュー用：「実施回」シートを開く（なければ作る） */
function showOccasionListSheet() {
  assertTeacher_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(OCCASION_LIST_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OCCASION_LIST_SHEET);
    sheet.getRange(1, 1, 1, 3).setValues([['実施回の名前', '受付中', 'メモ']])
      .setFontWeight('bold').setBackground('#f0ece3');
    sheet.getRange(2, 2, 199, 1).insertCheckboxes(); // 先にチェックボックスを入れる（入れると値が外れるため）
    sheet.getRange(2, 1, 1, 2).setValues([['第1回 ふりかえり', true]]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 320);
    sheet.setColumnWidth(3, 320);
  }
  sheet.activate();
  CacheService.getScriptCache().remove(OCCASION_CACHE_KEY);
  SpreadsheetApp.getUi().alert(
    '「実施回」シートの使い方\n\n' +
    '・A列に実施回の名前を書きます（例：第2回 中間発表のふりかえり）。\n' +
    '・B列にチェックを入れた回だけが、生徒画面の選択肢に出ます。終わった回はチェックを外してください。\n' +
    '・チェックが1つもないときは、これまでどおり生徒がタイトルを自由に入力します。\n' +
    '・生徒画面には、開き直したときに反映されます。');
}

/**
 * 生徒がWebアプリのURLを開いたときに、生徒配布用HTML（Indexファイル）を返す。
 * ※Apps Scriptプロジェクトに「Index」という名前のHTMLファイルが必要。
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('探究的な学びを通じて身に付けたい資質・能力21 ふりかえりアプリ（総合的な探究の時間）')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * スプレッドシートを開いたときにメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('探究アセスメント')
    .addItem('🖨 レポートを作る・印刷する（クラス集計／推移／生徒の変容／個人の成長）', 'showReportCenter')
    .addSeparator()
    .addItem('👥 生徒名簿を設定する', 'showRosterSheet')
    .addItem('🗓 実施回を設定する', 'showOccasionListSheet')
    .addSeparator()
    .addItem('版を確認', 'showAppVersion')
    .addToUi();
}

/* =====================================================================
 * レポート印刷（クラス集計／学年・クラスの推移／クラスの生徒の変容／生徒個人の成長）
 * 画面（Report.html）でクラスや表示する値を選び、プレビューをそのまま印刷・PDF保存する。
 * 計算は画面側で行い、ここでは1年度分の回答と名簿をまとめて渡すだけにする。
 * ===================================================================== */

function showReportCenter() {
  assertTeacher_();
  const html = HtmlService.createHtmlOutputFromFile('Report').setWidth(1180).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'レポートを作る・印刷する');
}

/**
 * レポート画面用：指定した年度（省略時は今年度）の回答・名簿・実施回をまとめて返す。
 * 回答は点数だけを返し、振り返りの文章やメッセージは含めない。
 */
function getReportData(year) {
  assertTeacher_();
  const all = loadResponses_();
  const thisYear = currentFiscalYear_();
  const yearSet = new Set([thisYear]);
  all.forEach(r => { const y = fiscalYearOf_(r.date); if (y) yearSet.add(y); });
  const targetYear = Number(year) || thisYear;

  const inYear = all.filter(r => fiscalYearOf_(r.date) === targetYear);
  const roster = loadRoster_(targetYear);

  // 実施回の並び：「実施回」シートの上からの順 → 登録されていない回は実施日の早い順
  const registered = getRegisteredOccasionNames_();
  const dates = {};
  inYear.forEach(r => {
    const t = normalizeTitle_(r.title);
    const ms = new Date(r.date).getTime();
    if (!dates[t]) dates[t] = [];
    if (!isNaN(ms)) dates[t].push(ms);
  });
  const avg = list => list.length ? list.reduce((a, b) => a + b, 0) / list.length : 0;
  const titles = Object.keys(dates).sort((a, b) => {
    const ia = registered.indexOf(a), ib = registered.indexOf(b);
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return avg(dates[a]) - avg(dates[b]);
  });
  const titleIndex = {};
  titles.forEach((t, i) => { titleIndex[t] = i; });

  return {
    year: targetYear,
    years: Array.from(yearSet).sort((a, b) => b - a),
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy年M月d日'),
    competencies: COMPETENCIES,
    categoryLabel: CATEGORY_LABEL,
    roster: {
      exists: roster.exists,
      sheetName: roster.sheetName,
      problems: roster.problems,
      students: roster.students.map(s => ({ number: s.number, name: s.name, kana: s.kana }))
    },
    occasions: titles,
    responses: inYear.map(r => ({
      n: normalizeStudentNumber_(r.studentNumber),
      name: String(r.studentName || ''),
      t: titleIndex[normalizeTitle_(r.title)],
      d: new Date(r.date).getTime() || 0,
      m: String(r.mode || ''),
      s: COMPETENCIES.map(c => Number(r.scores[c.id]) || 0)
    }))
  };
}

function showAppVersion() {
  SpreadsheetApp.getUi().alert(
    'このスプレッドシートのプログラムの版：' + APP_VERSION + '\n\n' +
    '生徒画面のいちばん下に出る版と同じなら、貼り直しとデプロイは完了しています。\n' +
    '違うときは「デプロイ」→「デプロイを管理」→ 鉛筆マーク →「新バージョン」→「デプロイ」を押してください。');
}

/**
 * 出席番号（半角4桁）を 学年／組／番号 に分解する
 * 例: "1215" → { grade: '1', klass: '2', number: '15' }
 */
function parseStudentNumber_(studentNumber) {
  const s = String(studentNumber || '').trim().toUpperCase();
  // 1文字目=学年(数字)、2文字目=組(数字 or 英字)、3〜4文字目=番号(数字)
  if (!/^\d[A-Z0-9]\d{2}$/.test(s)) return null;
  const isMiddle = /[A-Z]/.test(s.charAt(1)); // 組が英字なら中学、数字なら高校
  return { grade: (isMiddle ? '中' : '高') + s.charAt(0), klass: s.charAt(1), number: s.substring(2) };
}

/**
 * 「回答」シートの全行を読み込み、パース済みレコードの配列を返す
 */
function loadResponses_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  // 古いシート（列が少ない）でも安全に読めるよう、実際の列数とRESPONSE_HEADERの小さい方に合わせる
  const colCount = Math.min(sheet.getLastColumn(), RESPONSE_HEADER.length);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();
  const records = [];
  values.forEach(row => {
    const [receivedAt, date, studentNumber, studentName, title, mode, scoresJson, reflection, refinedReflection,
      recordId, senderEmail, notesJson, , peerCount, peerAvgJson, peerEvalsJson] = row;
    let peerEvals = [];
    try { peerEvals = JSON.parse(peerEvalsJson || '[]') || []; } catch (e) { peerEvals = []; }
    records.push({
      receivedAt, date, studentNumber: normalizeStudentNumber_(studentNumber), studentName, title, mode,
      scores: parseJsonSafe_(scoresJson), reflection, refinedReflection, recordId,
      senderEmail: senderEmail || '',
      notes: parseJsonSafe_(notesJson),
      peerScores: parseJsonSafe_(peerAvgJson),
      peerEvals: peerEvals,
      parsed: parseStudentNumber_(studentNumber)
    });
  });
  return records;
}

/**
 * タイトル文字列を正規化する（全角英数字→半角、全角/半角スペースの統一、前後の空白除去）
 * 例: '第１回　ふりかえり ' → '第1回 ふりかえり'
 */
function normalizeTitle_(title) {
  const cleaned = String(title || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  return cleaned || '（タイトル未設定）';
}
