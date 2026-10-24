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
 * - 「個人推移」　: 生徒ごとに提出履歴を時系列で並べた一覧（同一生徒の複数回答を追跡できる）
 * - 「レポート_◯◯」: 実施回（タイトル or 日付）を選んで作成する、要約文＋グラフのレポート
 *
 * ▼レポート印刷（メニュー「レポートを作る・印刷する」）
 * クラス集計（クラスを選ぶ）／学年・クラスの推移／クラスの生徒の変容（名簿順の一覧表）を
 * 画面で選んでプレビューし、そのままA4で印刷・PDF保存する。年度ごと（4月〜3月）に集計する。
 *
 * ▼安全のための約束
 * 「実行ユーザー: 自分」で公開しているため、名前の末尾が「_」でない関数は生徒の画面からも呼べてしまう。
 * 先生用の関数は最初に assertTeacher_() を呼び、スプレッドシートの編集者以外は止める。
 * 内部だけで使う関数は、名前の末尾を「_」にする（外から呼べなくなる）。
 *
 * ▼集計に使う回答
 * 「5項目選択」「自己×他己比較」モードは生徒が自分で選んだ5項目だけを答えるため、平均が高めに出やすい。
 * そのためクラス集計・各レポートは「全21項目」モードの回答だけで計算する
 * （全21項目の回答が1件もないときに限り、5項目モードの回答で計算し、その旨をレポートに書く）。
 */

// 生徒画面（Index）の APP_VERSION と同じ値にそろえる。貼り直しの確認に使う
const APP_VERSION = '2026-10-24.2';

const SHEET_NAME = '回答';
const TREND_SHEET = '個人推移';

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
    .addItem('🖨 レポートを作る・印刷する（クラス集計／推移／生徒の変容）', 'showReportCenter')
    .addSeparator()
    .addItem('👥 生徒名簿を設定する', 'showRosterSheet')
    .addItem('🗓 実施回を設定する', 'showOccasionListSheet')
    .addSeparator()
    .addItem('わかりやすいレポートを作成（実施回ごと）', 'showOccasionReportDialog')
    .addItem('生徒個人の成長レポートを作成', 'showPersonalGrowthDialog')
    .addItem('「個人推移」シート（回答の一覧）を更新', 'updateTrendListSheet')
    .addSeparator()
    .addItem('版を確認', 'showAppVersion')
    .addToUi();
}

/** メニュー用：「個人推移」シート（1回答1行の一覧）を作り直す */
function updateTrendListSheet() {
  assertTeacher_();
  const records = loadResponses_();
  const ui = SpreadsheetApp.getUi();
  if (records.length === 0) {
    ui.alert('集計対象のデータがまだありません。「回答」シートにデータが蓄積されてから再実行してください。');
    return;
  }
  buildTrendSheet_(records);
  ui.alert(`「${TREND_SHEET}」シートを更新しました（${records.length}件）。`);
}

/* =====================================================================
 * レポート印刷（クラス集計／学年・クラスの推移／クラスの生徒の変容）
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
 * 平均の計算に使う回答を選ぶ。
 * 5項目モード（SELECT5/COMPARE）は生徒が自分で選んだ項目だけを答えるため、平均が高めに出やすい。
 * 全21項目モードの回答があればそれだけを使い、1件もないときだけ全回答を使う。
 * @return {{records: Array, fullOnly: boolean, note: string}} note はレポートに書き添える説明（不要なら空文字）
 */
function selectRecordsForAggregate_(records) {
  const full = records.filter(r => r.mode === 'ALL');
  if (full.length > 0) {
    const excluded = records.length - full.length;
    return {
      records: full,
      fullOnly: true,
      note: excluded > 0
        ? `※「全21項目」モードの回答だけで集計しています（5項目モードの回答${excluded}件は、生徒が自分で選んだ項目だけのため平均に含めていません）。`
        : ''
    };
  }
  return {
    records: records,
    fullOnly: false,
    note: records.length > 0
      ? '※「全21項目」モードの回答がないため、5項目モードの回答で集計しています。5項目は生徒が自分で選んでいるため、平均は高めに出やすい点にご注意ください。'
      : ''
  };
}

/**
 * 生徒ごとの提出履歴を時系列で並べた一覧を「個人推移」シートに書き出す
 */
function buildTrendSheet_(records) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TREND_SHEET);
  if (sheet) {
    const existingFilter = sheet.getFilter();
    if (existingFilter) existingFilter.remove();
    sheet.clear();
  } else {
    sheet = ss.insertSheet(TREND_SHEET);
  }

  const catIds = { Thinking: [], Communication: [], Leadership: [] };
  COMPETENCIES.forEach(c => catIds[c.category].push(c.id));

  const avgOf = (scores, ids) => {
    const vals = ids.map(id => Number(scores[id])).filter(v => v > 0);
    if (vals.length === 0) return '';
    return Number((vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(2));
  };

  const header = ['学年', '組', '出席番号', '氏名', '実施日', 'モード', '総合平均', 'Thinking平均', 'Communication平均', 'Leadership平均', 'タイトル'];
  const rows = [header];

  const sorted = records.slice().sort((a, b) => {
    const numA = a.studentNumber, numB = b.studentNumber;
    if (numA !== numB) return numA < numB ? -1 : 1;
    return new Date(a.date) - new Date(b.date);
  });

  sorted.forEach(r => {
    const allIds = Object.keys(r.scores).map(Number);
    rows.push([
      r.parsed ? r.parsed.grade : '',
      r.parsed ? r.parsed.klass : '',
      r.studentNumber,
      r.studentName,
      r.date ? new Date(r.date) : '',
      r.mode,
      avgOf(r.scores, allIds),
      avgOf(r.scores, catIds.Thinking),
      avgOf(r.scores, catIds.Communication),
      avgOf(r.scores, catIds.Leadership),
      r.title
    ]);
  });

  sheet.getRange(1, 1, rows.length, header.length).setValues(rows);
  sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f0ece3');
  sheet.getRange(2, 5, Math.max(rows.length - 1, 1), 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, 7, Math.max(rows.length - 1, 1), 4).setNumberFormat('0.00');
  sheet.setFrozenRows(1);
  autoFitTableColumns_(sheet, 1, header.length, 220);

  // 学年・組・氏名等で絞り込み／並べ替えできるよう、シート全体にフィルタを設定する
  if (rows.length > 1) {
    sheet.getRange(1, 1, rows.length, header.length).createFilter();
  }
}

/* =====================================================================
 * 「わかりやすいレポート」機能
 * ふりかえりの実施回を、①アセスメントタイトル or ②実施日 のどちらかで
 * 絞り込み、専門知識がなくても一目でわかる文章＋グラフのレポートシートを生成する。
 * 個人名・自由記述コメントは含めず、集計値のみを扱う（プライバシー配慮）。
 *
 * ▼タイトルの表記ゆれ対策について
 * 生徒が自由入力するタイトルは、余分な空白・全角半角の違い・入力ミスなどで
 * 本来同じ回のはずが別々の文字列として記録されることがある。これに対し：
 *  1. normalizeTitle_() で空白の統一・前後trimを自動的に行い、些細な表記ゆれは吸収する
 *  2. それでも異なる文字列になったものは、レーベンシュタイン距離が近いタイトル同士を
 *     「似ている候補」としてグループ表示し、チェックボックスで複数選択・統合できる
 *  3. さらに、タイトルに頼らず「実施日」を基準に選ぶ方法も用意した（本セクション）。
 *     近い日付（3日以内）は自動的にひとまとまりで提示され、離れた日付も
 *     チェックボックスで自由に組み合わせて1つのレポートに集計できる。
 * 根本対策として、「実施回」シートに実施回を登録すると、生徒はタイトルを入力せずに一覧から選ぶ
 * （showOccasionListSheet / getOpenOccasions）。表記ゆれが起きるのは、登録前の回答と自由入力の回答だけになる。
 * ===================================================================== */

const OCCASION_SHEET_PREFIX = 'レポート_';

/**
 * タイトル文字列を正規化する（全角英数字→半角、全角/半角スペースの統一、前後の空白除去）
 * 例: '第１回　ふりかえり ' → '第1回 ふりかえり'
 */
function normalizeTitle_(title) {
  const cleaned = String(title || '').normalize('NFKC').replace(/\s+/g, ' ').trim();
  return cleaned || '（タイトル未設定）';
}

/**
 * レーベンシュタイン距離（編集距離）を計算する
 */
function levenshteinDistance_(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = [];
  for (let i = 0; i <= m; i++) dp.push(new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      dp[i][j] = a.charAt(i - 1) === b.charAt(j - 1)
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
    }
  }
  return dp[m][n];
}

/**
 * 正規化済みタイトルの一覧を、表記ゆれの可能性がある者同士でグルーピングする
 * @param {Array<{title:string, count:number}>} list
 * @param {Set<string>} [keepApart] 「実施回」シートに登録済みの名前。登録済みの名前どうしは
 *        （「第1回」と「第2回」のように似ていても）別の回なので、同じグループにしない
 * @return {Array<{variants: Array<{title:string, count:number}>, totalCount: number}>}
 */
function clusterTitles_(list, keepApart) {
  const items = list.slice().sort((a, b) => b.count - a.count);
  const used = new Array(items.length).fill(false);
  const clusters = [];

  for (let i = 0; i < items.length; i++) {
    if (used[i]) continue;
    used[i] = true;
    const cluster = { variants: [items[i]], totalCount: items[i].count };
    for (let j = i + 1; j < items.length; j++) {
      if (used[j]) continue;
      const a = items[i].title, b = items[j].title;
      if (keepApart && keepApart.has(a) && keepApart.has(b)) continue;
      const dist = levenshteinDistance_(a, b);
      const threshold = Math.max(2, Math.floor(Math.max(a.length, b.length) * 0.2));
      if (dist > 0 && dist <= threshold) {
        cluster.variants.push(items[j]);
        cluster.totalCount += items[j].count;
        used[j] = true;
      }
    }
    clusters.push(cluster);
  }

  clusters.sort((a, b) => b.totalCount - a.totalCount);
  return clusters;
}

/**
 * 「回答」シート内に存在する実施回（タイトル）ごとの件数一覧を、表記ゆれグループ化して返す
 * （ダイアログの選択肢を組み立てるため、クライアント側から呼ばれる）
 */
function getDistinctTitlesForDialog() {
  assertTeacher_();
  const records = loadResponses_();
  const counts = {};
  records.forEach(r => {
    const t = normalizeTitle_(r.title);
    counts[t] = (counts[t] || 0) + 1;
  });
  const list = Object.keys(counts).map(t => ({ title: t, count: counts[t] }));
  return clusterTitles_(list, new Set(getRegisteredOccasionNames_()));
}

/**
 * 日付文字列（実施日）を 'yyyy-MM-dd' 形式に変換する。無効な日付はnullを返す
 */
function toDateKey_(dateValue) {
  const d = new Date(dateValue);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd');
}

/**
 * 2つの日付キー（yyyy-MM-dd）の間隔（日数）を計算する
 */
function daysBetween_(dateKeyA, dateKeyB) {
  const msPerDay = 24 * 60 * 60 * 1000;
  return Math.abs((new Date(dateKeyA) - new Date(dateKeyB)) / msPerDay);
}

/**
 * 「回答」シート内に存在する実施日ごとの件数一覧を、近い日付（3日以内）でグループ化して返す
 */
function getDistinctDatesForDialog() {
  assertTeacher_();
  const records = loadResponses_();
  const counts = {};
  records.forEach(r => {
    const key = toDateKey_(r.date);
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });
  const list = Object.keys(counts)
    .map(k => ({ title: k, count: counts[k] }))
    .sort((a, b) => a.title < b.title ? -1 : 1); // 日付の昇順

  // 日付順に並んでいる前提で、クラスタの起点（最初の日付）から3日以内なら同じクラスタにまとめる
  // （直前の1件とだけ比較する方式だと、3日おきの日付が連鎖して際限なく1クラスタに広がってしまうため、
  //   起点との距離で判定することでクラスタの範囲を「最大3日間」に確実に収める）
  const clusters = [];
  let current = null;
  list.forEach(item => {
    const anchor = current ? current.variants[0] : null;
    if (anchor && daysBetween_(anchor.title, item.title) <= 3) {
      current.variants.push(item);
      current.totalCount += item.count;
    } else {
      current = { variants: [item], totalCount: item.count };
      clusters.push(current);
    }
  });
  return clusters;
}

/**
 * メニューから起動：実施回を選ぶダイアログを表示する（タイトルで選ぶ／日付で選ぶ の2方式）
 */
function showOccasionReportDialog() {
  assertTeacher_();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: sans-serif; font-size: 13px; padding: 4px 10px 10px; color: #2a2420; }
      .tabs { display: flex; gap: 6px; margin-bottom: 10px; }
      .tab-btn { flex: 1; padding: 8px; font-size: 12px; font-weight: bold; border: 1px solid #d8cfb8;
                 border-radius: 6px; background: #f5f0e8; color: #7a6f5c; cursor: pointer; }
      .tab-btn.active { background: #2c5f8a; color: white; border-color: #2c5f8a; }
      .panel { display: none; }
      .panel.active { display: block; }
      .group { border: 1px solid #e6dfd0; border-radius: 8px; padding: 8px 10px; margin-bottom: 8px; background: #fdfbf6; }
      .group.warn { border-color: #e0b060; background: #fff8ec; }
      .group-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 4px; }
      .group-title { font-weight: bold; font-size: 12px; }
      .warn-badge { font-size: 10px; color: #a8402c; font-weight: bold; }
      .variant-row { display: flex; align-items: center; gap: 6px; padding: 2px 0; font-size: 12px; }
      .link-btn { font-size: 11px; color: #2c5f8a; text-decoration: underline; cursor: pointer; background: none; border: none; padding: 0; }
      #reportName { width: 100%; padding: 6px; font-size: 13px; margin: 6px 0 12px; box-sizing: border-box; }
      .groups-box { max-height: 200px; overflow-y: auto; margin: 8px 0; }
      button.primary { padding: 8px 18px; font-size: 13px; font-weight: bold; border: none; border-radius: 6px;
               background: #2c5f8a; color: white; cursor: pointer; }
      button.primary:disabled { background: #aaa; cursor: default; }
      #msg { margin-top: 10px; font-size: 12px; font-weight: bold; }
      .empty { color: #7a6f5c; display: none; }
    </style>
    <div>
      <div class="tabs">
        <button class="tab-btn active" id="tabTitleBtn" onclick="switchTab('title')">実施回タイトルで選ぶ</button>
        <button class="tab-btn" id="tabDateBtn" onclick="switchTab('date')">実施日で選ぶ</button>
      </div>

      <div class="panel active" id="panelTitle">
        入力ミス等で表記が分かれているものは「似ています」の枠にまとめて表示されます。
        <div id="loadingTitle">読み込み中...</div>
        <div class="groups-box" id="groupsTitle"></div>
        <div class="empty" id="emptyTitle">まだ集計できる回答データがありません。</div>
      </div>

      <div class="panel" id="panelDate">
        近い日付（3日以内）は自動的にまとめて表示されます。チェックボックスで自由に組み合わせて選べます。
        <div id="loadingDate">読み込み中...</div>
        <div class="groups-box" id="groupsDate"></div>
        <div class="empty" id="emptyDate">まだ集計できる回答データがありません。</div>
      </div>

      <div id="form" style="display:none;">
        <label style="font-size:11px;color:#7a6f5c;">レポートの表示タイトル（編集可）</label>
        <input type="text" id="reportName">
        <button class="primary" id="runBtn" onclick="runReport()">選択したデータでレポートを作成</button>
      </div>
      <div id="msg"></div>
    </div>
    <script>
      var activeTab = 'title';

      function switchTab(tab) {
        activeTab = tab;
        document.getElementById('tabTitleBtn').className = 'tab-btn' + (tab === 'title' ? ' active' : '');
        document.getElementById('tabDateBtn').className = 'tab-btn' + (tab === 'date' ? ' active' : '');
        document.getElementById('panelTitle').className = 'panel' + (tab === 'title' ? ' active' : '');
        document.getElementById('panelDate').className = 'panel' + (tab === 'date' ? ' active' : '');
        updateReportNameSuggestion();
      }

      function renderGroups(list, containerId, loadingId, emptyId, labelFn) {
        document.getElementById(loadingId).style.display = 'none';
        if (!list || list.length === 0) {
          document.getElementById(emptyId).style.display = 'block';
          return;
        }
        const container = document.getElementById(containerId);
        list.forEach(function(cluster, gi) {
          const box = document.createElement('div');
          box.className = 'group' + (cluster.variants.length > 1 ? ' warn' : '');
          const head = document.createElement('div');
          head.className = 'group-head';
          const titleSpan = document.createElement('span');
          titleSpan.className = 'group-title';
          titleSpan.textContent = labelFn(cluster.variants[0].title) + '（計' + cluster.totalCount + '件）';
          head.appendChild(titleSpan);
          if (cluster.variants.length > 1) {
            const warnSpan = document.createElement('span');
            warnSpan.className = 'warn-badge';
            warnSpan.textContent = containerId === 'groupsDate' ? '近い日付をまとめて表示中' : '⚠ 表記ゆれの可能性';
            head.appendChild(warnSpan);
          }
          box.appendChild(head);
          cluster.variants.forEach(function(v, vi) {
            const row = document.createElement('div');
            row.className = 'variant-row';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = v.title;
            cb.id = containerId + '_cb_' + gi + '_' + vi;
            cb.className = containerId + '-checkbox';
            cb.onchange = updateReportNameSuggestion;
            row.appendChild(cb);
            const label = document.createElement('label');
            label.setAttribute('for', cb.id);
            label.textContent = labelFn(v.title) + '（' + v.count + '件）';
            row.appendChild(label);
            box.appendChild(row);
          });
          if (cluster.variants.length > 1) {
            const allBtn = document.createElement('button');
            allBtn.className = 'link-btn';
            allBtn.textContent = 'この' + cluster.variants.length + '件をまとめて選択';
            allBtn.onclick = function() {
              cluster.variants.forEach(function(v, vi) {
                document.getElementById(containerId + '_cb_' + gi + '_' + vi).checked = true;
              });
              updateReportNameSuggestion();
            };
            box.appendChild(allBtn);
          }
          container.appendChild(box);
        });
        document.getElementById('form').style.display = 'block';
      }

      google.script.run.withSuccessHandler(function(list) {
        renderGroups(list, 'groupsTitle', 'loadingTitle', 'emptyTitle', function(t) { return t; });
      }).getDistinctTitlesForDialog();

      google.script.run.withSuccessHandler(function(list) {
        renderGroups(list, 'groupsDate', 'loadingDate', 'emptyDate', function(t) { return formatDateLabelClient(t); });
      }).getDistinctDatesForDialog();

      function formatDateLabelClient(dateKey) {
        const parts = dateKey.split('-');
        if (parts.length !== 3) return dateKey;
        return parts[1] + '月' + parts[2] + '日';
      }

      function updateReportNameSuggestion() {
        const selector = activeTab === 'title' ? '.groupsTitle-checkbox:checked' : '.groupsDate-checkbox:checked';
        const checked = Array.prototype.slice.call(document.querySelectorAll(selector));
        if (checked.length > 0) {
          document.getElementById('reportName').value = activeTab === 'title'
            ? checked[0].value
            : formatDateLabelClient(checked[0].value) + '実施分';
        }
      }

      function runReport() {
        const selector = activeTab === 'title' ? '.groupsTitle-checkbox:checked' : '.groupsDate-checkbox:checked';
        const checked = Array.prototype.slice.call(document.querySelectorAll(selector)).map(function(cb) { return cb.value; });
        const msgEl = document.getElementById('msg');
        if (checked.length === 0) {
          msgEl.style.color = '#a8402c';
          msgEl.textContent = '少なくとも1つは選択してください。';
          return;
        }
        const reportName = document.getElementById('reportName').value || checked[0];
        const btn = document.getElementById('runBtn');
        btn.disabled = true;
        btn.textContent = '作成中...';
        msgEl.textContent = '';
        google.script.run
          .withSuccessHandler(function(sheetName) {
            msgEl.style.color = '#2c5f8a';
            msgEl.textContent = '「' + sheetName + '」シートを作成しました。このダイアログを閉じてご確認ください。';
            btn.disabled = false;
            btn.textContent = '選択したデータでレポートを作成';
          })
          .withFailureHandler(function(err) {
            msgEl.style.color = '#a8402c';
            msgEl.textContent = 'エラーが発生しました: ' + err.message;
            btn.disabled = false;
            btn.textContent = '選択したデータでレポートを作成';
          })
          .generateOccasionReport({ mode: activeTab, values: checked }, reportName);
      }
    </script>
  `).setWidth(480).setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(html, 'わかりやすいレポートを作成');
}

/**
 * 指定した列範囲を内容に合わせて自動調整する共通ヘルパー。
 * 項目名や日付など列ごとに内容の長さが大きく異なると、列幅がバラバラで見づらくなるため、
 * 生成する全シートでこの関数を通して幅を揃える。
 * @param {Sheet} sheet 対象シート
 * @param {number} startCol 開始列（1始まり）
 * @param {number} numCols 列数
 * @param {number} [maxWidthPx] この幅を超える列は切り詰める（長い項目名等で極端に広がるのを防ぐ）
 */
function autoFitTableColumns_(sheet, startCol, numCols, maxWidthPx) {
  sheet.autoResizeColumns(startCol, numCols);
  if (maxWidthPx) {
    for (let c = startCol; c < startCol + numCols; c++) {
      if (sheet.getColumnWidth(c) > maxWidthPx) sheet.setColumnWidth(c, maxWidthPx);
    }
  }
}

/**
 * シート名として使えない文字を除去し、長さも制限する
 * @param {string} name 表示名のもとになる文字列
 * @param {string} [prefix] シート名の接頭辞（省略時は実施回レポート用のOCCASION_SHEET_PREFIX）
 */
function sanitizeSheetName_(name, prefix) {
  const cleaned = String(name || '').replace(/[:\\\/\?\*\[\]]/g, '　').trim();
  const base = (prefix !== undefined ? prefix : OCCASION_SHEET_PREFIX) + (cleaned || '無題');
  return base.length > 95 ? base.substring(0, 95) : base;
}

/**
 * データテーブルをシートに書き込む。カテゴリーごとに列を分けることで、
 * Google純正グラフの「1系列＝1色」制約を利用し、棒グラフをカテゴリー別の色分け表示にする。
 * @return {number} 使用した列数（常に4列: 項目名 + Thinking + Communication + Leadership）
 */
function writeCategorySeriesTable_(sheet, headerRow, col, items) {
  const header = ['項目', CATEGORY_LABEL.Thinking, CATEGORY_LABEL.Communication, CATEGORY_LABEL.Leadership];
  sheet.getRange(headerRow, col, 1, 4).setValues([header]);
  items.forEach((it, i) => {
    const row = [it.label, '', '', ''];
    const colIndex = { Thinking: 1, Communication: 2, Leadership: 3 }[it.category];
    row[colIndex] = Number(it.value.toFixed(2));
    sheet.getRange(headerRow + 1 + i, col, 1, 4).setValues([row]);
  });
  return 4;
}

/**
 * 選択された実施回のデータだけを集計し、文章＋グラフのレポートシートを作る
 * @param {{mode: 'title'|'date', values: string[]}} criteria 絞り込み条件
 *        mode='title': valuesはnormalizeTitle()済みのタイトル文字列の配列
 *        mode='date' : valuesは'yyyy-MM-dd'形式の日付文字列の配列
 * @param {string} reportName レポートの表示名・シート名のもとになる文字列
 * @return {string} 生成したシート名（ダイアログ側に返す）
 */
function generateOccasionReport(criteria, reportName) {
  assertTeacher_();
  const mode = (criteria && criteria.mode) || 'title';
  const targetValues = ((criteria && criteria.values) || []).filter(v => v);
  if (targetValues.length === 0) {
    throw new Error('対象の実施回が選択されていません。');
  }

  const loadedRecords = loadResponses_();
  const occasionRecords = loadedRecords.filter(r => {
    if (mode === 'date') {
      const key = toDateKey_(r.date);
      return key !== null && targetValues.indexOf(key) !== -1;
    }
    return targetValues.indexOf(normalizeTitle_(r.title)) !== -1;
  });

  if (occasionRecords.length === 0) {
    throw new Error('選択した実施回のデータが見つかりませんでした。');
  }

  // 今回分も、比べる相手の「全期間平均」も、同じモードの回答どうしで計算する
  const target = selectRecordsForAggregate_(occasionRecords);
  const records = target.records;
  const allRecords = target.fullOnly ? loadedRecords.filter(r => r.mode === 'ALL') : loadedRecords;

  const displayTitle = (reportName || targetValues[0]).trim() || targetValues[0];

  // ── 項目別・カテゴリー別の平均を集計 ──
  const itemAgg = {}; // compId -> {sum, count}
  records.forEach(r => {
    Object.keys(r.scores).forEach(compIdStr => {
      const compId = parseInt(compIdStr, 10);
      const score = Number(r.scores[compIdStr]);
      if (!score) return;
      if (!itemAgg[compId]) itemAgg[compId] = { sum: 0, count: 0 };
      itemAgg[compId].sum += score;
      itemAgg[compId].count += 1;
    });
  });

  const itemStats = COMPETENCIES
    .filter(c => itemAgg[c.id])
    .map(c => ({
      id: c.id,
      name: c.name,
      category: c.category,
      avg: itemAgg[c.id].sum / itemAgg[c.id].count,
      count: itemAgg[c.id].count
    }));

  if (itemStats.length === 0) {
    throw new Error('集計できる項目スコアが見つかりませんでした。');
  }

  const categoryStats = ['Thinking', 'Communication', 'Leadership'].map(cat => {
    const items = itemStats.filter(it => it.category === cat);
    const avg = items.length ? items.reduce((s, it) => s + it.avg, 0) / items.length : null;
    return { category: cat, label: CATEGORY_LABEL[cat], avg };
  }).filter(c => c.avg !== null);

  // 回答者ごとの総合平均（レベル分布用）
  const respondentAverages = records.map(r => {
    const ids = Object.keys(r.scores).map(Number);
    const vals = ids.map(id => Number(r.scores[id])).filter(v => v > 0);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  }).filter(v => v !== null);

  // 上限は「その値未満」で判定し、最後のバケットだけ4.0ちょうどを含める（境界での漏れ・二重計上を防ぐ）
  const levelBuckets = [
    { label: 'Level 1台（1.0〜1.9）', min: 1.0, max: 2.0 },
    { label: 'Level 2台（2.0〜2.9）', min: 2.0, max: 3.0 },
    { label: 'Level 3台（3.0〜3.9）', min: 3.0, max: 4.0 },
    { label: 'Level 4（4.0）',       min: 4.0, max: Infinity }
  ];
  const levelCounts = levelBuckets
    .map(b => ({ label: b.label, count: respondentAverages.filter(v => v >= b.min && v < b.max).length }))
    .filter(l => l.count > 0); // 0件のレベル帯は円グラフに出さない

  const overallAvg = respondentAverages.length
    ? respondentAverages.reduce((a, b) => a + b, 0) / respondentAverages.length
    : 0;

  // ── 全期間平均（ベースライン）の算出（今回と比較するため、全「回答」データから計算）──
  const baselineItemAgg = {};
  allRecords.forEach(r => {
    Object.keys(r.scores).forEach(compIdStr => {
      const compId = parseInt(compIdStr, 10);
      const score = Number(r.scores[compIdStr]);
      if (!score) return;
      if (!baselineItemAgg[compId]) baselineItemAgg[compId] = { sum: 0, count: 0 };
      baselineItemAgg[compId].sum += score;
      baselineItemAgg[compId].count += 1;
    });
  });
  const baselineCategoryStats = ['Thinking', 'Communication', 'Leadership'].map(cat => {
    const ids = COMPETENCIES.filter(c => c.category === cat).map(c => c.id);
    const avgs = ids.filter(id => baselineItemAgg[id]).map(id => baselineItemAgg[id].sum / baselineItemAgg[id].count);
    return { category: cat, label: CATEGORY_LABEL[cat], avg: avgs.length ? avgs.reduce((a, b) => a + b, 0) / avgs.length : null };
  });
  const baselineValidCats = baselineCategoryStats.filter(c => c.avg !== null);
  const baselineOverallAvg = baselineValidCats.length
    ? baselineValidCats.reduce((s, c) => s + c.avg, 0) / baselineValidCats.length
    : null;
  const hasComparableBaseline = baselineOverallAvg !== null && allRecords.length > records.length;

  // ── 文章（ひとことまとめ）の自動生成 ──
  const sortedByAvgDesc = itemStats.slice().sort((a, b) => b.avg - a.avg);
  const top3 = sortedByAvgDesc.slice(0, 3);
  const topIds = new Set(top3.map(it => it.id));
  // top3と重複しないよう除外してから下位3件を選ぶ（SELECT5/COMPAREモードは対象が5項目のため、
  // 単純にslice(-3)すると上位・下位の両方に同じ項目が含まれてしまうことがあるため）
  const bottom3 = sortedByAvgDesc.filter(it => !topIds.has(it.id)).slice(-3).reverse();

  const catSortedDesc = categoryStats.slice().sort((a, b) => b.avg - a.avg);
  const strongestCat = catSortedDesc[0];
  const weakestCat = catSortedDesc[catSortedDesc.length - 1];

  const level3PlusCount = respondentAverages.filter(v => v >= 3).length;
  const level3PlusPct = respondentAverages.length ? Math.round((level3PlusCount / respondentAverages.length) * 100) : 0;

  const summarySentences = [];
  summarySentences.push(
    `「${displayTitle}」では、${records.length}人の生徒が自己評価に回答しました。全体の平均レベルは4段階中${overallAvg.toFixed(1)}で、そのうち${level3PlusPct}%にあたる${level3PlusCount}人がレベル3以上に到達しています。`
  );
  if (strongestCat && weakestCat && strongestCat.category !== weakestCat.category) {
    summarySentences.push(
      `カテゴリー別に見ると「${strongestCat.label}」が最も高く（平均${strongestCat.avg.toFixed(1)}）、クラス全体としての強みと言えます。一方「${weakestCat.label}」は平均${weakestCat.avg.toFixed(1)}にとどまっており、次の探究活動で重点的に声かけしたい領域です。`
    );
  }
  if (top3.length > 0) {
    const topLine = top3.map(it => `「${it.name}」（平均${it.avg.toFixed(1)}）`).join('、');
    summarySentences.push(`個別項目では${topLine}の評価が特に高く、多くの生徒がこれらの力を発揮できたと感じているようです。`);
  }
  if (bottom3.length > 0) {
    const bottomLine = bottom3.map(it => `「${it.name}」（平均${it.avg.toFixed(1)}）`).join('、');
    summarySentences.push(`反対に${bottomLine}は他の項目と比べて評価が低めでした。次回の探究活動では、これらを意識した声かけや振り返りの機会を設けると効果的です。`);
  }
  if (hasComparableBaseline) {
    const diff = overallAvg - baselineOverallAvg;
    if (Math.abs(diff) < 0.05) {
      summarySentences.push(`これまでの全期間平均（${baselineOverallAvg.toFixed(1)}）とほぼ同水準の結果でした。`);
    } else if (diff > 0) {
      summarySentences.push(`これまでの全期間平均（${baselineOverallAvg.toFixed(1)}）と比べると、今回は${diff.toFixed(1)}ポイント高く、成長が感じられる結果です。`);
    } else {
      summarySentences.push(`これまでの全期間平均（${baselineOverallAvg.toFixed(1)}）と比べると、今回は${Math.abs(diff).toFixed(1)}ポイント低めでした。実施のタイミングや設問の受け止め方など、背景も合わせて確認してみてください。`);
    }
  }
  if (targetValues.length > 1) {
    summarySentences.push(
      mode === 'date'
        ? `（${targetValues.length}日分のデータをまとめて集計しています）`
        : `（表記ゆれの可能性がある${targetValues.length}種類のタイトルをまとめて集計しています）`
    );
  }
  if (target.note) summarySentences.push(target.note);
  summarySentences.push('一人ひとりの成長のペースは様々です。この結果を、次の探究活動に向けた生徒との対話のきっかけとしてご活用ください。');

  // ── シート作成 ──
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName_(displayTitle);
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // 見出し・サマリー文章
  sheet.getRange('A1').setValue(`『${displayTitle}』ふりかえりレポート`)
    .setFontSize(18).setFontWeight('bold').setFontColor('#2a2420');
  sheet.getRange('A2').setValue(`回答者数: ${records.length}人　／　作成日: ${Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy年M月d日')}`)
    .setFontSize(10).setFontColor('#7a6f5c');

  sheet.getRange('A4').setValue('ひとことまとめ').setFontSize(13).setFontWeight('bold').setFontColor('#a8402c');
  sheet.getRange('A5').setValue(summarySentences.join('\n\n')).setFontSize(12).setWrap(true);
  sheet.setRowHeight(5, 200);
  sheet.getRange('A5:J5').merge();

  // ── グラフ用データ（右側の作業エリアに書き出し、あとで列を隠す）──
  const WORK_COL = 14; // N列以降を作業エリアとして使用
  const WORK_COL_WIDTH = 4; // writeCategorySeriesTableが使う列数

  // カテゴリー別平均データ（カテゴリー別に列を分け、棒の色をカテゴリーごとに変える）
  const catHeaderRow = 1;
  const catItems = categoryStats.map(c => ({ label: c.label, category: c.category, value: c.avg }));
  writeCategorySeriesTable_(sheet, catHeaderRow, WORK_COL, catItems);

  // 項目別平均データ（高い順、カテゴリー別に色分け）
  const itemHeaderRow = catHeaderRow + catItems.length + 3;
  const itemItems = sortedByAvgDesc.map(it => ({ label: it.name, category: it.category, value: it.avg }));
  writeCategorySeriesTable_(sheet, itemHeaderRow, WORK_COL, itemItems);

  // レベル分布データ
  const levelHeaderRow = itemHeaderRow + itemItems.length + 3;
  sheet.getRange(levelHeaderRow, WORK_COL, 1, 2).setValues([['レベル帯', '人数']]);
  levelCounts.forEach((l, i) => {
    sheet.getRange(levelHeaderRow + 1 + i, WORK_COL, 1, 2).setValues([[l.label, l.count]]);
  });

  // 「今回 vs 全期間平均」比較データ（全期間データが今回以外にも存在する場合のみ）
  const compareHeaderRow = levelHeaderRow + levelCounts.length + 3;
  if (hasComparableBaseline) {
    const compareRows = [['カテゴリー', '今回', '全期間平均']];
    ['Thinking', 'Communication', 'Leadership'].forEach(cat => {
      const current = categoryStats.find(c => c.category === cat);
      const baseline = baselineCategoryStats.find(c => c.category === cat);
      compareRows.push([
        CATEGORY_LABEL[cat],
        current ? Number(current.avg.toFixed(2)) : null,
        (baseline && baseline.avg !== null) ? Number(baseline.avg.toFixed(2)) : null
      ]);
    });
    sheet.getRange(compareHeaderRow, WORK_COL, compareRows.length, 3).setValues(compareRows);
  }

  const seriesColors = [CATEGORY_COLOR.Thinking, CATEGORY_COLOR.Communication, CATEGORY_COLOR.Leadership];

  // ── グラフ挿入（縦に並べ、それぞれ十分な余白を確保して見やすくする）──
  const catChart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange(catHeaderRow, WORK_COL, catItems.length + 1, WORK_COL_WIDTH))
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(7, 1, 0, 0)
    .setOption('title', 'カテゴリー別 平均レベル')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 11 } })
    .setOption('vAxis', { minValue: 0, maxValue: 4, gridlines: { count: 5 }, title: '平均レベル（4段階）' })
    .setOption('hAxis', { textStyle: { fontSize: 11 } })
    .setOption('colors', seriesColors)
    .setOption('bar', { groupWidth: '55%' })
    .setOption('width', 620).setOption('height', 340)
    .build();
  sheet.insertChart(catChart);

  const itemChart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(itemHeaderRow, WORK_COL, itemItems.length + 1, WORK_COL_WIDTH))
    .setNumHeaders(1)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(25, 1, 0, 0)
    .setOption('title', '項目別 平均レベル（高い順）')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('legend', { position: 'top', textStyle: { fontSize: 11 } })
    .setOption('hAxis', { minValue: 0, maxValue: 4, gridlines: { count: 5 }, title: '平均レベル（4段階）' })
    .setOption('vAxis', { textStyle: { fontSize: 10 } })
    .setOption('colors', seriesColors)
    .setOption('bar', { groupWidth: '70%' })
    .setOption('width', 680).setOption('height', 620)
    .build();
  sheet.insertChart(itemChart);

  const levelChart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sheet.getRange(levelHeaderRow, WORK_COL, levelCounts.length + 1, 2))
    .setNumHeaders(1)
    .setPosition(60, 1, 0, 0)
    .setOption('title', '総合レベル帯の分布（回答者数の割合）')
    .setOption('titleTextStyle', { fontSize: 14, bold: true })
    .setOption('pieSliceText', 'value-and-percentage')
    .setOption('legend', { position: 'right', textStyle: { fontSize: 11 } })
    .setOption('colors', ['#e2dccb', '#8fb8d8', '#3d7a5c', '#a8402c'])
    .setOption('width', 560).setOption('height', 360)
    .build();
  sheet.insertChart(levelChart);

  if (hasComparableBaseline) {
    const compareChart = sheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(sheet.getRange(compareHeaderRow, WORK_COL, 4, 3))
      .setNumHeaders(1)
      .setOption('useFirstColumnAsDomain', true)
      .setPosition(95, 1, 0, 0)
      .setOption('title', '今回 vs 全期間平均')
      .setOption('titleTextStyle', { fontSize: 14, bold: true })
      .setOption('legend', { position: 'top' })
      .setOption('vAxis', { minValue: 0, maxValue: 4, gridlines: { count: 5 }, title: '平均レベル（4段階）' })
      .setOption('colors', ['#2c5f8a', '#c9bd9c'])
      .setOption('bar', { groupWidth: '55%' })
      .setOption('width', 620).setOption('height', 340)
      .build();
    sheet.insertChart(compareChart);
  }

  // 作業用データ列は非表示にする（各テーブルはWORK_COL_WIDTH=4列以内に収まっている）
  sheet.hideColumns(WORK_COL, WORK_COL_WIDTH);

  sheet.setColumnWidth(1, 700);
  sheet.activate();

  return sheetName;
}

/* =====================================================================
 * 「生徒個人の成長レポート」機能
 * 出席番号を指定して、その生徒のこれまでの全提出履歴を時系列にたどり、
 * 総合平均・カテゴリー別平均の推移グラフと文章による成長の要約を生成する。
 * さらに、生成したシートをそのままPDF化してDriveに保存し、リンクを返す。
 * ===================================================================== */

const GROWTH_SHEET_PREFIX = '成長記録_'; // PDFファイル名の接頭辞として使用
const GROWTH_WORK_SHEET_NAME = '成長レポート（作業用）'; // 生徒ごとにシートを増やさず、毎回この1枚を作り直す

/**
 * 「回答」シートに存在する生徒（出席番号）の一覧を、提出回数つきで返す
 * （ダイアログの選択肢を組み立てるため、クライアント側から呼ばれる）
 */
function getStudentsForDialog() {
  assertTeacher_();
  const records = loadResponses_();
  const byStudent = {};
  records.forEach(r => {
    const key = r.studentNumber || '';
    if (!key) return;
    if (!byStudent[key]) byStudent[key] = { studentNumber: key, studentName: r.studentName || '', count: 0 };
    byStudent[key].count += 1;
    if (r.studentName) byStudent[key].studentName = r.studentName; // 最新の表記を優先
  });
  return Object.values(byStudent).sort((a, b) => a.studentNumber < b.studentNumber ? -1 : 1);
}

/**
 * メニューから起動：成長レポートを作りたい生徒を選ぶダイアログを表示する
 */
function showPersonalGrowthDialog() {
  assertTeacher_();
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: sans-serif; font-size: 13px; padding: 4px 10px 10px; color: #2a2420; }
      .student-row { display: flex; align-items: center; gap: 8px; padding: 6px 8px; border: 1px solid #e6dfd0;
                     border-radius: 6px; margin-bottom: 4px; background: #fdfbf6; cursor: pointer; }
      .student-row:hover { background: #f5f0e8; }
      .student-row input { margin: 0; }
      .student-label { font-size: 12px; }
      #list { max-height: 320px; overflow-y: auto; margin: 8px 0; }
      button.primary { padding: 8px 18px; font-size: 13px; font-weight: bold; border: none; border-radius: 6px;
               background: #2c5f8a; color: white; cursor: pointer; }
      button.primary:disabled { background: #aaa; cursor: default; }
      #msg { margin-top: 10px; font-size: 12px; font-weight: bold; }
      #msg a { color: #2c5f8a; }
      .empty { color: #7a6f5c; }
    </style>
    <div>
      成長レポートを作成したい生徒を1人選んでください。
      <div id="loading">読み込み中...</div>
      <div id="list"></div>
      <div class="empty" id="empty" style="display:none;">まだ回答データがありません。</div>
      <button class="primary" id="runBtn" style="display:none;" onclick="runReport()">このデータでレポートを作成</button>
      <div id="msg"></div>
    </div>
    <script>
      var selected = null;

      google.script.run.withSuccessHandler(function(students) {
        document.getElementById('loading').style.display = 'none';
        if (!students || students.length === 0) {
          document.getElementById('empty').style.display = 'block';
          return;
        }
        const listEl = document.getElementById('list');
        students.forEach(function(s, i) {
          const row = document.createElement('label');
          row.className = 'student-row';
          const radio = document.createElement('input');
          radio.type = 'radio'; radio.name = 'student'; radio.value = s.studentNumber;
          const labelSpan = document.createElement('span');
          labelSpan.className = 'student-label';
          labelSpan.textContent = s.studentNumber + '　' + (s.studentName || '（氏名未入力）') + '　（' + s.count + '回分）';
          row.appendChild(radio);
          row.appendChild(labelSpan);
          row.querySelector('input').addEventListener('change', function() {
            selected = s.studentNumber;
            document.getElementById('runBtn').style.display = 'inline-block';
          });
          listEl.appendChild(row);
        });
      }).getStudentsForDialog();

      function runReport() {
        if (!selected) return;
        const btn = document.getElementById('runBtn');
        const msgEl = document.getElementById('msg');
        btn.disabled = true;
        btn.textContent = '作成中...（PDF生成に数秒かかります）';
        msgEl.textContent = '';
        google.script.run
          .withSuccessHandler(function(result) {
            msgEl.innerHTML = '「' + result.sheetName + '」シートを作成しました。' +
              (result.pdfUrl ? '<br><a href="' + result.pdfUrl + '" target="_blank">PDFを開く</a>' : '（PDFの生成はスキップされました）');
            btn.disabled = false;
            btn.textContent = 'このデータでレポートを作成';
          })
          .withFailureHandler(function(err) {
            msgEl.style.color = '#a8402c';
            msgEl.textContent = 'エラーが発生しました: ' + err.message;
            btn.disabled = false;
            btn.textContent = 'このデータでレポートを作成';
          })
          .generatePersonalGrowthReport(selected);
      }
    </script>
  `).setWidth(440).setHeight(520);

  SpreadsheetApp.getUi().showModalDialog(html, '生徒個人の成長レポートを作成');
}

/**
 * 指定した出席番号の生徒について、時系列の成長レポートシートを生成し、
 * 可能であればPDF化してDriveに保存する
 * @param {string} studentNumber 出席番号
 * @return {{sheetName: string, pdfUrl: (string|null)}}
 */
function generatePersonalGrowthReport(studentNumber) {
  assertTeacher_();
  const target = String(studentNumber || '').trim();
  if (!target) {
    throw new Error('生徒が選択されていません。');
  }

  const studentRecords = loadResponses_().filter(r => r.studentNumber === target);
  if (studentRecords.length === 0) {
    throw new Error('対象の生徒の記録が見つかりませんでした。');
  }

  // 5項目モードは回ごとに項目が違うため、推移は全21項目の記録どうしで比べる
  const selection = selectRecordsForAggregate_(studentRecords);
  const records = selection.records.slice().sort((a, b) => new Date(a.date) - new Date(b.date));

  const studentName = records[records.length - 1].studentName || '';

  const catIds = { Thinking: [], Communication: [], Leadership: [] };
  COMPETENCIES.forEach(c => catIds[c.category].push(c.id));

  const avgOf = (scores, ids) => {
    const vals = ids.map(id => Number(scores[id])).filter(v => v > 0);
    return vals.length ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  };

  const points = records.map(r => {
    const allIds = Object.keys(r.scores).map(Number);
    return {
      date: r.date ? new Date(r.date) : null,
      title: r.title || '',
      mode: r.mode || '',
      overall: avgOf(r.scores, allIds),
      Thinking: avgOf(r.scores, catIds.Thinking),
      Communication: avgOf(r.scores, catIds.Communication),
      Leadership: avgOf(r.scores, catIds.Leadership)
    };
  });

  const fmtDate = (d) => d ? Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy年M月d日') : '（日付不明）';

  // ── 文章（成長の要約）の自動生成 ──
  const first = points[0];
  const latest = points[points.length - 1];
  const summarySentences = [];
  summarySentences.push(`${target}　${studentName} さんの、これまで${points.length}回分の記録をもとにした成長レポートです。`);

  if (points.length === 1) {
    summarySentences.push('現時点では記録が1回分のみのため、成長の推移はまだ確認できません。次回以降の記録が蓄積されると、変化がグラフに表示されるようになります。');
  } else if (first.overall !== null && latest.overall !== null) {
    const diff = latest.overall - first.overall;
    if (Math.abs(diff) < 0.05) {
      summarySentences.push(`初回（${fmtDate(first.date)}、平均${first.overall.toFixed(1)}）から直近（${fmtDate(latest.date)}、平均${latest.overall.toFixed(1)}）にかけて、総合的な自己評価レベルはほぼ横ばいで推移しています。`);
    } else if (diff > 0) {
      summarySentences.push(`初回（${fmtDate(first.date)}、平均${first.overall.toFixed(1)}）から直近（${fmtDate(latest.date)}、平均${latest.overall.toFixed(1)}）にかけて、総合的な自己評価レベルが${diff.toFixed(1)}ポイント向上しており、着実な成長が見られます。`);
    } else {
      summarySentences.push(`初回（${fmtDate(first.date)}、平均${first.overall.toFixed(1)}）から直近（${fmtDate(latest.date)}、平均${latest.overall.toFixed(1)}）にかけて、総合的な自己評価レベルは${Math.abs(diff).toFixed(1)}ポイント低下しています。活動内容や心境の変化など、背景を本人と対話で確認してみるとよいかもしれません。`);
    }

    const catDiffs = ['Thinking', 'Communication', 'Leadership']
      .map(cat => (first[cat] !== null && latest[cat] !== null) ? { cat, diff: latest[cat] - first[cat] } : null)
      .filter(d => d !== null);
    if (catDiffs.length > 0) {
      const best = catDiffs.slice().sort((a, b) => b.diff - a.diff)[0];
      if (best.diff > 0.05) {
        summarySentences.push(`カテゴリー別では特に「${CATEGORY_LABEL[best.cat]}」の伸びが大きく、初回から${best.diff.toFixed(1)}ポイント向上しています。`);
      }
    }
  }
  if (selection.note) summarySentences.push(selection.note);
  summarySentences.push('数値の変化だけでなく、本人がどんな経験を通じてそう感じるようになったのか、対話のきっかけとしてこのレポートをご活用ください。');

  // ── シート作成 ──
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = GROWTH_WORK_SHEET_NAME;
  const pdfFileName = sanitizeSheetName_(`${target}_${studentName}`, GROWTH_SHEET_PREFIX);
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  sheet.getRange('A1').setValue(`${studentName || target} さんの成長記録`)
    .setFontSize(18).setFontWeight('bold').setFontColor('#2a2420');
  sheet.getRange('A2').setValue(`出席番号: ${target}　／　記録数: ${points.length}回　／　作成日: ${fmtDate(new Date())}`)
    .setFontSize(10).setFontColor('#7a6f5c');

  sheet.getRange('A4').setValue('成長の要約').setFontSize(13).setFontWeight('bold').setFontColor('#a8402c');
  sheet.getRange('A5').setValue(summarySentences.join('\n\n')).setFontSize(12).setWrap(true);
  sheet.setRowHeight(5, 180);
  sheet.getRange('A5:J5').merge();

  // ── 提出履歴の一覧表 ──
  const tableStartRow = 7;
  sheet.getRange(tableStartRow, 1).setValue('提出履歴').setFontSize(13).setFontWeight('bold').setFontColor('#a8402c');
  const tableHeader = ['実施日', 'タイトル', 'モード', '総合平均', 'Thinking平均', 'Communication平均', 'Leadership平均'];
  const tableRows = [tableHeader].concat(points.map(p => [
    p.date || '', p.title, p.mode,
    p.overall !== null ? Number(p.overall.toFixed(2)) : '',
    p.Thinking !== null ? Number(p.Thinking.toFixed(2)) : '',
    p.Communication !== null ? Number(p.Communication.toFixed(2)) : '',
    p.Leadership !== null ? Number(p.Leadership.toFixed(2)) : ''
  ]));
  const tableHeaderRow = tableStartRow + 1;
  sheet.getRange(tableHeaderRow, 1, tableRows.length, tableHeader.length).setValues(tableRows);
  sheet.getRange(tableHeaderRow, 1, 1, tableHeader.length).setFontWeight('bold').setBackground('#f0ece3');
  sheet.getRange(tableHeaderRow + 1, 1, points.length, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(tableHeaderRow + 1, 4, points.length, 4).setNumberFormat('0.00');

  // ── 推移グラフ（折れ線）用データ：作業エリアに書き出してから列を隠す ──
  const WORK_COL = 14;
  const chartDataHeaderRow = tableHeaderRow;
  const chartHeader = ['実施日', '総合平均', CATEGORY_LABEL.Thinking, CATEGORY_LABEL.Communication, CATEGORY_LABEL.Leadership];
  const chartRows = [chartHeader].concat(points.map(p => [
    p.date ? Utilities.formatDate(p.date, Session.getScriptTimeZone() || 'Asia/Tokyo', 'M/d') : '',
    p.overall !== null ? Number(p.overall.toFixed(2)) : null,
    p.Thinking !== null ? Number(p.Thinking.toFixed(2)) : null,
    p.Communication !== null ? Number(p.Communication.toFixed(2)) : null,
    p.Leadership !== null ? Number(p.Leadership.toFixed(2)) : null
  ]));
  // 「9/19」のような日付ラベルを、シートに日付へ自動変換されないよう、文字列として書き込む
  sheet.getRange(chartDataHeaderRow + 1, WORK_COL, points.length, 1).setNumberFormat('@');
  sheet.getRange(chartDataHeaderRow, WORK_COL, chartRows.length, chartHeader.length).setValues(chartRows);

  if (points.length > 1) {
    const growthChart = sheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(sheet.getRange(chartDataHeaderRow, WORK_COL, chartRows.length, chartHeader.length))
      .setNumHeaders(1)
      .setOption('useFirstColumnAsDomain', true)
      .setPosition(tableStartRow + points.length + 4, 1, 0, 0)
      .setOption('title', '総合・カテゴリー別 平均レベルの推移')
      .setOption('titleTextStyle', { fontSize: 14, bold: true })
      .setOption('legend', { position: 'top' })
      .setOption('vAxis', { minValue: 0, maxValue: 4, gridlines: { count: 5 }, title: '平均レベル（4段階）' })
      .setOption('hAxis', { title: '実施日' })
      .setOption('colors', ['#a8402c', CATEGORY_COLOR.Thinking, CATEGORY_COLOR.Communication, CATEGORY_COLOR.Leadership])
      .setOption('pointSize', 6)
      .setOption('width', 780).setOption('height', 420)
      .build();
    sheet.insertChart(growthChart);
  }

  sheet.hideColumns(WORK_COL, chartHeader.length);
  autoFitTableColumns_(sheet, 1, tableHeader.length, 220);
  sheet.setFrozenRows(0);
  sheet.activate();

  // ── PDF化（失敗してもシート自体は完成しているため、エラーは握りつぶして通知のみ）──
  let pdfUrl = null;
  try {
    SpreadsheetApp.flush(); // 作り直したグラフを確実に反映してからPDF化する
    pdfUrl = exportSheetAsPdf_(sheet, pdfFileName);
  } catch (e) {
    console.error('PDF生成に失敗しました:', e);
  }

  return { sheetName: sheetName, pdfUrl: pdfUrl };
}

/**
 * 指定したシート1枚だけをPDF化し、スプレッドシートと同じフォルダ内に保存して、そのURLを返す
 */
function exportSheetAsPdf_(sheet, fileName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export'
    + '?format=pdf&size=A4&portrait=true&fitw=true&gridlines=false'
    + '&printtitle=false&sheetnames=false&pagenumbers=false'
    + '&horizontal_alignment=CENTER&vertical_alignment=TOP'
    + '&gid=' + sheet.getSheetId();

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('PDFの生成に失敗しました（HTTP ' + response.getResponseCode() + '）。');
  }

  const blob = response.getBlob().setName(fileName + '.pdf');
  const folder = getOrCreateGrowthPdfFolder_();
  // 同じ生徒のPDFが既にあればゴミ箱に移してから作り直す（何度作成してもファイルが増えないように）
  const oldFiles = folder.getFilesByName(fileName + '.pdf');
  while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);
  const file = folder.createFile(blob);
  // ※プライバシー配慮のため共有設定は変更しない（デフォルトのままスプレッドシートの所有者のみ閲覧可）
  return file.getUrl();
}

/**
 * 成長レポートPDFの保存先フォルダを取得（無ければ、スプレッドシートと同じ場所に作成）する
 */
function getOrCreateGrowthPdfFolder_() {
  const FOLDER_NAME = '成長レポートPDF';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parents = DriveApp.getFileById(ss.getId()).getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const existing = parentFolder.getFoldersByName(FOLDER_NAME);
  if (existing.hasNext()) return existing.next();
  return parentFolder.createFolder(FOLDER_NAME);
}

