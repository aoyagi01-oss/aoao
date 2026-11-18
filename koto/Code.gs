/**
 * デジタルお箏
 * スプレッドシート「デジタルお箏 曲と調弦」に貼り付けて、ウェブアプリとして公開します。
 * 生徒の画面は Index.html。曲と調弦はスプレッドシートから読み込みます。
 */

const SHEET_SONGS = '曲';
const SHEET_TUNINGS = '調弦';
const STRING_NAMES = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '斗', '為', '巾'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('デジタルお箏')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** 生徒の画面から呼ばれる。曲と調弦をまとめて返す。 */
function getData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return {
    tunings: readTunings_(ss.getSheetByName(SHEET_TUNINGS)),
    songs: readSongs_(ss.getSheetByName(SHEET_SONGS)),
  };
}

function readTunings_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0].map(String);
  const nameCol = header.indexOf('調弦名');
  const cols = STRING_NAMES.map(n => header.indexOf(n));
  const list = [];
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][nameCol] || '').trim();
    if (!name) continue;
    const notes = cols.map(c => (c >= 0 ? String(values[r][c] || '').trim() : ''));
    if (notes.some(n => !n)) continue; // 13本そろっていない行は使わない
    list.push({ name: name, notes: notes });
  }
  return list;
}

function readSongs_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const header = values[0].map(String);
  const col = name => header.findIndex(h => h.indexOf(name) === 0);
  const c = {
    show: col('表示'), title: col('曲名'), tuning: col('調弦'),
    tempo: col('速さ'), score: col('楽譜'), lyrics: col('歌詞'), memo: col('メモ'),
  };
  const get = (row, i) => (i >= 0 ? row[i] : '');
  const list = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const title = String(get(row, c.title) || '').trim();
    const score = String(get(row, c.score) || '').trim();
    if (!title || !score) continue;
    if (c.show >= 0 && get(row, c.show) === false) continue;
    list.push({
      title: title,
      tuning: String(get(row, c.tuning) || '').trim(),
      tempo: Number(get(row, c.tempo)) || 80,
      score: score,
      lyrics: String(get(row, c.lyrics) || ''),
      memo: String(get(row, c.memo) || ''),
    });
  }
  return list;
}
