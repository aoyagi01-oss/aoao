// 画面配信（静止画）：先生の画面の画像をあずかり、生徒のページに渡すプログラム
// 使い方は gamen-kyoyu/README.md の「最初に1回だけ」を見てください。
// スプレッドシートは使いません。画像は Google のキャッシュに最大6時間だけ置かれ、自動で消えます。

// ▼ 先生用ページで「配信開始」するときの合言葉。必ず自分だけの言葉に変えてください。
const PASSWORD = 'ここを合言葉に変える';

const CHUNK = 90000;          // キャッシュ1件あたりの文字数（上限 100KB より少し小さく）
const KEEP = 6 * 60 * 60;     // キャッシュに置く秒数（最大6時間）
const STALE = 90 * 1000;      // 先生のページから合図がこの時間ないと「止まっている」とみなす

// 生徒用・先生用ページからの読み出し（<script> タグで呼ぶ JSONP）
function doGet(e) {
  const p = e.parameter || {};
  let out;
  try {
    if (p.action === 'check') {
      checkPassword_(p.key);
      out = { ok: true };
    } else if (p.action === 'meta') {
      out = Object.assign({ ok: true }, status_());
    } else {
      out = get_(p.have);
    }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  const json = JSON.stringify(out);
  if (p.callback && /^[\w$.]+$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// 先生用ページからの書き込み（画像・合図・終了）
function doPost(e) {
  let out;
  try {
    const p = JSON.parse(e.postData.contents);
    checkPassword_(p.key);
    if (p.action === 'put') put_(String(p.ver), String(p.img), !!p.paused);
    else if (p.action === 'beat') beat_(!!p.paused);
    else if (p.action === 'stop') stop_();
    out = { ok: true };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function checkPassword_(key) {
  if (PASSWORD === 'ここを合言葉に変える') throw new Error('Apps Script の PASSWORD（合言葉）がまだ設定されていません');
  const norm = function (v) { return String(v || '').normalize('NFKC').trim(); };
  if (norm(key) !== norm(PASSWORD)) throw new Error('合言葉がちがいます');
}

function cache_() {
  return CacheService.getScriptCache();
}

function meta_() {
  const s = cache_().get('meta');
  return s ? JSON.parse(s) : null;
}

function setMeta_(m) {
  cache_().put('meta', JSON.stringify(m), KEEP);
}

function put_(ver, img, paused) {
  if (!/^[\w-]{1,40}$/.test(ver)) throw new Error('ver が不正です');
  const parts = {};
  let n = 0;
  for (let i = 0; i < img.length; i += CHUNK) parts['img:' + ver + ':' + (n++)] = img.slice(i, i + CHUNK);
  cache_().putAll(parts, KEEP);
  // 画像をすべて置いてから meta を切り替える（読み手が途中の画像を見ないように）
  setMeta_({ ver: ver, n: n, live: true, paused: paused, at: Date.now() });
}

function beat_(paused) {
  const m = meta_();
  if (!m) return;
  m.live = true;
  m.paused = paused;
  m.at = Date.now();
  setMeta_(m);
}

function stop_() {
  const m = meta_();
  if (!m) return;
  m.live = false;
  m.paused = false;
  setMeta_(m);
}

function status_() {
  const m = meta_();
  if (!m) return { ver: '', live: false, paused: false };
  return { ver: m.ver, live: !!m.live && Date.now() - m.at < STALE, paused: !!m.paused, at: m.at };
}

// have：生徒がすでに持っている画像の番号。同じなら画像は送らない
function get_(have) {
  const m = meta_();
  const st = status_();
  if (!m || (have && have === m.ver)) return Object.assign({ ok: true, same: !!m }, st);
  const keys = [];
  for (let i = 0; i < m.n; i++) keys.push('img:' + m.ver + ':' + i);
  const got = cache_().getAll(keys);
  const img = keys.map(function (k) { return got[k]; });
  if (img.some(function (v) { return v == null; })) return Object.assign({ ok: true, same: true }, st);
  return Object.assign({ ok: true, img: img.join('') }, st);
}
