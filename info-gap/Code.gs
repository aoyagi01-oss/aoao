// インフォメーションギャップ配信：A役・B役に別々のスライドを配り、先生の操作で全員いっしょにページを進めるプログラム
// 使い方は info-gap/README.md の「最初に1回だけ」を見てください。
// 画面共有（gamen-kyoyu）とは別の Apps Script として作ってください。

// ▼ 先生用ページで使う合言葉。必ず自分だけの言葉に変えてください。
const PASSWORD = 'ここを合言葉に変える';

const FOLDER_NAME = 'インフォギャップ配信の画像（自動で作られます）';
const CHUNK = 90000;             // キャッシュ1件あたりの文字数（上限 100KB より少し小さく）
const KEEP = 6 * 60 * 60;        // キャッシュに置く秒数（最大6時間。消えてもドライブから読み直す）
const MAX_PAGES = 40;            // 1つのスライドから画像にする最大ページ数
const ROLES = ['A', 'B'];

// ── 生徒用・先生用ページからの読み出し（<script> タグで呼ぶ JSONP）──
function doGet(e) {
  const p = e.parameter || {};
  let out;
  try {
    if (p.action === 'check') {
      checkPassword_(p.key);
      out = { ok: true };
    } else if (p.action === 'img') {
      out = Object.assign({ ok: true }, image_(String(p.ver || ''), String(p.role || ''), Number(p.i)));
    } else {
      out = Object.assign({ ok: true }, state_());
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

// ── 先生用ページからの書き込み（スライドの準備・ページ送り）──
function doPost(e) {
  let out;
  try {
    const p = JSON.parse(e.postData.contents);
    checkPassword_(p.key);
    if (p.action === 'prepare') prepare_(String(p.a || ''), String(p.b || ''));
    else if (p.action === 'go') go_(p);
    out = { ok: true };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// 最初に1回、エディタでこの関数を選んで「▶ 実行」し、スライド・ドライブ・外部への接続の許可を出しておく
function setup() {
  folder_();
  Logger.log('準備できました。次は「デプロイ → 新しいデプロイ」です。');
}

function checkPassword_(key) {
  if (PASSWORD === 'ここを合言葉に変える') throw new Error('Apps Script の PASSWORD（合言葉）がまだ設定されていません');
  const norm = function (v) { return String(v || '').normalize('NFKC').trim(); };
  if (norm(key) !== norm(PASSWORD)) throw new Error('合言葉がちがいます');
}

const cache_ = function () { return CacheService.getScriptCache(); };
const props_ = function () { return PropertiesService.getScriptProperties(); };

// ── いまの状態（何枚目か・かくしているか・入れかえ中か）──
// 生徒が何度も読むので、ふだんはキャッシュから読み、なければ保存しておいた値を使う
function readJson_(name) {
  let s = cache_().get(name);
  if (!s) {
    s = props_().getProperty(name);
    if (s) cache_().put(name, s, KEEP);
  }
  return s ? JSON.parse(s) : null;
}

function writeJson_(name, obj) {
  const s = JSON.stringify(obj);
  props_().setProperty(name, s);
  cache_().put(name, s, KEEP);
}

function state_() {
  const deck = readJson_('deck') || {};
  const st = readJson_('state') || {};
  const prep = readJson_('prep') || {};
  return {
    ver: deck.ver || '',
    n: deck.n || 0,
    counts: deck.counts || {},
    names: deck.names || {},
    step: st.step || 0,
    hidden: !!st.hidden,
    swap: !!st.swap,
    prep: prep
  };
}

function go_(p) {
  const deck = readJson_('deck') || {};
  const n = deck.n || 0;
  const step = Math.max(0, Math.min(n, Math.floor(Number(p.step) || 0)));
  writeJson_('state', { step: step, hidden: !!p.hidden, swap: !!p.swap, at: Date.now() });
}

// ── Google スライドを1ページずつ画像にして、ドライブのフォルダに保存する ──
// 時間がかかったり、Google に「少し待って」と言われたりしたときは途中で止め、もう一度押すと続きから画像にする
const TIME_LIMIT = 4.5 * 60 * 1000;   // Apps Script は1回6分までしか動けないので、その前に止める
const RESUME_MS = 30 * 60 * 1000;     // この時間内にもう一度押すと、続きから
const GAP_MS = 1200;                  // 1ページごとにあける間（まとめて頼むと断られるため）

function prepare_(urlA, urlB) {
  const started = Date.now();
  const ids = { A: slidesId_(urlA, 'A'), B: slidesId_(urlB, 'B') };
  const setPrep = function (o) { writeJson_('prep', Object.assign({ at: Date.now() }, o)); };
  let work = readJson_('work');
  if (!work || work.ids.A !== ids.A || work.ids.B !== ids.B || Date.now() - work.at > RESUME_MS) {
    if (work && work.ver) { try { trashOld_(work.ver); } catch (e) {} }   // 前の途中の画像は消す
    work = { ver: Date.now().toString(36), ids: ids, decks: openDecks_(ids), files: { A: [], B: [] } };
  }
  const total = work.decks.A.pages.length + work.decks.B.pages.length;
  const doneCount = function () { return work.files.A.length + work.files.B.length; };
  const save = function () { work.at = Date.now(); writeJson_('work', work); };
  save();
  setPrep({ doing: true, done: doneCount(), total: total });

  const folder = folder_();
  try {
    for (let ri = 0; ri < ROLES.length; ri++) {
      const r = ROLES[ri];
      const pages = work.decks[r].pages;
      for (let i = work.files[r].length; i < pages.length; i++) {
        if (Date.now() - started > TIME_LIMIT) throw new Error('時間がかかっています');
        if (doneCount() > 0) Utilities.sleep(GAP_MS);
        const blob = pageImage_(work.ids[r], pages[i], started).setName(work.ver + '-' + r + '-' + (i + 1));
        work.files[r].push(folder.createFile(blob).getId());
        save();
        setPrep({ doing: true, done: doneCount(), total: total });
      }
    }
  } catch (err) {
    const msg = String(err && err.message || err);
    setPrep({ doing: false, done: doneCount(), total: total, error: msg + '。' + doneCount() + ' / ' + total + ' ページまでできています。1分ほど待ってから、もう一度「スライドを準備する」を押すと続きから画像にします' });
    throw err;
  }

  // 前に準備した画像は、新しい画像ができてから消す
  const old = readJson_('deck');
  const counts = { A: work.files.A.length, B: work.files.B.length };
  writeJson_('files', work.files);
  writeJson_('deck', {
    ver: work.ver,
    n: Math.max(counts.A, counts.B),
    counts: counts,
    names: { A: work.decks.A.name, B: work.decks.B.name }
  });
  writeJson_('state', { step: 0, hidden: false, swap: false, at: Date.now() });
  props_().deleteProperty('work');
  cache_().remove('work');
  setPrep({ doing: false, done: total, total: total, ver: work.ver });
  if (old && old.ver && old.ver !== work.ver) trashOld_(old.ver);
}

function openDecks_(ids) {
  const decks = {};
  ROLES.forEach(function (r) {
    let pres;
    try { pres = SlidesApp.openById(ids[r]); }
    catch (err) { throw new Error(r + '役のスライドを開けません。URL と、自分のアカウントで開けるかを確かめてください'); }
    const pages = pres.getSlides().slice(0, MAX_PAGES).map(function (s) { return s.getObjectId(); });
    if (!pages.length) throw new Error(r + '役のスライドにページがありません');
    decks[r] = { name: pres.getName(), pages: pages };
  });
  return decks;
}

function slidesId_(url, role) {
  const m = String(url).match(/\/presentation\/d\/([\w-]{20,})/) || String(url).match(/^([\w-]{25,})$/);
  if (!m) throw new Error(role + '役の URL が Google スライドの URL ではありません');
  return m[1];
}

function folder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

// スライド1ページを画像にする。PNG と JPEG の小さいほうを使う
// 「少し待って」（429）と言われたら、間をあけてやり直す。だめならスライドの縮小画像を使う
function pageImage_(presId, pageId, started) {
  const auth = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const url = 'https://docs.google.com/presentation/d/' + presId + '/export/png?pageid=' + encodeURIComponent(pageId);
  const waits = [3000, 8000, 15000, 25000];
  let code = 0, png = null;
  for (let tries = 0; ; tries++) {
    const res = UrlFetchApp.fetch(url, { headers: auth, muteHttpExceptions: true, followRedirects: true });
    code = res.getResponseCode();
    if (code === 200) { png = res.getBlob(); break; }
    if (tries >= waits.length || (code !== 429 && code < 500)) break;
    if (Date.now() - started + waits[tries] > TIME_LIMIT) break;
    Utilities.sleep(waits[tries]);
  }
  if (!png) png = thumbnail_(presId, pageId, auth);
  if (!png) {
    throw new Error(code === 429
      ? 'Google から「少し待ってください」と言われました（429）'
      : 'スライドを画像にできませんでした（' + code + '）');
  }
  png.setContentType('image/png');
  let jpg = null;
  try { jpg = png.getAs('image/jpeg'); } catch (err) {}
  return jpg && jpg.getBytes().length < png.getBytes().length ? jpg : png;
}

// 予備：Google スライドの機能で、ページの縮小画像（横1600）を作ってもらう
function thumbnail_(presId, pageId, auth) {
  try {
    const api = 'https://slides.googleapis.com/v1/presentations/' + presId + '/pages/' + encodeURIComponent(pageId) +
      '/thumbnail?thumbnailProperties.thumbnailSize=LARGE&thumbnailProperties.mimeType=PNG';
    const res = UrlFetchApp.fetch(api, { headers: auth, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    const img = UrlFetchApp.fetch(JSON.parse(res.getContentText()).contentUrl, { muteHttpExceptions: true });
    return img.getResponseCode() === 200 ? img.getBlob() : null;
  } catch (err) {
    return null;
  }
}

function trashOld_(ver) {
  const files = folder_().getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf(ver + '-') === 0) f.setTrashed(true);
  }
}

// ── 生徒用ページに1ページ分の画像を渡す（キャッシュになければドライブから読む）──
function image_(ver, role, i) {
  const deck = readJson_('deck');
  if (!deck || deck.ver !== ver) throw new Error('スライドが新しくなりました');
  if (ROLES.indexOf(role) < 0 || !(i >= 0)) throw new Error('ページの指定が不正です');
  const head = 'img:' + ver + ':' + role + ':' + i;
  const c = cache_();
  const meta = c.get(head);
  if (meta) {
    const m = JSON.parse(meta);
    const keys = [];
    for (let k = 0; k < m.n; k++) keys.push(head + ':' + k);
    const got = c.getAll(keys);
    const parts = keys.map(function (k) { return got[k]; });
    if (parts.every(function (v) { return v != null; })) return { mime: m.mime, img: parts.join('') };
  }
  const files = readJson_('files') || {};
  const id = (files[role] || [])[i];
  if (!id) return { none: true };
  const blob = DriveApp.getFileById(id).getBlob();
  const b64 = Utilities.base64Encode(blob.getBytes());
  const mime = blob.getContentType();
  const put = {};
  let n = 0;
  for (let k = 0; k < b64.length; k += CHUNK) put[head + ':' + (n++)] = b64.slice(k, k + CHUNK);
  try {
    c.putAll(put, KEEP);
    c.put(head, JSON.stringify({ n: n, mime: mime }), KEEP);
  } catch (err) {}
  return { mime: mime, img: b64 };
}
