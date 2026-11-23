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
function prepare_(urlA, urlB) {
  const ids = { A: slidesId_(urlA, 'A'), B: slidesId_(urlB, 'B') };
  const ver = Date.now().toString(36);
  const setPrep = function (o) { writeJson_('prep', Object.assign({ at: Date.now() }, o)); };
  try {
    const decks = {};
    let total = 0;
    ROLES.forEach(function (r) {
      let pres;
      try { pres = SlidesApp.openById(ids[r]); }
      catch (err) { throw new Error(r + '役のスライドを開けません。URL と、自分のアカウントで開けるかを確かめてください'); }
      const pages = pres.getSlides().slice(0, MAX_PAGES).map(function (s) { return s.getObjectId(); });
      if (!pages.length) throw new Error(r + '役のスライドにページがありません');
      decks[r] = { name: pres.getName(), pages: pages };
      total += pages.length;
    });

    const folder = folder_();
    const files = { A: [], B: [] };
    let done = 0;
    setPrep({ doing: true, done: 0, total: total });
    ROLES.forEach(function (r) {
      decks[r].pages.forEach(function (pageId, i) {
        const blob = pageImage_(ids[r], pageId).setName(ver + '-' + r + '-' + (i + 1));
        files[r].push(folder.createFile(blob).getId());
        done++;
        setPrep({ doing: true, done: done, total: total });
      });
    });

    // 前に準備した画像は、新しい画像ができてから消す
    const old = readJson_('deck');
    const counts = { A: files.A.length, B: files.B.length };
    writeJson_('files', files);
    writeJson_('deck', {
      ver: ver,
      n: Math.max(counts.A, counts.B),
      counts: counts,
      names: { A: decks.A.name, B: decks.B.name }
    });
    writeJson_('state', { step: 0, hidden: false, swap: false, at: Date.now() });
    setPrep({ doing: false, done: total, total: total, ver: ver });
    if (old && old.ver) trashOld_(old.ver);
  } catch (err) {
    setPrep({ doing: false, error: String(err && err.message || err) });
    try { trashOld_(ver); } catch (e2) {}   // 途中まで作った画像は消す
    throw err;
  }
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
function pageImage_(presId, pageId) {
  const url = 'https://docs.google.com/presentation/d/' + presId + '/export/png?pageid=' + encodeURIComponent(pageId);
  let res;
  for (let tries = 0; tries < 3; tries++) {
    res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() === 200) break;
    Utilities.sleep(1500 * (tries + 1));
  }
  if (res.getResponseCode() !== 200) throw new Error('スライドを画像にできませんでした（' + res.getResponseCode() + '）');
  const png = res.getBlob().setContentType('image/png');
  let jpg = null;
  try { jpg = png.getAs('image/jpeg'); } catch (err) {}
  return jpg && jpg.getBytes().length < png.getBytes().length ? jpg : png;
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
