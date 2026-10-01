// 生徒用ページ(index.html)と先生用ページ(teacher.html)で共有する処理

// 教材データを URL の # 部分に入れるためのエンコード（日本語対応の base64url）
function encodeLesson(lesson) {
  const bytes = new TextEncoder().encode(JSON.stringify(lesson));
  let bin = '';
  bytes.forEach(b => { bin += String.fromCharCode(b); });
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

function decodeLesson(str) {
  let s = str.replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  const bin = atob(s);
  return JSON.parse(new TextDecoder().decode(Uint8Array.from(bin, c => c.charCodeAt(0))));
}

function lessonFromHash() {
  const m = location.hash.match(/[#&]d=([^&]+)/);
  if (!m) return null;
  try { return decodeLesson(m[1]); } catch (e) { return null; }
}

// 生徒がボタンを押したときに開く ChatGPT の URL
// share（共有チャットのリンク）があればそれを優先し、なければプロンプトを ?q= で渡して新しいチャットを始める
function chatUrl(lesson) {
  if (lesson.share) return lesson.share;
  return 'https://chatgpt.com/?q=' + encodeURIComponent(lesson.prompt);
}

// 英語の読み上げ（Chromebook の音声合成を使用）
function speak(text) {
  if (!('speechSynthesis' in window)) return;
  speechSynthesis.cancel();
  const u = new SpeechSynthesisUtterance(text);
  u.lang = 'en-US';
  u.rate = 0.85;
  speechSynthesis.speak(u);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

// スプレッドシート（Apps Script）を呼び出す。学校アカウントのログイン状態でも動くよう <script> タグで読み込む
function callApi(params) {
  return new Promise((resolve, reject) => {
    if (typeof API_URL === 'undefined' || !API_URL) { reject(new Error('スプレッドシートとまだつながっていません')); return; }
    const cb = '__aiEikaiwa' + Date.now() + Math.floor(Math.random() * 1000);
    const s = document.createElement('script');
    // Apps Script は最初の呼び出しが遅いことがあるので、長めに待つ
    const timer = setTimeout(() => { done(); reject(new Error('スプレッドシートから応答がありません。少し待ってから、もう一度押してください')); }, 40000);
    function done() { clearTimeout(timer); delete window[cb]; s.remove(); }
    window[cb] = data => { done(); data && data.ok ? resolve(data) : reject(new Error((data && data.error) || '読み込みに失敗しました')); };
    s.onerror = () => { done(); reject(new Error('スプレッドシートに接続できませんでした')); };
    s.src = API_URL + '?' + new URLSearchParams(Object.assign({}, params, { callback: cb, t: Date.now() }));
    document.head.appendChild(s);
  });
}

// スプレッドシートの1行を、生徒ページで使う教材データにする
function lessonFromRow(row) {
  return {
    title: row.title,
    scene: row.scene,
    phrases: parsePhrases(row.phrases),
    prompt: row.prompt || buildPrompt(row)
  };
}
