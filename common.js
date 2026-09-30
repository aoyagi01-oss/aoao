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
