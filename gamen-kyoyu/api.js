// 生徒用ページ(index.html)と先生用ページ(teacher.html)で共有する処理

// Apps Script から読み出す。学校アカウントのログイン状態でも動くよう <script> タグ（JSONP）で読み込む
function callApi(params, timeoutMs) {
  return new Promise((resolve, reject) => {
    if (typeof API_URL === 'undefined' || !API_URL) { reject(new Error('Apps Script とまだつながっていません（config.js の API_URL）')); return; }
    const cb = '__gamen' + Date.now() + Math.floor(Math.random() * 100000);
    const s = document.createElement('script');
    const timer = setTimeout(() => { done(); reject(new Error('応答がありません')); }, timeoutMs || 30000);
    function done() { clearTimeout(timer); delete window[cb]; s.remove(); }
    window[cb] = data => { done(); data && data.ok ? resolve(data) : reject(new Error((data && data.error) || '読み込みに失敗しました')); };
    s.onerror = () => { done(); reject(new Error('Apps Script に接続できませんでした')); };
    s.src = API_URL + '?' + new URLSearchParams(Object.assign({}, params, { callback: cb, t: Date.now() }));
    document.head.appendChild(s);
  });
}

// Apps Script に書き込む（先生用）。返事は読めない（no-cors）ので、届いたかは callApi の meta で確かめる
function postApi(body) {
  return fetch(API_URL, {
    method: 'POST',
    mode: 'no-cors',
    credentials: 'include',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
}
