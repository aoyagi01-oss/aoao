// 画面の一部を光らせながら順番に説明する「使い方ガイド」

const TEACHER_TOUR = [
  {
    title: 'ようこそ！',
    text: 'このページで、授業ごとの英会話の教材を作ります。<br>やることは <b>①選ぶ → ②直す → ③公開する</b> の3つだけです。<br>生徒は、いつも同じページを開くだけで今日の教材が表示されます。'
  },
  {
    target: '#startBox',
    title: '① どこから作る？',
    text: '<b>ひな形</b>のボタンを押すと、下の内容がまとめて入ります。<br>前に作った教材は<b>「これまでの教材」</b>の「使う」で呼び出せます。'
  },
  {
    target: '#basicsBox',
    title: '公開日とタイトル',
    text: '<b>公開日</b>の日から、生徒の画面にこの教材が出ます。<br>明日の日付にしておけば、前日のうちに準備（予約）できます。'
  },
  {
    target: '#sceneBox',
    title: '場面',
    text: 'どんな場面で会話するかを書きます。日本語でOK。<br>生徒の画面にもそのまま表示されます。'
  },
  {
    target: '#rolesBox',
    title: 'AIの役と生徒の役',
    text: 'AIが演じる役と、生徒の役です。<br>入力欄をクリックすると<b>候補</b>が出るので、選ぶだけでもOKです。'
  },
  {
    target: '#levelBox',
    title: 'レベルと回数',
    text: '生徒のレベルに合わせて、AIが<b>使う単語や文の長さ</b>を調整します。<br>回数の目安が来ると、AIが会話を終えて<b>日本語でフィードバック</b>します。'
  },
  {
    target: '#phraseBox',
    title: '今日の目標表現',
    text: '左に英語、右に日本語を入れます。<br>🔊で発音を確認、✕で削除、<b>＋表現を追加</b>で行を増やせます。<br>AIは会話の中で、生徒がこの表現を使えるように話を進めます。'
  },
  {
    target: '#publishBox',
    title: '③ 生徒に公開する',
    text: '<b>合言葉</b>を入れて「📢 生徒に公開する」を押せば完了です。<br>生徒のページが、この教材に切り替わります。'
  },
  {
    target: '#studentBox',
    title: '生徒に配るURLは1つだけ',
    text: '生徒用ページはいつも同じURLです。<br>Classroom に<b>最初に1回</b>貼っておけば、あとは公開するだけで切り替わります。'
  },
  {
    title: 'さっそく試してみましょう',
    text: '「レストラン」などのひな形を選んで、「📢 生徒に公開する」を押してみてください。<br>表示される<b>「生徒の画面を確認」</b>から、生徒が見る画面も確かめられます。<br><br>このガイドは右上の <b>❓ 使い方ガイド</b> でいつでも見られます。'
  }
];

let tourSteps = [];
let tourIndex = 0;
let tourEls = null;

function startTour(steps) {
  tourSteps = steps;
  tourIndex = 0;
  if (!tourEls) buildTour();
  tourEls.layer.classList.remove('hidden');
  showTourStep();
}

function endTour() {
  clearTourFocus();
  tourEls.layer.classList.add('hidden');
}

function buildTour() {
  const layer = document.createElement('div');
  layer.className = 'tour-layer hidden';
  layer.innerHTML =
    '<div class="tour-dim"></div>' +
    '<div class="tour-bubble" role="dialog" aria-live="polite">' +
    '  <div class="tour-count"></div>' +
    '  <h3 class="tour-title"></h3>' +
    '  <div class="tour-text"></div>' +
    '  <div class="tour-nav">' +
    '    <button class="btn small tour-skip" type="button">とじる</button>' +
    '    <span style="flex:1"></span>' +
    '    <button class="btn small tour-prev" type="button">← もどる</button>' +
    '    <button class="btn small primary tour-next" type="button">つぎへ →</button>' +
    '  </div>' +
    '</div>';
  document.body.appendChild(layer);
  tourEls = {
    layer,
    dim: layer.querySelector('.tour-dim'),
    bubble: layer.querySelector('.tour-bubble'),
    count: layer.querySelector('.tour-count'),
    title: layer.querySelector('.tour-title'),
    text: layer.querySelector('.tour-text'),
    prev: layer.querySelector('.tour-prev'),
    next: layer.querySelector('.tour-next')
  };
  layer.querySelector('.tour-skip').addEventListener('click', endTour);
  tourEls.prev.addEventListener('click', () => { tourIndex--; showTourStep(); });
  tourEls.next.addEventListener('click', () => {
    if (tourIndex >= tourSteps.length - 1) { endTour(); return; }
    tourIndex++;
    showTourStep();
  });
  document.addEventListener('keydown', e => {
    if (layer.classList.contains('hidden')) return;
    if (e.key === 'Escape') endTour();
    if (e.key === 'ArrowRight') tourEls.next.click();
    if (e.key === 'ArrowLeft' && tourIndex > 0) tourEls.prev.click();
  });
  window.addEventListener('resize', placeBubble);
  window.addEventListener('scroll', placeBubble, { passive: true });
}

function clearTourFocus() {
  document.querySelectorAll('.tour-focus').forEach(el => el.classList.remove('tour-focus'));
}

function showTourStep() {
  const step = tourSteps[tourIndex];
  clearTourFocus();
  tourEls.count.textContent = (tourIndex + 1) + ' / ' + tourSteps.length;
  tourEls.title.innerHTML = step.title;
  tourEls.text.innerHTML = step.text;
  tourEls.prev.classList.toggle('hidden', tourIndex === 0);
  tourEls.next.textContent = tourIndex === tourSteps.length - 1 ? 'はじめる！' : 'つぎへ →';

  const target = step.target && document.querySelector(step.target);
  // 対象があるときは対象のまわりを明るく残し、ないときは全体を暗くして中央に表示
  tourEls.dim.classList.toggle('hidden', !!target);
  if (target) {
    target.classList.add('tour-focus');
    // 説明の吹き出しが下に入るよう、対象を画面の上のほうへ
    const top = target.getBoundingClientRect().top + window.scrollY - 90;
    window.scrollTo({ top: Math.max(0, top), behavior: 'smooth' });
  }
  placeBubble();
  setTimeout(placeBubble, 350);
  setTimeout(placeBubble, 800);
  tourEls.next.focus({ preventScroll: true });
}

function placeBubble() {
  if (!tourEls || tourEls.layer.classList.contains('hidden')) return;
  const b = tourEls.bubble;
  const target = document.querySelector('.tour-focus');
  const margin = 16;
  const bw = Math.min(380, window.innerWidth - margin * 2);
  b.style.width = bw + 'px';
  if (!target) {
    b.style.left = (window.innerWidth - bw) / 2 + 'px';
    b.style.top = Math.max(margin, (window.innerHeight - b.offsetHeight) / 2) + 'px';
    return;
  }
  const r = target.getBoundingClientRect();
  const bh = b.offsetHeight;
  let top;
  if (r.bottom + 22 + bh < window.innerHeight - margin) top = r.bottom + 22;      // 下に置ける
  else if (r.top - 22 - bh > margin) top = r.top - 22 - bh;                          // 上に置ける
  else top = window.innerHeight - bh - margin;                                        // 画面の下に重ねる
  let left = Math.min(Math.max(margin, r.left), window.innerWidth - bw - margin);
  // 横に余白があれば、対象の右側に並べる
  if (r.right + 22 + bw < window.innerWidth - margin && r.height > bh) {
    left = r.right + 22;
    top = Math.min(Math.max(margin, r.top), window.innerHeight - bh - margin);
  }
  b.style.left = left + 'px';
  b.style.top = top + 'px';
}
