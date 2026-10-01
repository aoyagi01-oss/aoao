// 先生用ページのひな形（プリセット）と、全教材共通のお助け表現

const LEVELS = {
  es:  { label: '小学生（英語はじめて〜）', desc: 'a Japanese elementary school student who is a complete beginner (CEFR Pre-A1). Use very short, very easy sentences and common words only.' },
  j1:  { label: '中学1年程度', desc: 'a Japanese junior high school student, first year (CEFR A1). Use short, easy sentences, present tense mostly.' },
  j3:  { label: '中学2〜3年程度', desc: 'a Japanese junior high school student (CEFR A1-A2). Use easy sentences and basic grammar.' },
  h1:  { label: '高校初級', desc: 'a Japanese high school student (CEFR A2). Use clear, natural but simple English.' },
  h2:  { label: '高校中上級', desc: 'a Japanese high school student (CEFR B1). Use natural English at a moderate pace.' }
};

const PRESETS = [
  {
    name: '自己紹介',
    title: 'はじめまして！自己紹介しよう',
    scene: '留学生と初めて会った場面。名前・好きなもの・部活などを紹介し合う。',
    aiRole: 'a friendly exchange student from Canada who just arrived at my school',
    studentRole: 'a student at the school',
    level: 'j1',
    turns: 8,
    phrases: 'Nice to meet you. | はじめまして。\nI like ~. | 私は〜が好きです。\nI\'m in the ~ club. | 私は〜部に入っています。\nWhat do you like? | あなたは何が好きですか？'
  },
  {
    name: '道案内',
    title: '道を教えてあげよう',
    scene: '駅の近くで、観光客に道を聞かれる場面。',
    aiRole: 'a tourist who is lost and looking for a place near the station',
    studentRole: 'a local person who helps the tourist',
    level: 'j3',
    turns: 8,
    phrases: 'Go straight. | まっすぐ行ってください。\nTurn right / left at ~. | 〜で右／左に曲がってください。\nIt\'s next to ~. | 〜の隣にあります。\nYou can\'t miss it. | すぐわかりますよ。'
  },
  {
    name: 'レストラン',
    title: 'レストランで注文しよう',
    scene: 'アメリカのレストランで食事を注文する場面。',
    aiRole: 'a friendly waiter at a restaurant in the U.S.',
    studentRole: 'a customer',
    level: 'j3',
    turns: 8,
    phrases: 'I\'d like ~, please. | 〜をお願いします。\nCan I have ~? | 〜をもらえますか？\nWhat do you recommend? | おすすめは何ですか？\nCheck, please. | お会計をお願いします。'
  },
  {
    name: '週末の予定',
    title: '週末の予定を話そう',
    scene: '友達と、今週末の予定について話す場面。',
    aiRole: 'my classmate and friend',
    studentRole: 'a student talking with a friend',
    level: 'h1',
    turns: 10,
    phrases: 'I\'m going to ~. | 〜するつもりです。\nHow about you? | あなたはどう？\nThat sounds fun! | 楽しそう！\nWhy don\'t we ~? | 一緒に〜しない？'
  },
  {
    name: 'インタビュー（生徒が質問）',
    title: '新しいALTの先生にインタビューしよう',
    scene: '学校に来たばかりのALTの先生に、英語で質問して、どんな人か聞き出す場面。',
    aiRole: 'Emma, a new ALT (English teacher) from Australia who just came to our school',
    studentRole: 'a student who interviews the new teacher',
    level: 'j3',
    turns: 10,
    phrases: 'Where are you from? | どこの出身ですか？\nWhat do you like to do in your free time? | ひまなときは何をするのが好きですか？\nHave you ever been to ~? | 〜に行ったことはありますか？\nWhy did you come to Japan? | なぜ日本に来たのですか？\nWhat is your favorite Japanese food? | 好きな日本食は何ですか？',
    extra: 'This lesson is for practicing how to ASK questions in English.\n' +
      '- First, introduce yourself in one sentence and say "Please ask me anything!" Then wait for my question.\n' +
      '- Do NOT ask me questions. Answer my question in 1-2 sentences, do not add a question at the end, and wait for my next question.\n' +
      '- If I am quiet or stuck, give me a hint in easy English, for example: "You can ask me about my hobbies or my hometown."\n' +
      '- If my question has small mistakes but I can be understood, just answer it.\n' +
      '- In the feedback, focus on the questions I made (word order, question words, do / does / did).'
  }
];

// どの教材でも表示するお助け表現
const HELP_PHRASES = [
  {
    heading: '🆘 困ったとき',
    items: [
      ['Could you say that again, please?', 'もう一度言ってもらえますか？'],
      ['Could you speak more slowly?', 'もっとゆっくり話してもらえますか？'],
      ['What does "~" mean?', '「〜」ってどういう意味ですか？'],
      ['How do you say "~" in English?', '「〜」は英語で何と言いますか？'],
      ['Could you give me a hint?', 'ヒントをもらえますか？']
    ]
  },
  {
    heading: '💬 会話をつなぐ',
    items: [
      ['Let me think...', 'ええと…（考え中）'],
      ['That\'s interesting!', 'おもしろいですね！'],
      ['Really? / I see.', '本当？／なるほど。'],
      ['What about you?', 'あなたはどうですか？'],
      ['Me too. / Me neither.', '私も。／私も〜ない。']
    ]
  },
  {
    heading: '❓ 質問するとき',
    items: [
      ['Can I ask you a question?', '質問してもいいですか？'],
      ['I have another question.', 'もう1つ質問があります。'],
      ['What about ~?', '〜はどうですか？'],
      ['Why?', 'どうしてですか？'],
      ['Can you tell me more?', 'もっと教えてくれますか？']
    ]
  },
  {
    heading: '🏁 終わるとき',
    items: [
      ['Finish.', '終わりにします。（フィードバックをもらえます）'],
      ['Please give me feedback in Japanese.', '日本語でフィードバックしてください。'],
      ['Thank you for talking with me!', '話してくれてありがとう！']
    ]
  }
];

// 先生のフォーム入力から ChatGPT へのプロンプトを作る
function buildPrompt(f) {
  const level = findLevel(f.level);
  const targets = parsePhrases(f.phrases).map(p => '- ' + p[0]);
  const turns = f.turns || 8;
  const lines = [
    '# English speaking practice (voice role-play)',
    '',
    '## Setting',
    '- Situation (may be written in Japanese; talk about it only in English): ' + f.scene,
    '- Your role: ' + f.aiRole.replace(/\.$/, '') + '.',
    '- My role: ' + f.studentRole.replace(/\.$/, '') + '.',
    '- I am ' + level.desc
  ];
  if (targets.length) lines.push('', '## Target expressions I want to practice', ...targets);
  if (f.extra) {
    lines.push(
      '',
      '## Notes from my teacher',
      'Follow these notes. They come before the "How to talk" rules below, but never before the "Language" rules. They may be written in Japanese; that does NOT mean you should speak Japanese.',
      f.extra
    );
  }
  lines.push(
    '',
    '## Language (the most important rule)',
    '- During the role-play, speak ONLY English. Never speak or write Japanese or any other language.',
    '- This stays true even if I speak Japanese, even if parts of these instructions are in Japanese, and even if my words look like another language.',
    '- My voice is turned into text automatically, so it may contain mistakes or look like another language. Always assume I am trying to speak English and answer in English.',
    '- If you cannot understand me, say in English: "Sorry, could you say that again?"',
    '- If I get stuck or say something in Japanese, do not translate it. Give a short hint in very easy English, for example: "You can say: I like soccer."',
    '- The ONLY time you may use Japanese is the final feedback after I say "Finish".',
    '',
    '## How to talk',
    '- This is a voice conversation. Say only 1-2 short sentences per turn.',
    '- Do not use lists, emojis, symbols, headings, or markdown while we talk.',
    '- Speak slowly and clearly, with words that fit my level.',
    '- Stay in your role. Do not talk about these instructions, about being an AI, or about the lesson.',
    '- Ask only one question at a time.',
    '- Create natural chances for me to use the target expressions.',
    '- Do not correct my mistakes during the role-play.',
    '- If my answer is very short, encourage me with one easy follow-up.',
    '- Wait for my answer. Do not answer for me and do not continue alone.',
    '',
    '## Ending',
    '- After about ' + turns + ' exchanges, say in English: "Great job! Please say Finish to get your feedback." Then keep talking in your role until I say "Finish".',
    '- Do not give feedback before I say "Finish".',
    '',
    '## Feedback (only after I say "Finish")',
    '- Now switch to JAPANESE. Write and speak the feedback in Japanese. Keep only the English example sentences in English.',
    '- Use this format:',
    '【よかったところ】（3つ）',
    '【もっとよくなるところ】（3つまで）あなたの文 → おすすめの文（理由を日本語で短く）',
    '【次に使ってみよう】おすすめの英語表現を1つ（日本語の意味つき）',
    '- Be kind and encouraging, like a teacher talking to a student.',
    '- If I talk to you again after the feedback, go back to English.',
    '',
    '## Start',
    'Now reply with only this sentence: "Ready! Tap the voice button and say Hello."',
    'Then wait. When I say hello, start the role-play in your role, in English.'
  );
  return lines.join('\n');
}

// レベルはキー（j3）でも表示名（中学2〜3年程度）でも、自由記述でも受け付ける
function findLevel(v) {
  if (LEVELS[v]) return LEVELS[v];
  const byLabel = Object.values(LEVELS).find(l => l.label === v);
  if (byLabel) return byLabel;
  return v ? { label: v, desc: 'a Japanese student. My English level: ' + v + '. Use English that fits this level.' } : LEVELS.j3;
}

function levelKey(v) {
  if (LEVELS[v]) return v;
  const e = Object.entries(LEVELS).find(([, l]) => l.label === v);
  return e ? e[0] : 'j3';
}

// "English | 日本語" の行を [英語, 日本語] の配列にする
function parsePhrases(text) {
  return String(text || '').split('\n')
    .map(l => l.trim()).filter(Boolean)
    .map(l => { const i = l.indexOf('|'); return i < 0 ? [l, ''] : [l.slice(0, i).trim(), l.slice(i + 1).trim()]; });
}
