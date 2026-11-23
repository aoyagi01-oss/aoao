// 課題提出バーコード管理のテスト：node kadai-barcode/test/run.js
// （Apps Script には貼らないでください。スプレッドシートのまねをして Code.gs を動かします）
const assert = require('assert');
const { load } = require('./mock');

const tests = [];
const test = (name, fn) => tests.push({ name, fn });

// 2026-10-20（火）12:00（日本時間）として動かす
const NOW = '2026-10-20T03:00:00Z';

function fresh(opts) {
  const app = load(Object.assign({ now: NOW }, opts));
  app.call('setup_', false);
  app.call('importRoster', [
    { gakuseki: '1101', name: '青木 さくら' }, { gakuseki: '1102', name: '石田 はると' }, { gakuseki: '1201', name: '江藤 そうた' },
  ]);
  return app;
}

function addTask(app, t) {
  return app.call('addTask', Object.assign({ subject: '数学', name: 'ワーク', classes: [], given: '2026-10-13', due: '2026-10-19', late: '遅れとして記録' }, t));
}

const logRows = (app) => app.sheet('提出記録').data.slice(1).filter((r) => r.some((v) => v !== ''));

test('読み取り：提出・2回目・未完成・再提出・取り消し', () => {
  const app = fresh();
  const t = addTask(app, { due: '2026-10-30' });
  let r = app.call('recordScan', t.id, '1101', 'submit');
  assert.strictEqual(r.kind, 'ok');
  assert.strictEqual(r.stats.submitted, 1);
  assert.strictEqual(r.stats.target, 3);
  r = app.call('recordScan', t.id, '1101', 'submit');
  assert.strictEqual(r.kind, 'dup');
  r = app.call('recordScan', t.id, '1102', 'incomplete');
  assert.strictEqual(r.kind, 'incomplete');
  assert.strictEqual(r.stats.submitted, 1);
  r = app.call('recordScan', t.id, '1102', 'submit');
  assert.ok(/再提出/.test(r.message));
  assert.strictEqual(r.stats.submitted, 2);
  r = app.call('recordScan', t.id, '1102', 'undo');
  assert.strictEqual(r.kind, 'undone');
  assert.strictEqual(r.stats.submitted, 1);
  assert.strictEqual(logRows(app).length, 1);
  // 書いた記録と、読み直した記録が同じになっている
  assert.strictEqual(app.call('getTaskStats', t.id).submitted, 1);
});

test('読み取りで、教務手帳のその生徒の行が書きかわる', () => {
  const app = fresh();
  const t = addTask(app, { due: '2026-10-30' });
  app.call('refreshAll');
  app.call('recordScan', t.id, '1201', 'submit');
  const st = app.sheet('提出状況');
  const row = st.data.find((r) => r[1] === '1201');
  assert.ok(row[3] instanceof Date, '提出した日付が入る');
  assert.strictEqual(row[4], 1); // 提出数
});

test('遅れ：締切の次の日からの提出は遅れ', () => {
  const app = fresh();
  const t = addTask(app, { due: '2026-10-19' });
  const r = app.call('recordScan', t.id, '1101', 'submit');
  assert.strictEqual(r.kind, 'late');
});

test('除外の生徒のバーコードは記録しない', () => {
  const app = fresh();
  const t = addTask(app);
  app.sheet('生徒名簿').put(2, 7, true);
  const id = app.sheet('生徒名簿').cell(2, 1);
  const r = app.call('recordScan', t.id, id, 'submit');
  assert.strictEqual(r.kind, 'error');
  assert.ok(/除外/.test(r.message));
  assert.strictEqual(logRows(app).length, 0);
});

test('直前の1件を取り消す：ほかの先生が読んだ記録は消さない', () => {
  const app = fresh();
  app.as('A'); app.call('saveViewSettings', { teacher: '青柳' });
  const t = addTask(app, { due: '2026-10-30' });
  app.call('recordScan', t.id, '1101', 'submit');
  app.as('B'); app.call('saveViewSettings', { teacher: '佐藤' });
  app.call('recordScan', t.id, '1102', 'submit');
  app.as('A');
  let r = app.call('undoLast', t.id);
  assert.strictEqual(r.kind, 'undone');
  assert.ok(/青木/.test(r.message), r.message);
  const rows = logRows(app);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0][10], '佐藤');
  r = app.call('undoLast', t.id);
  assert.strictEqual(r.kind, 'error');
  assert.strictEqual(logRows(app).length, 1);
});

test('教務手帳に手で書く：まとめて記録し、記録を読むのは1回', () => {
  const app = fresh();
  const t = addTask(app, { due: '2026-10-30' });
  app.call('refreshAll');
  const st = app.sheet('提出状況');
  const rowOf = (gk) => st.data.findIndex((r) => r[1] === gk) + 1;
  const r1 = rowOf('1101');
  assert.strictEqual(rowOf('1102'), r1 + 1);
  st.put(r1, 4, '○');
  st.put(r1 + 1, 4, '△');
  const before = app.ss.stats.cellsRead;
  app.call('onEdit', { range: app.ss.getSheetByName('提出状況').getRange(r1, 4, 2, 1) });
  const rows = logRows(app);
  assert.strictEqual(rows.length, 2);
  assert.strictEqual(rows[1][9], '未完成');
  // 消すと取り消し
  st.put(r1, 4, '');
  app.call('onEdit', { range: app.ss.getSheetByName('提出状況').getRange(r1, 4) });
  assert.strictEqual(logRows(app).length, 1);
  assert.ok(app.ss.stats.cellsRead - before < 5000);
});

test('教務手帳に日付を書くと、その日の提出', () => {
  const app = fresh();
  const t = addTask(app, { due: '2026-10-19' });
  app.call('refreshAll');
  const st = app.sheet('提出状況');
  const r = st.data.findIndex((x) => x[1] === '1101') + 1;
  st.put(r, 4, new Date('2026-10-16T01:00:00Z'));
  app.call('onEdit', { range: app.ss.getSheetByName('提出状況').getRange(r, 4) });
  assert.strictEqual(logRows(app)[0][9], '期限内');
});

test('祝日：振替休日・国民の休日・春分・秋分', () => {
  const app = load({ now: NOW });
  const h = (y) => Object.keys(app.ctx.jpHolidays_(y)).sort();
  const y26 = h(2026);
  ['2026-01-12', '2026-03-20', '2026-05-06', '2026-07-20', '2026-09-21', '2026-09-22', '2026-09-23', '2026-10-12', '2026-11-23'].forEach((k) => assert.ok(y26.includes(k), k));
  assert.ok(!y26.includes('2026-05-07'));
  assert.ok(h(2025).includes('2025-11-24')); // 11/23（日）の振替
  assert.ok(h(2027).includes('2027-03-21'));
});

test('◯日まで期限内：土日・祝日・設定した休みの日をとばして数える', () => {
  const app = fresh();
  // 10/9（金）締切・1日まで期限内 → 10/10土・10/11日・10/12スポーツの日をとばして 10/13
  assert.strictEqual(app.call('addSchoolDays_', '2026-10-09', 1), '2026-10-13');
  app.call('saveViewSettings', { holidays: '10/13, 12/25〜1/7' });
  assert.strictEqual(app.call('addSchoolDays_', '2026-10-09', 1), '2026-10-14');
  assert.strictEqual(app.call('addSchoolDays_', '2026-12-24', 1), '2027-01-08');
  assert.strictEqual(app.call('getHomeData').view.holidays, '10/13, 12/25〜1/7');
});

test('休みの日の書き方', () => {
  const app = load({ now: NOW });
  const keys = (v, today) => Object.keys(app.ctx.parseHolidays_(v, today || '2026-10-20')).sort();
  assert.deepStrictEqual(keys('3/25〜3/27'), ['2027-03-25', '2027-03-26', '2027-03-27']);
  assert.deepStrictEqual(keys('2026/11/2、１１月４日'), ['2026-11-02', '2026-11-04']);
  assert.deepStrictEqual(keys('4/1', '2027-02-01'), ['2026-04-01']);
  assert.deepStrictEqual(keys('あいう, 13/40'), []);
});

test('シートの点検は版ごとに1回だけ', () => {
  const app = fresh();
  assert.ok(app.docProps.getProperty('setupVersion'));
  const before = app.ss.stats.reads;
  app.call('setup_', false);
  assert.ok(app.ss.stats.reads - before <= 1, 'ほとんど読まない');
  // シートを消したら作り直す
  app.ss.sheets = app.ss.sheets.filter((s) => s.name !== '欠席・配慮');
  app.call('setup_', false);
  assert.ok(app.sheet('欠席・配慮'));
});

test('旧版の提出記録（10列）に「読み取った先生」の列を足す', () => {
  const app = fresh();
  const log = app.sheet('提出記録');
  log.put(1, 11, '');
  app.docProps.setProperty('setupVersion', 'old');
  app.call('setup_', false);
  assert.strictEqual(log.cell(1, 11), '読み取った先生');
});

test('記録が多くても、1回の読み取りで記録を読むのは1回', () => {
  const app = fresh();
  const ids = [];
  for (let i = 0; i < 30; i++) ids.push(addTask(app, { name: '課題' + i, due: '2026-10-30' }).id);
  app.call('refreshAll');
  ids.forEach((id) => { app.call('recordScan', id, '1101', 'submit'); app.call('recordScan', id, '1102', 'submit'); });
  const logSheet = app.sheet('提出記録');
  let logReads = 0;
  const orig = logSheet.getRange.bind(logSheet);
  logSheet.getRange = (r, c, nr, nc) => { if (r === 2 && c === 1 && nr > 1) logReads++; return orig(r, c, nr, nc); };
  app.call('recordScan', ids[0], '1201', 'submit');
  assert.strictEqual(logReads, 1);
});

let failed = 0;
tests.forEach(({ name, fn }) => {
  try { fn(); console.log('✓ ' + name); } catch (e) { failed++; console.log('✗ ' + name + '\n  ' + (e && e.stack || e)); }
});
console.log(failed ? failed + ' 件失敗' : 'すべて成功（' + tests.length + ' 件）');
process.exit(failed ? 1 : 0);
