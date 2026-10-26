#!/usr/bin/env python3
"""ほかの学校に渡す配布版を dist/ に作る。

使い方: python3 tools/make_dist.py "石川県立羽咋高等学校" "羽咋"
  1つ目: 生徒画面の上に出す学校名
  2つ目: 氏名の入力例に使う姓（例: 羽咋　太郎）

錦丘の学校名・教育目標の言い回し・学校のピクトグラムを外す。Code.gs と Report.html はそのまま写す。
"""
import os, re, sys

here = os.path.dirname(os.path.abspath(__file__))
src_dir = os.path.dirname(here)
out_dir = os.path.join(src_dir, 'dist')
school, family = sys.argv[1], sys.argv[2]

s = open(os.path.join(src_dir, 'Index.html'), encoding='utf-8').read()

def rep(old, new, count):
    global s
    n = s.count(old)
    assert n == count, f'{old!r}: {n} 件（{count} 件のはず）'
    s = s.replace(old, new)

rep('>金沢錦丘中学校・高等学校<', f'>{school}<', 1)
rep('錦丘　太郎', f'{family}　太郎', 3)
# フィードバックの「錦丘高校が目指す「…」」は学校の目標なので、目標の文だけ残す
rep('錦丘高校が目指す', '', 3)
# 学校のピクトグラムは錦丘のものなので外す（各項目の線画アイコンが代わりに出る）
s, n = re.subn(r'const PICTOGRAMS = \{.*?\n    \};', 'const PICTOGRAMS = {};', s, count=1, flags=re.S)
assert n == 1, 'PICTOGRAMS が見つからない'
assert '錦丘' not in s, '錦丘 が残っている'

os.makedirs(out_dir, exist_ok=True)
open(os.path.join(out_dir, 'Index.html'), 'w', encoding='utf-8').write(s)
for name in ('Code.gs', 'Report.html'):
    open(os.path.join(out_dir, name), 'w', encoding='utf-8').write(
        open(os.path.join(src_dir, name), encoding='utf-8').read())
print('dist/ を作りました:', school)
