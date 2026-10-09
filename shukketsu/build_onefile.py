#!/usr/bin/env python3
# 1ファイル版を作る：Code.gs と画面の HTML（Take・Setup・Period）を、貼るだけで動く1つのファイルにまとめる
#   使い方：python3 shukketsu/build_onefile.py
#   できるもの：shukketsu/1ファイル版/コード.gs（Code.gs や HTML を直したら、これを実行し直す）
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
PAGES = ['Take', 'Setup', 'Period']
OUT = os.path.join(HERE, '1ファイル版', 'コード.gs')


def read(name):
    with open(os.path.join(HERE, name), encoding='utf-8') as f:
        return f.read()


def build():
    code = read('Code.gs').rstrip('\n')
    parts = [
        '// ─────────────────────────────────────────────',
        '// 授業の出欠（教務手帳）1ファイル版',
        '// このファイルだけを Apps Script の「コード.gs」に貼れば動きます（HTML ファイルを作る必要はありません）。',
        '// shukketsu/build_onefile.py で自動で作ったファイルです。直すときは Code.gs・Take.html などを直して作り直してください。',
        '// ─────────────────────────────────────────────',
        '',
        code,
        '',
        '// ───────── 画面の HTML（Take.html・Setup.html・Period.html の中身） ─────────',
        '',
        'var HTML_FILES = {};',
    ]
    for name in PAGES:
        lines = read(name + '.html').rstrip('\n').split('\n')
        parts.append('HTML_FILES[' + json.dumps(name) + '] = [')
        parts.extend('  ' + json.dumps(line, ensure_ascii=False) + ',' for line in lines)
        parts.append("].join('\\n');")
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('\n'.join(parts) + '\n')
    return OUT


if __name__ == '__main__':
    print(build())
