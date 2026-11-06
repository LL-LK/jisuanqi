# -*- coding: utf-8 -*-
"""列出免勤表里所有学院名写法及其记录数，找出与明细表学院名的变体/错字。"""

from __future__ import annotations

import glob
import os
import re
import sys
from collections import Counter, defaultdict

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)


def strip_count(label):
    return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(label or '')).strip()


def main() -> int:
    exempt_path = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '免勤人员统计表' in os.path.basename(p):
            exempt_path = p
    src = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '明细表' in os.path.basename(p):
            src = p

    wb = load_workbook(exempt_path, data_only=True)
    counter = Counter()
    for sn in wb.sheetnames:
        ws = wb[sn]
        for r in range(3, ws.max_row + 1):
            v = ws.cell(r, 3).value
            if v:
                counter[str(v).strip()] += 1
    wb.close()

    wb_src = load_workbook(src, data_only=True)
    detail_names = set()
    for sn in wb_src.sheetnames:
        ws = wb_src[sn]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 2).value
            if v:
                detail_names.add(strip_count(v))
    wb_src.close()

    print('明细表学院名 %d 个：' % len(detail_names))
    for n in sorted(detail_names):
        print('   %s' % n)

    print('\n免勤表学院名 %d 种写法：' % len(counter))
    for name, cnt in counter.most_common():
        flag = '' if name in detail_names else '   <== 明细表没有此写法'
        print('  %-30s %5d%s' % (repr(name), cnt, flag))

    print('\n--- 归一化后能否对上 ---')
    norm_counter = Counter()
    for name, cnt in counter.items():
        norm_counter[strip_count(name).replace(' ', '')] += cnt
    for name, cnt in sorted(norm_counter.items(), key=lambda x: -x[1]):
        flag = '' if name in {d.replace(' ', '') for d in detail_names} else '   <== 仍对不上'
        print('  %-30s %5d%s' % (name, cnt, flag))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
