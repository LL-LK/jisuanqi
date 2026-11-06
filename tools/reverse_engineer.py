# -*- coding: utf-8 -*-
"""反向推导基线口径：把各种候选计数口径与基线教室合计并列，找出基线用的那一种。"""

from __future__ import annotations

import glob
import os
import sys
from collections import defaultdict

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_exempt import load_exemptions      # noqa: E402

EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}


import re

def strip_count(label):
    return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(label or '')).strip()


def college_of(ws, row):
    col = 2
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.min_row <= row <= rng.max_row:
            return strip_count(ws.cell(rng.min_row, rng.min_col).value)
    return strip_count(ws.cell(row, col).value)


def baseline_sums(baseline):
    out = {}
    wb = load_workbook(baseline, data_only=True)
    for sheet in ('雁山（早）', '雁山（晚）'):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        agg = defaultdict(int)
        for row in range(4, ws.max_row + 1):
            college = college_of(ws, row)
            room = ws.cell(row, 3).value
            if not college or room is None:
                continue
            for day, col in EXEMPT_COLS.items():
                v = ws.cell(row, col).value
                if isinstance(v, (int, float)):
                    agg[(college, day)] += int(v)
        out[sheet] = agg
    wb.close()
    return out


def find_exempt():
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '免勤人员统计表' in os.path.basename(path):
            return path
    return None


def main() -> int:
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    base = baseline_sums(baseline)
    records = load_exemptions(find_exempt())

    # 候选口径
    cand = defaultdict(lambda: defaultdict(set))
    for rec in records:
        if rec.campus != '雁山':
            continue
        for day in rec.weekdays:
            if not rec.sid:
                continue
            # A: 周次含 7 且 星期含该天
            if 7 in rec.weeks:
                cand[('A_含第7周', rec.session, rec.college)][day].add(rec.sid)
            # B: 不看周次，只看星期（但记录本身属于晚/早表）
            cand[('B_忽略周次', rec.session, rec.college)][day].add(rec.sid)
            # C: 星期含该天，且该学院该天的所有记录去重（跨早/晚表合并）
            cand[('C_合并早晚表', rec.college)][day].add(rec.sid)
    # D: 学院整体去重（不分天）
    cand_d = defaultdict(set)
    for rec in records:
        if rec.campus == '雁山' and rec.sid and 7 in rec.weeks:
            cand_d[(rec.session, rec.college)] |= set(rec.sid)

    for sheet in ('雁山（早）', '雁山（晚）'):
        session = '早自习' if '早' in sheet else '晚自习'
        print('\n===== %s =====' % sheet)
        keys = sorted(base.get(sheet, {}).keys())
        for (college, day) in keys:
            b = base[sheet][(college, day)]
            a = len(cand[('A_含第7周', session, college)].get(day, ()))
            bb = len(cand[('B_忽略周次', session, college)].get(day, ()))
            cc = len(cand[('C_合并早晚表', college)].get(day, ()))
            dd = len(cand_d.get((session, college), ()))
            flag = '  <== 基线' if b in (a, bb, cc, dd) else ''
            if b != a:
                print('  %-22s %s 基线=%-4d A含7周=%-4d B忽略周次=%-4d C合并早晚=%-4d D学院整体=%-4d%s'
                      % (college, day, b, a, bb, cc, dd, flag))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
