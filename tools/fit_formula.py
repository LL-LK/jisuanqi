# -*- coding: utf-8 -*-
"""系统性枚举候选口径，找出能复现基线数字的那一种。

候选：
  H2  唯一学号去重（对话口径）
  H3  记录数（不按学号去重）
  H4  唯一学号（该学院该天，跨早晚表合并）
  H5  唯一学号 + '全班' 占位记录按明细表班级人数计入
  H6  唯一学号 + 无学号集体申请人数计入
  H7  唯一学号 + 所有占位/集体记录按明细表班级人数计入
"""

from __future__ import annotations

import glob
import os
import re
import sys
from collections import defaultdict

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_exempt import load_exemptions      # noqa: E402
from core.parsing import normalize_class       # noqa: E402

EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}


def strip_count(label):
    return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(label or '')).strip()


def class_sizes(text):
    """'会计26-1（41）、电商26-3（10）' → {'会计26-1': 41, '电商26-3': 10}"""
    out = {}
    for part in re.split(r'[、,，;；]', str(text or '')):
        m = re.match(r'\s*(.+?)\s*[（(]\s*(\d+)\s*人\s*[)）]\s*$', part)
        if m:
            out[normalize_class(m.group(1))] = int(m.group(2))
    return out


def main() -> int:
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    src = exempt = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        name = os.path.basename(p)
        if '明细表' in name:
            src = p
        elif '免勤人员统计表' in name:
            exempt = p

    wb_src = load_workbook(src, data_only=True)
    wb_base = load_workbook(baseline, data_only=True)

    # 班级人数表（按 学院 → 班级key → 人数）
    sizes = defaultdict(dict)
    base_totals = defaultdict(int)
    for sheet in ('雁山（早）', '雁山（晚）'):
        ws = wb_src[sheet]
        current = None
        for row in range(3, ws.max_row + 1):
            label = ws.cell(row, 2).value
            if label:
                current = strip_count(label)
            if ws.cell(row, 3).value is None:
                continue
            for key, n in class_sizes(ws.cell(row, 4).value).items():
                sizes[current].setdefault(key, n)
        wsb = wb_base[sheet]
        current = None
        for row in range(3, wsb.max_row + 1):
            label = wsb.cell(row, 2).value
            if label:
                current = strip_count(label)
            if wsb.cell(row, 3).value is None:
                continue
            for day, col in EXEMPT_COLS.items():
                v = wsb.cell(row, col).value
                if isinstance(v, (int, float)):
                    base_totals[(sheet, current, day)] += int(v)

    records = load_exemptions(exempt)
    print('班级人数表覆盖学院 %d 个' % len(sizes))

    scores = defaultdict(int)
    total = 0
    detail_rows = []
    for (sheet, college, day), base in sorted(base_totals.items()):
        if base == 0:
            continue
        session = '早自习' if '早' in sheet else '晚自习'
        hits = [r for r in records
                if r.college == college and r.session == session and day in r.weekdays
                and (not r.weeks or 7 in r.weeks)]
        sids = {r.sid for r in hits if r.sid and r.sid != '全班'}
        h2 = len(sids)
        h3 = len(hits)
        h4 = len({r.sid for r in records
                  if r.college == college and day in r.weekdays and r.sid
                  and (not r.weeks or 7 in r.weeks)})
        h5 = h2 + sum(sizes[college].get(r.class_key, 0)
                      for r in hits if r.sid == '全班')
        h6 = h2 + sum(r.group_size or 0 for r in hits if not r.sid)
        h7 = h2 + sum(sizes[college].get(r.class_key, 0)
                      for r in hits if (not r.sid) or r.sid == '全班')
        cand = {'H2': h2, 'H3': h3, 'H4': h4, 'H5': h5, 'H6': h6, 'H7': h7}
        total += 1
        for k, v in cand.items():
            if v == base:
                scores[k] += 1
        detail_rows.append((sheet[2:-1], college, day, base, cand))

    for row in detail_rows:
        sheet, college, day, base, cand = row
        if not any(v == base for v in cand.values()):
            print('  %-3s %-16s %s 基线=%-5d %s' % (sheet, college, day, base, cand))
    print('\n命中统计（共 %d 个组合）：' % total)
    for k in sorted(scores):
        print('  %s 命中 %d / %d' % (k, scores[k], total))
    wb_src.close(); wb_base.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
