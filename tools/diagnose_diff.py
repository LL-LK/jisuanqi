# -*- coding: utf-8 -*-
"""归因诊断：把"免勤列差异"拆成两类——总量差异（去重/过滤口径）还是分摊差异。

对每个 (工作表, 学院, 星期)：
    本次学院总免勤   = 引擎按唯一学号去重算出的学院总免勤
    本次教室合计     = 本次输出里该学院各教室免勤之和
    基线教室合计     = 基线文件里该学院各教室免勤之和
若"本次学院总免勤 == 基线教室合计"，说明差异纯粹是分摊方式不同；
否则说明去重/过滤口径本身有分歧。
"""

from __future__ import annotations

import glob
import os
import sys
from collections import defaultdict

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core import RunConfig, run                      # noqa: E402
from core.io_detail import load_detail              # noqa: E402

EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}
RATE_COLS = {'周一': 14, '周三': 15, '周四': 16, '周五': 17}


def college_of(ws, row):
    import re as _re
    col = 2
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.min_row <= row <= rng.max_row:
            raw = ws.cell(rng.min_row, rng.min_col).value
            return _re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(raw or '')).strip()
    raw = ws.cell(row, col).value
    return _re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(raw or '')).strip()


def room_sums(path, sheet):
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = defaultdict(int)
    for row in range(4, ws.max_row + 1):
        college = college_of(ws, row)
        room = ws.cell(row, 3).value
        if not college or room is None:
            continue
        for day, col in EXEMPT_COLS.items():
            v = ws.cell(row, col).value
            if isinstance(v, (int, float)):
                out[(college, day)] += int(v)
    wb.close()
    return out


def main() -> int:
    outs = sorted(glob.glob(os.path.join(WORKSPACE, '_work', '输出', '*.xlsx')))
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    if not outs:
        print('缺少本次输出')
        return 1

    detail_path = None
    exempt_path = None
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        name = os.path.basename(path)
        if '明细表' in name:
            detail_path = path
        elif '免勤人员统计表' in name:
            exempt_path = path

    result = run(RunConfig(detail_path=detail_path, exempt_path=exempt_path,
                           week_no=7,
                           dictionary_path=os.path.join(ROOT, '免勤词典.xlsx')))
    if not result.ok:
        print('运行失败', result.error)
        return 1

    engine_total = {(r.sheet, r.college, r.day): r.exempt for r in result.college_rates}

    for sheet in ('雁山（早）', '雁山（晚）'):
        mine = room_sums(outs[-1], sheet)
        base = room_sums(baseline, sheet)
        print('\n===== %s =====' % sheet)
        total_diff = 0
        dist_diff = 0
        rows = []
        for key in sorted(set(mine) | set(base)):
            college, day = key
            eng = engine_total.get((sheet, college.replace('（', '(').replace('）', ')'), day))
            if eng is None:
                for (s, c, d), v in engine_total.items():
                    if s == sheet and d == day and c.replace('（', '(').replace('）', ')') == college.replace('（', '(').replace('）', ')'):
                        eng = v
                        break
            m, b = mine.get(key, 0), base.get(key, 0)
            if m == b:
                continue
            if eng is not None and eng == b:
                dist_diff += abs(m - b)
                rows.append(('分摊差异', college, day, eng, m, b))
            else:
                total_diff += abs(m - b)
                rows.append(('总量差异', college, day, eng, m, b))
        for kind, college, day, eng, m, b in rows:
            print('  [%s] %-18s %s 引擎学院总免勤=%s 本次教室合计=%s 基线教室合计=%s'
                  % (kind, college, day, eng, m, b))
        print('  → 涉及 %d 个 (学院,星期) 组合：其中总量差异 %d、分摊差异 %d'
              % (len(rows), total_diff, dist_diff))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
