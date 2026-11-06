# -*- coding: utf-8 -*-
"""逐教室对拍某个学院：把基线写下的免勤人数与明细表里的班级/人数标注并排，找出基线的口径。"""

from __future__ import annotations

import glob
import os
import re
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)

EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}
RATE_COLS = {'周一': 14, '周三': 15, '周四': 16, '周五': 17}


def main() -> int:
    sheet = sys.argv[1] if len(sys.argv) > 1 else '雁山（晚）'
    college_want = sys.argv[2] if len(sys.argv) > 2 else '商学院'
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    src = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '明细表' in os.path.basename(p):
            src = p
    outs = sorted(glob.glob(os.path.join(WORKSPACE, '_work', '输出', '*.xlsx')))

    wb_src = load_workbook(src, data_only=True)
    ws_src = wb_src[sheet]
    wb_base = load_workbook(baseline, data_only=True)
    ws_base = wb_base[sheet]
    wb_mine = load_workbook(outs[-1], data_only=True) if outs else None
    ws_mine = wb_mine[sheet] if wb_mine else None

    print('=== %s / %s ===' % (sheet, college_want))
    print('%-4s %-7s %-46s %-5s %s' % ('行', '教室', '分配专业班级', '应到', '基线免勤(周一/周三/周四/周五)  本次免勤'))
    current = None
    for row in range(3, ws_src.max_row + 1):
        label = ws_src.cell(row, 2).value
        if label:
            current = re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(label)).strip()
        if current != college_want:
            continue
        room = ws_src.cell(row, 3).value
        classes = ws_src.cell(row, 4).value
        expected = ws_src.cell(row, 5).value
        base_vals = [ws_base.cell(row, c).value for c in EXEMPT_COLS.values()]
        mine_vals = [ws_mine.cell(row, c).value for c in EXEMPT_COLS.values()] if ws_mine else []
        print('%-4d %-7s %-46s %-5s %s | %s'
              % (row, room, str(classes)[:46], expected, base_vals, mine_vals))
    wb_src.close(); wb_base.close()
    if wb_mine:
        wb_mine.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
