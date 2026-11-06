# -*- coding: utf-8 -*-
"""严格对拍：把 None 与 '' 视为等价，只报真实数值差异，并给出定位上下文。"""

from __future__ import annotations

import glob
import os
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)


def norm(v):
    if v is None:
        return ''
    if isinstance(v, str):
        return v.strip()
    return v


def college_of(ws, row):
    for col in (2,):
        for rng in ws.merged_cells.ranges:
            if rng.min_col <= col <= rng.max_col and rng.min_row <= row <= rng.max_row:
                return ws.cell(rng.min_row, rng.min_col).value
        return ws.cell(row, col).value
    return ''


def main() -> int:
    outs = [p for p in glob.glob(os.path.join(WORKSPACE, '_work', '输出', '*.xlsx'))
            if not os.path.basename(p).startswith('_explore')]
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    if not outs or not os.path.exists(baseline):
        print('缺少对比文件')
        return 1
    # 取**最新**的一份产物，而不是文件名排序最后的一份
    out_path = max(outs, key=os.path.getmtime)
    print('本次：', os.path.basename(out_path))
    print('基线：', os.path.basename(baseline))

    a = load_workbook(out_path, data_only=True)
    b = load_workbook(baseline, data_only=True)

    for sheet in b.sheetnames:
        if sheet not in a.sheetnames:
            print('  [%s] 输出缺少该表' % sheet)
            continue
        wa, wb_ = a[sheet], b[sheet]
        real = []
        for r in range(1, max(wa.max_row, wb_.max_row) + 1):
            for c in range(1, max(wa.max_column, wb_.max_column) + 1):
                va, vb = norm(wa.cell(r, c).value), norm(wb_.cell(r, c).value)
                if va == vb:
                    continue
                if isinstance(va, float) and isinstance(vb, float) and abs(va - vb) < 1e-9:
                    continue
                real.append((r, c, wa.cell(r, c).value, wb_.cell(r, c).value))
        print('\n  [%s] 真实差异 %d 处' % (sheet, len(real)))
        for r, c, va, vb in real[:40]:
            head2 = norm(wb_.cell(2, c).value)
            head3 = norm(wb_.cell(3, c).value)
            col_label = head2 or head3 or ''
            print('     %s%d  本次=%r  基线=%r  | 列=%s | 行归属=%r'
                  % (chr(64 + c) if c <= 26 else c, r, va, vb, col_label,
                     college_of(wb_, r)))
    a.close()
    b.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
