# -*- coding: utf-8 -*-
"""在真实样本上跑一遍引擎，用于快速验证与对拍。

用法：
    python tools/quickrun.py                 # 自动定位工作区里的明细表与免勤表
    python tools/quickrun.py --compare       # 与 Week7_Attendance_Final.xlsx 对拍
"""

from __future__ import annotations

import argparse
import glob
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core import RunConfig, run          # noqa: E402
from core.io_detail import load_detail   # noqa: E402


def find_inputs():
    detail = None
    exempt = None
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        name = os.path.basename(path)
        if '明细表' in name and detail is None:
            detail = path
        elif '免勤人员统计表' in name and exempt is None:
            exempt = path
    return detail, exempt


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--compare', action='store_true')
    parser.add_argument('--week', type=int, default=7)
    parser.add_argument('--out', default=None)
    args = parser.parse_args()

    detail_path, exempt_path = find_inputs()
    if not detail_path:
        print('找不到明细表')
        return 1
    print('明细表：', os.path.basename(detail_path))
    print('免勤表：', os.path.basename(exempt_path) if exempt_path else '(未找到)')

    config = RunConfig(
        detail_path=detail_path,
        exempt_path=exempt_path,
        output_path=args.out,
        week_no=args.week,
        dictionary_path=os.path.join(ROOT, '免勤词典.xlsx'),
    )
    result = run(config)
    if not result.ok:
        print('计算失败：', result.error)
        return 1

    print('\n=== 汇总 ===')
    for key, value in result.summary.items():
        print('%-26s %s' % (key, value))
    print('\n=== 免勤构成 ===')
    for key, value in result.exempt_stats.items():
        print('%-26s %s' % (key, value))
    print('\n输出：', result.output_path)

    print('\n=== 异常清单（前 25 条）===')
    for item in result.anomalies[:25]:
        print('[%s] %s | %s | %s' % (item['level'], item['type'], item['scope'], item['note']))

    print('\n=== 学院出勤率（按工作表/学院/星期）===')
    for rec in result.college_rates:
        rate = '%.2f%%' % (rec.rate * 100) if rec.rate is not None else '—'
        print('%-10s %-18s %-3s 应到=%-5s 实到=%-5s 免勤=%-5s 率=%-8s %s'
              % (rec.sheet, rec.college, rec.day, rec.expected, rec.actual,
                 rec.exempt, rate, rec.status))

    if args.compare:
        compare(result.output_path)
    return 0


def compare(output_path: str) -> None:
    """与上一轮产物 Week7_Attendance_Final.xlsx 逐单元格对拍（仅比较免勤列与出勤率区）。"""
    from openpyxl import load_workbook
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    if not os.path.exists(baseline):
        print('\n[对拍] 基线文件不存在，跳过')
        return
    a = load_workbook(output_path, data_only=True)
    b = load_workbook(baseline, data_only=True)
    print('\n=== 与 Week7_Attendance_Final.xlsx 对拍 ===')
    for sheet in b.sheetnames:
        if sheet not in a.sheetnames:
            print('  [%s] 输出缺少该表' % sheet)
            continue
        wa, wb_ = a[sheet], b[sheet]
        diff_exempt = 0
        diff_rate = 0
        samples = []
        for r in range(1, max(wa.max_row, wb_.max_row) + 1):
            for c in range(1, max(wa.max_column, wb_.max_column) + 1):
                va, vb = wa.cell(r, c).value, wb_.cell(r, c).value
                if va == vb:
                    continue
                header = wb_.cell(2, c).value or wb_.cell(3, c).value or ''
                if '免勤人数' in str(header):
                    diff_exempt += 1
                elif c >= 14 and c <= 18:
                    diff_rate += 1
                if len(samples) < 12:
                    samples.append((r, c, va, vb))
        print('  [%s] 免勤列差异 %d，N:R 出勤率区差异 %d' % (sheet, diff_exempt, diff_rate))
        for s in samples:
            print('      r%s c%s 本次=%r 基线=%r' % s)
    a.close()
    b.close()


if __name__ == '__main__':
    raise SystemExit(main())
