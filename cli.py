# -*- coding: utf-8 -*-
"""命令行入口：与桌面界面共用同一套计算引擎。

用法：
    python cli.py --detail 明细表.xlsx --exempt 免勤表.xlsx --week 7 --out 输出.xlsx
"""

from __future__ import annotations

import argparse
import os
import sys
from zipfile import BadZipFile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import RunConfig, run          # noqa: E402
from core.exempt_report import export_exempt_report
from core.export import check_output_path, compare_actuals, verify_workbook_integrity
from core.io_detail import infer_workbook_week
from core.pipeline import output_path_for, summary_path_for


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description='三早一晚免勤自动统计系统')
    parser.add_argument('--verify-tamper', default=None, help='校验指定结果 Excel 文件的防篡改签名完整性')
    parser.add_argument('--against-detail', help='防篡改校验时与可信原始明细逐格核对实到人数')
    parser.add_argument('--detail', required=False, help='执行情况明细表路径')
    parser.add_argument('--exempt', action='extend', nargs='+', default=[], help='一个或多个免勤表路径，可重复使用')
    parser.add_argument('--out', default=None, help='输出文件路径')
    parser.add_argument('--exempt-report', default=None, help='另行输出免勤统计工作簿（.xlsx）')
    parser.add_argument('--week', type=int, choices=range(1, 31), metavar='1–30',
                        help='统计周次，省略时优先读取表内标题，缺省第 7 周')
    parser.add_argument('--source', default='未分类',
                        choices=['全免勤大表', '在线存档', '未分类'],
                        help='免勤表来源分类')
    parser.add_argument('--dictionary', default=None, help='免勤词典路径')
    parser.add_argument('--campus', default='雁山', help='计入统计的校区')
    parser.add_argument('--match-classes', action='store_true', help='无可匹配教室号时先尝试按班级推定教室')
    parser.add_argument('--college-exempt', action='store_true', help='按学院回填免勤（合并免勤列的模板自动使用）')
    parser.add_argument('--include-exempt-summary', action='store_true', help='将免勤统计一并写入明细结果工作簿')
    args = parser.parse_args(argv)

    if args.against_detail and not args.verify_tamper:
        parser.error('--against-detail 必须与 --verify-tamper 同时使用')
    if args.verify_tamper:
        try:
            ok, violations = verify_workbook_integrity(args.verify_tamper)
            if args.against_detail:
                changes = compare_actuals(args.against_detail, args.verify_tamper)
                if changes:
                    violations.extend(f'{cell} 实到人数与原表不符' for cell in changes)
                else:
                    print('实到人数逐格核对通过')
                ok = ok and not changes
        except (OSError, ValueError, BadZipFile) as exc:
            parser.error(str(exc))
        if ok:
            print('防篡改校验通过（%s）' % args.verify_tamper)
            return 0
        print('防篡改校验失败，发现 %d 处异常：' % len(violations))
        for violation in violations:
            print('  - %s' % violation)
        return 2

    if not args.detail:
        parser.error('the following arguments are required: --detail')
    try:
        week = args.week or infer_workbook_week(args.detail) or 7
    except (OSError, ValueError, BadZipFile) as exc:
        parser.error(str(exc))

    config = RunConfig(
        detail_path=args.detail,
        exempt_paths=tuple(args.exempt),
        output_path=args.out,
        week_no=week,
        exempt_source=args.source,
        dictionary_path=args.dictionary,
        campus=args.campus,
        match_classes=args.match_classes,
        college_exempt=args.college_exempt,
        include_exempt_summary=args.include_exempt_summary,
    )
    if args.exempt_report:
        try:
            check_output_path(args.exempt_report, (config.detail_path, *config.exemption_files,
                                                   config.dictionary_path, output_path_for(config), summary_path_for(config)))
        except ValueError as exc:
            parser.error(str(exc))
    result = run(config)
    if not result.ok:
        print('计算失败：%s' % result.error)
        return 1
    if args.exempt_report:
        try:
            export_exempt_report(args.exempt_report, result)
        except (OSError, ValueError) as exc:
            print('明细结果已生成（%s），免勤统计表导出失败：%s' % (result.output_path, exc))
            return 1
        print('免勤统计表：%s' % args.exempt_report)

    print('=== 汇总 ===')
    for key, value in result.summary.items():
        print('%-26s %s' % (key, value))
    print('\n=== 免勤构成 ===')
    for key, value in result.exempt_stats.items():
        print('%-26s %s' % (key, value))
    print('\n输出文件：%s' % result.output_path)

    print('\n=== 异常清单（前 30 条）===')
    for item in result.anomalies[:30]:
        print('[%s] %s | %s | %s' % (item['level'], item['type'], item['scope'], item['note']))
    if len(result.anomalies) > 30:
        print('... 其余 %d 条见界面异常面板' % (len(result.anomalies) - 30))

    print('\n=== 学院出勤率 ===')
    for rec in result.college_rates:
        rate = '%.2f%%' % (rec.rate * 100) if rec.rate is not None else '—'
        print('%-10s %-16s %-4s 应到=%-6s 实到=%-6s 免勤=%-5s 出勤率=%-8s %s'
              % (rec.sheet, rec.college, rec.day, rec.expected, rec.actual,
                 rec.exempt, rate, rec.status))
    print('\n=== 学院每周出勤率 ===')
    for rec in result.weekly_rates:
        rate = '%.2f%%' % (rec.rate * 100) if rec.rate is not None else '—'
        print(f'{rec.sheet} / {rec.college}：{rate}（{rec.calculated_days}/{rec.scheduled_days} 天）')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
