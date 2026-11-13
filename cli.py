# -*- coding: utf-8 -*-
"""命令行入口：与桌面界面共用同一套计算引擎。

用法：
    # 完整计算：
    python cli.py --detail 明细表.xlsx --exempt 免勤表.xlsx --week 7 --out 输出.xlsx

    # 结构化 JSON 输出：
    python cli.py --detail 明细表.xlsx --exempt 免勤表.xlsx --week 7 --out 输出.xlsx --json

    # 前置自检：
    python cli.py --preflight 明细表.xlsx 免勤表.xlsx [--json]

    # 独立免勤分项导出：
    python cli.py --exempt-report 免勤报表.xlsx --exempt 免勤表.xlsx --week 7 [--json]

    # 防篡改校验与原始明细实到对拍：
    python cli.py --verify-tamper 结果.xlsx [--against-detail 原明细.xlsx] [--json]
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional
from zipfile import BadZipFile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core import RunConfig, run          # noqa: E402
from core.excel_safety import (
    DiagnosticMasker,
    _sanitize_anomaly_text,
    generate_diagnostic_log,
    preflight_excel_file,
)
from core.exempt_report import (
    export_exempt_report,
    export_standalone_exempt_report,
    reconcile_backfill_discrepancies,
)
from core.export import (
    check_output_path,
    compare_actuals,
    diff_actuals_detailed,
    verify_workbook_integrity,
)
from core.io_detail import infer_workbook_week
from core.pipeline import output_path_for, summary_path_for
from core.storage import AttendanceStorage

SID_PATTERN = re.compile(r'(?<![0-9A-Fa-f])(?!99)\d{13}(?![0-9A-Fa-f])')


def _print_json(data: Any) -> None:
    """标准 JSON 输出，强制执行 13 位学号零泄露断言。"""
    json_str = json.dumps(data, ensure_ascii=False, indent=2)
    assert not SID_PATTERN.search(json_str), "Privacy violation: unmasked 13-digit student ID detected in JSON output"
    print(json_str)


def _is_json_mode(argv=None) -> bool:
    args_list = argv if argv is not None else sys.argv[1:]
    return '--json' in args_list


def main(argv=None) -> int:
    t_start = time.perf_counter()
    json_mode = _is_json_mode(argv)

    class JsonArgumentParser(argparse.ArgumentParser):
        def error(self, message):
            if json_mode:
                _print_json({
                    'action': 'cli_error',
                    'status': 'error',
                    'code': 1,
                    'error': _sanitize_anomaly_text(message),
                    'duration_ms': int((time.perf_counter() - t_start) * 1000),
                })
                sys.exit(1)
            super().error(message)

    parser = JsonArgumentParser(description='三早一晚免勤自动统计系统')
    parser.add_argument('--json', action='store_true', help='以标准化 JSON 格式输出结果')
    parser.add_argument('--preflight', nargs='+', default=None, metavar='FILE',
                        help='执行前置 Excel 健康自检与表型自动识别')
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
    parser.add_argument('--campus', default='雁山', help='计入统计的校区，当前模板支持雁山')
    parser.add_argument('--match-classes', action='store_true', help='无可匹配教室号时先尝试按班级推定教室')
    parser.add_argument('--college-exempt', action='store_true', help='按学院回填免勤（合并免勤列的模板自动使用）')
    parser.add_argument('--include-exempt-summary', action='store_true', help='将免勤统计一并写入明细结果工作簿')
    parser.add_argument('--complete-mode', choices=['auto', 'strict', 'tolerant'], default='auto',
                        help='出勤率完整计算模式：auto (按词典), strict (严格契约), tolerant (容错完整)')
    parser.add_argument('--summary-out', default=None, help='独立汇总终稿工作簿输出路径（.xlsx）')
    parser.add_argument('--no-storage', action='store_true', help='禁用 SQLite 历史计算批次记录')
    parser.add_argument('--list-runs', action='store_true', help='查询最近历史计算批次记录')
    parser.add_argument('--show-run', type=int, default=None, help='查询指定 ID 的计算批次详情')
    parser.add_argument('--diagnostic-log', default=None, help='导出脱敏技术诊断报告路径（.md）')
    args = parser.parse_args(argv)

    # 1. 前置健康自检与表型推断 (--preflight)
    if args.preflight is not None:
        reports = [preflight_excel_file(p) for p in args.preflight]
        for r in reports:
            r.pop('raw_path', None)
        all_valid = all(r['is_valid'] for r in reports)
        code = 0 if all_valid else 1
        duration_ms = int((time.perf_counter() - t_start) * 1000)
        if args.json:
            _print_json({
                'action': 'preflight',
                'status': 'ok' if all_valid else 'error',
                'code': code,
                'duration_ms': duration_ms,
                'reports': reports,
            })
            return code

        print('=== 前置 Excel 健康诊断与表型自检 ===')
        for r in reports:
            status_tag = '✅ 正常' if r['is_valid'] else f"❌ 异常 ({r['error_code']})"
            print(f"[{status_tag}] {r['path']} -> 类型: {r['table_type']} | {r['user_message']}")
            if r['tech_detail']:
                print(f"   技术细节: {r['tech_detail']}")
        return code

    # 2. 防篡改签名校验与实到人数逐格对拍 (--verify-tamper / --against-detail)
    if args.against_detail and not args.verify_tamper:
        parser.error('--against-detail 必须与 --verify-tamper 同时使用')

    if args.verify_tamper:
        diffs = []
        violations = []
        try:
            ok, violations = verify_workbook_integrity(args.verify_tamper)
            if args.against_detail:
                diffs = diff_actuals_detailed(args.against_detail, args.verify_tamper)
                if diffs:
                    violations.extend(f"{d['cell']} 实到人数与原表不符" for d in diffs)
                    ok = False
                elif not args.json:
                    print('实到人数逐格核对通过')
        except (OSError, ValueError, BadZipFile) as exc:
            if args.json:
                _print_json({
                    'action': 'verify_tamper',
                    'status': 'error',
                    'code': 1,
                    'duration_ms': int((time.perf_counter() - t_start) * 1000),
                    'error': _sanitize_anomaly_text(str(exc)),
                })
                return 1
            parser.error(str(exc))

        code = 0 if ok else 2
        status = 'ok' if ok else 'tamper_detected'
        duration_ms = int((time.perf_counter() - t_start) * 1000)
        if args.json:
            _print_json({
                'action': 'verify_tamper',
                'status': status,
                'code': code,
                'duration_ms': duration_ms,
                'target': args.verify_tamper,
                'against_detail': args.against_detail,
                'verified': ok,
                'violations_count': len(violations),
                'violations': [_sanitize_anomaly_text(v) for v in violations],
                'diffs': diffs,
            })
            return code

        if ok:
            print('防篡改校验通过（%s）' % args.verify_tamper)
            return 0
        print('防篡改校验失败，发现 %d 处异常：' % len(violations))
        for violation in violations:
            print('  - %s' % violation)
        return 2

    # 2.5 历史运行记录查询 (--list-runs / --show-run)
    if args.list_runs:
        storage = AttendanceStorage()
        runs = storage.list_runs(limit=30)
        masked_runs = []
        for r in runs:
            item = dict(r)
            for pkey in ['detail_path', 'output_path', 'summary_path']:
                if pkey in item and item[pkey]:
                    item[pkey] = DiagnosticMasker.mask_path(item[pkey])
            masked_runs.append(item)

        if args.json:
            _print_json({
                'action': 'list_runs',
                'status': 'ok',
                'code': 0,
                'duration_ms': int((time.perf_counter() - t_start) * 1000),
                'total_runs': len(masked_runs),
                'runs': masked_runs,
            })
            return 0

        print(f"=== 历史计算记录 (最近 {len(masked_runs)} 条) ===")
        if not masked_runs:
            print("暂无计算记录。")
        else:
            print(f"{'ID':<6} {'时间':<20} {'周次':<6} {'校区':<8} {'状态':<8} {'耗时(ms)':<10} {'明细文件'}")
            print("-" * 80)
            for r in masked_runs:
                print(f"{r.get('id', ''):<6} {str(r.get('created_at', ''))[:19]:<20} 第{r.get('week_no', '')}周 {r.get('campus', ''):<8} {r.get('status', ''):<8} {r.get('duration_ms', 0):<10} {r.get('detail_path', '')}")
        return 0

    if args.show_run is not None:
        storage = AttendanceStorage()
        run_record = storage.get_run(args.show_run)
        if not run_record:
            if args.json:
                _print_json({
                    'action': 'show_run',
                    'status': 'not_found',
                    'code': 1,
                    'error': f'Run ID {args.show_run} not found',
                    'duration_ms': int((time.perf_counter() - t_start) * 1000),
                })
                return 1
            print(f"未找到批次记录 ID: {args.show_run}")
            return 1

        masked_record = dict(run_record)
        for pkey in ['detail_path', 'output_path', 'summary_path']:
            if pkey in masked_record and masked_record[pkey]:
                masked_record[pkey] = DiagnosticMasker.mask_path(masked_record[pkey])

        if args.json:
            _print_json({
                'action': 'show_run',
                'status': 'ok',
                'code': 0,
                'duration_ms': int((time.perf_counter() - t_start) * 1000),
                'run': masked_record,
            })
            return 0

        print(f"=== 计算批次详情 (ID: {args.show_run}) ===")
        for k, v in masked_record.items():
            print(f"  {k}: {v}")
        return 0

    # 3. 独立免勤报表生成 (--exempt-report 且无 --detail)
    if args.exempt_report and not args.detail:
        if not args.exempt:
            parser.error('--exempt-report 独立运行模式必须指定至少一份 --exempt 免勤表')
        try:
            week = args.week or 7
            rep_result = export_standalone_exempt_report(
                output_path=args.exempt_report,
                exempt_paths=args.exempt,
                week_no=week,
                dictionary_path=args.dictionary,
                campus=args.campus,
                source=args.source,
            )
        except (OSError, ValueError, BadZipFile) as exc:
            if args.json:
                _print_json({
                    'action': 'exempt_report_standalone',
                    'status': 'error',
                    'code': 1,
                    'duration_ms': int((time.perf_counter() - t_start) * 1000),
                    'error': _sanitize_anomaly_text(str(exc)),
                })
                return 1
            parser.error(str(exc))

        duration_ms = int((time.perf_counter() - t_start) * 1000)
        if args.json:
            _print_json({
                'action': 'exempt_report_standalone',
                'status': 'ok',
                'code': 0,
                'duration_ms': duration_ms,
                'output_path': args.exempt_report,
                'metrics': rep_result,
            })
            return 0

        print(f"独立免勤报表导出成功: {args.exempt_report}（共 {rep_result['rows_count']} 条分项记录，免勤合计 {rep_result['total_exempt']} 人次）")
        return 0

    # 4. 完整计算流水线
    if not args.detail:
        parser.error('the following arguments are required: --detail')

    try:
        week = args.week or infer_workbook_week(args.detail) or 7
    except (OSError, ValueError, BadZipFile) as exc:
        if args.json:
            _print_json({
                'action': 'run',
                'status': 'error',
                'code': 1,
                'duration_ms': int((time.perf_counter() - t_start) * 1000),
                'error': _sanitize_anomaly_text(str(exc)),
            })
            return 1
        parser.error(str(exc))

    complete_mode_map = {
        'tolerant': True,
        'strict': False,
        'auto': None,
    }
    complete_mode_val = complete_mode_map.get(args.complete_mode, None)

    config = RunConfig(
        detail_path=args.detail,
        exempt_paths=tuple(args.exempt),
        output_path=args.out,
        week_no=week,
        exempt_source=args.source,
        dictionary_path=args.dictionary,
        campus=args.campus,
        complete_mode=complete_mode_val,
        match_classes=args.match_classes,
        college_exempt=args.college_exempt,
        include_exempt_summary=args.include_exempt_summary,
        enable_storage=not args.no_storage,
        summary_output_path=args.summary_out,
    )

    if args.exempt_report:
        try:
            check_output_path(args.exempt_report, (config.detail_path, *config.exemption_files,
                                                   config.dictionary_path, output_path_for(config), summary_path_for(config)))
        except ValueError as exc:
            if args.json:
                _print_json({
                    'action': 'run',
                    'status': 'error',
                    'code': 1,
                    'duration_ms': int((time.perf_counter() - t_start) * 1000),
                    'error': _sanitize_anomaly_text(str(exc)),
                })
                return 1
            parser.error(str(exc))

    result = run(config)

    if args.diagnostic_log:
        try:
            diag_text = generate_diagnostic_log(config, result=result)
            diag_dir = os.path.dirname(os.path.abspath(args.diagnostic_log))
            if diag_dir:
                os.makedirs(diag_dir, exist_ok=True)
            with open(args.diagnostic_log, 'w', encoding='utf-8') as f:
                f.write(diag_text)
        except Exception:
            pass

    if not result.ok:
        if args.json:
            out_err = {
                'action': 'run',
                'status': 'error',
                'code': 1,
                'duration_ms': result.duration_ms or int((time.perf_counter() - t_start) * 1000),
                'error': _sanitize_anomaly_text(result.error),
            }
            if args.diagnostic_log:
                out_err['diagnostic_log'] = args.diagnostic_log
            _print_json(out_err)
            return 1
        print('计算失败：%s' % result.error)
        if args.diagnostic_log:
            print('诊断日志已生成：%s' % args.diagnostic_log)
        return 1

    if args.exempt_report:
        try:
            export_exempt_report(args.exempt_report, result)
        except (OSError, ValueError) as exc:
            if args.json:
                _print_json({
                    'action': 'run',
                    'status': 'error',
                    'code': 1,
                    'duration_ms': result.duration_ms,
                    'error': _sanitize_anomaly_text(f'明细结果已生成（{result.output_path}），免勤统计表导出失败: {exc}'),
                    'output_file': result.output_path,
                })
                return 1
            print('明细结果已生成（%s），免勤统计表导出失败：%s' % (result.output_path, exc))
            return 1

    if args.json:
        college_rates_list = []
        for rec in result.college_rates:
            college_rates_list.append({
                'sheet': rec.sheet,
                'college': rec.college,
                'day': rec.day,
                'expected': rec.expected,
                'actual': rec.actual,
                'exempt': rec.exempt,
                'rate': rec.rate,
                'status': rec.status,
            })
        weekly_rates_list = []
        for rec in result.weekly_rates:
            weekly_rates_list.append({
                'sheet': rec.sheet,
                'college': rec.college,
                'rate': rec.rate,
                'calculated_days': rec.calculated_days,
                'scheduled_days': rec.scheduled_days,
            })
        discrepancies = reconcile_backfill_discrepancies(result.exempt_rows)

        def _sanitize_anomaly_item(item: Dict[str, Any]) -> Dict[str, Any]:
            cleaned = {}
            for k, v in item.items():
                if isinstance(v, str):
                    cleaned[k] = _sanitize_anomaly_text(v)
                elif isinstance(v, dict):
                    cleaned[k] = _sanitize_anomaly_item(v)
                elif isinstance(v, list):
                    cleaned[k] = [
                        _sanitize_anomaly_item(elem) if isinstance(elem, dict)
                        else (_sanitize_anomaly_text(elem) if isinstance(elem, str) else elem)
                        for elem in v
                    ]
                else:
                    cleaned[k] = v
            return cleaned

        sanitized_anomalies = [_sanitize_anomaly_item(a) for a in result.anomalies]

        data = {
            'action': 'run',
            'status': 'ok',
            'code': 0,
            'duration_ms': result.duration_ms or int((time.perf_counter() - t_start) * 1000),
            'phases': result.phase_durations,
            'output_file': result.output_path,
            'summary_output_file': result.summary_output_path,
            'exempt_report': args.exempt_report,
            'summary': result.summary,
            'exempt_stats': result.exempt_stats,
            'college_rates': college_rates_list,
            'weekly_rates': weekly_rates_list,
            'discrepancies_count': len(discrepancies),
            'discrepancies': discrepancies,
            'anomalies_count': len(result.anomalies),
            'anomalies': sanitized_anomalies,
        }
        if args.diagnostic_log:
            data['diagnostic_log'] = args.diagnostic_log
        _print_json(data)
        return 0

    if args.exempt_report:
        print('免勤统计表：%s' % args.exempt_report)

    if args.diagnostic_log:
        print('诊断报告已导出：%s' % args.diagnostic_log)

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
