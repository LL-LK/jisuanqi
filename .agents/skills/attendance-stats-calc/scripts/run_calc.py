# -*- coding: utf-8 -*-
"""Privacy-enforcing calculation runner wrapper for attendance-stats-calc.

This wrapper invokes the core calculation engine or cli.py while strictly
scrubbing student personal data (student IDs, individual names). It only emits
clean aggregate college-level summaries, weekly rates, anomaly counts, status
codes, and generated file paths.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
from typing import Any, Dict, List, Optional, Sequence

# 13-digit student ID pattern (safeguard against hexadecimal hash substrings)
SID_REGEX = re.compile(r'(?<![0-9A-Fa-f])\d{13}(?![0-9A-Fa-f])')


def discover_app_root(custom_path: Optional[str] = None) -> Optional[str]:
    """Dynamically discover the root directory containing cli.py and core package."""
    candidates = []
    if custom_path:
        candidates.append(os.path.abspath(custom_path))

    env_root = os.environ.get("ATTENDANCE_APP_ROOT")
    if env_root:
        candidates.append(os.path.abspath(env_root))

    cwd = os.getcwd()
    candidates.append(cwd)

    # Check if skill resides inside <repo>/.agents/skills/attendance-stats-calc/scripts/
    script_repo = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "..", ".."))
    candidates.append(script_repo)

    # Search upward and downward from cwd
    for depth in range(1, 4):
        up = os.path.abspath(os.path.join(cwd, *(['..'] * depth)))
        candidates.append(up)
        candidates.append(os.path.join(up, "免勤统计系统"))
        candidates.append(os.path.join(up, "ruanjian", "免勤统计系统"))

    candidates.append(os.path.join(cwd, "免勤统计系统"))
    candidates.append(os.path.join(cwd, "ruanjian", "免勤统计系统"))

    # Fallback to standard installation paths if not found via cwd climbing
    for d_letter in ("D", "d"):
        d_prefix = d_letter + ":"
        candidates.append(os.path.join(d_prefix, os.sep, "Software", "ruanjian", "免勤统计系统"))
        candidates.append(os.path.join(d_prefix, os.sep, "Software", "三早一晚免勤自动统计"))

    for cand in candidates:
        if not os.path.isdir(cand):
            continue
        cli_file = os.path.join(cand, "cli.py")
        core_dir = os.path.join(cand, "core")
        if os.path.isfile(cli_file) and os.path.isdir(core_dir):
            return cand

    return None


def sanitize_anomalies(raw_anomalies: Sequence[Dict[str, Any]]) -> Dict[str, Any]:
    """Aggregate anomalies by type and level without personal names or student details."""
    by_type: Dict[str, int] = {}
    by_level: Dict[str, int] = {}
    for item in raw_anomalies:
        atype = str(item.get("type", "UNKNOWN"))
        level = str(item.get("level", "INFO"))
        by_type[atype] = by_type.get(atype, 0) + 1
        by_level[level] = by_level.get(level, 0) + 1

    return {
        "total_anomalies": len(raw_anomalies),
        "by_type": by_type,
        "by_level": by_level,
    }


def assert_privacy_compliance(payload_str: str) -> None:
    """Internal privacy assertion: guarantee zero 13-digit student IDs in output."""
    match = SID_REGEX.search(payload_str)
    if match:
        raise ValueError(
            f"Privacy assertion violation: 13-digit student ID detected in output text (index {match.start()})"
        )


def execute_runner(args: argparse.Namespace) -> int:
    """Execute requested operation and emit sanitized aggregate JSON."""
    t_start = time.perf_counter()
    app_root = discover_app_root(args.app_root)

    if not app_root:
        err_out = {
            "action": "runner_error",
            "status": "error",
            "code": 1,
            "error": "Could not locate application root containing cli.py and core package.",
            "duration_ms": int((time.perf_counter() - t_start) * 1000),
        }
        print(json.dumps(err_out, ensure_ascii=False, indent=2))
        return 1

    if app_root not in sys.path:
        sys.path.insert(0, app_root)

    try:
        from core.excel_safety import preflight_excel_file
        from core.export import verify_workbook_integrity, diff_actuals_detailed
        from core.exempt_report import export_standalone_exempt_report
        from core.pipeline import RunConfig, run as run_pipeline, output_path_for, summary_path_for
        from core.io_detail import infer_workbook_week
    except ImportError as err:
        err_out = {
            "action": "runner_error",
            "status": "error",
            "code": 1,
            "error": f"Failed importing core modules: {err}",
            "duration_ms": int((time.perf_counter() - t_start) * 1000),
        }
        print(json.dumps(err_out, ensure_ascii=False, indent=2))
        return 1

    # Mode 1: Preflight Diagnosis
    if args.preflight is not None:
        reports = [preflight_excel_file(p) for p in args.preflight]
        all_valid = all(r.get("is_valid", False) for r in reports)
        sanitized_reports = []
        for r in reports:
            sanitized_reports.append({
                "path": r.get("path"),
                "is_valid": r.get("is_valid"),
                "table_type": r.get("table_type"),
                "error_code": r.get("error_code"),
                "user_message": r.get("user_message"),
                "file_size_formatted": r.get("file_size_formatted"),
                "sha256_prefix": r.get("sha256_prefix"),
                "sheet_names": r.get("sheet_names", []),
            })
        out_data = {
            "action": "preflight",
            "status": "ok" if all_valid else "error",
            "code": 0 if all_valid else 1,
            "duration_ms": int((time.perf_counter() - t_start) * 1000),
            "all_valid": all_valid,
            "reports": sanitized_reports,
        }
        out_str = json.dumps(out_data, ensure_ascii=False, indent=2)
        assert_privacy_compliance(out_str)
        print(out_str)
        return 0 if all_valid else 1

    # Mode 2: Verify Anti-Tamper
    if args.verify_tamper:
        v_ok, v_msg = verify_workbook_integrity(args.verify_tamper)
        diff_count = 0
        diff_summary: List[Dict[str, Any]] = []
        if args.against_detail:
            diffs = diff_actuals_detailed(args.against_detail, args.verify_tamper)
            diff_count = len(diffs)
            for d in diffs:
                diff_summary.append({
                    "sheet": d.get("sheet"),
                    "cell": d.get("cell"),
                    "source_value": d.get("source_value"),
                    "result_value": d.get("result_value"),
                })
        is_tampered = (not v_ok) or (diff_count > 0)
        out_data = {
            "action": "verify_tamper",
            "status": "tamper_detected" if is_tampered else "ok",
            "code": 2 if is_tampered else 0,
            "duration_ms": int((time.perf_counter() - t_start) * 1000),
            "verified": v_ok,
            "message": v_msg,
            "diff_count": diff_count,
            "diffs": diff_summary,
        }
        out_str = json.dumps(out_data, ensure_ascii=False, indent=2)
        assert_privacy_compliance(out_str)
        print(out_str)
        return 2 if is_tampered else 0

    # Mode 3: Standalone Exemption Report (without detail workbook)
    if args.exempt_report and not args.detail:
        if not args.exempt:
            err_data = {
                "action": "exempt_report_standalone",
                "status": "error",
                "code": 1,
                "error": "--exempt argument is required for standalone exemption export",
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
            }
            print(json.dumps(err_data, ensure_ascii=False, indent=2))
            return 1
        week = args.week or 7
        try:
            rep_res = export_standalone_exempt_report(
                output_path=args.exempt_report,
                exempt_paths=args.exempt,
                week_no=week,
                dictionary_path=args.dictionary,
                campus=args.campus,
                source=args.source,
            )
            out_data = {
                "action": "exempt_report_standalone",
                "status": "ok",
                "code": 0,
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
                "output_file": args.exempt_report,
                "rows_count": rep_res.get("rows_count", 0),
                "total_exempt": rep_res.get("total_exempt", 0),
                "anomalies_count": rep_res.get("anomalies_count", 0),
            }
            out_str = json.dumps(out_data, ensure_ascii=False, indent=2)
            assert_privacy_compliance(out_str)
            print(out_str)
            return 0
        except Exception as exc:
            err_data = {
                "action": "exempt_report_standalone",
                "status": "error",
                "code": 1,
                "error": str(exc),
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
            }
            print(json.dumps(err_data, ensure_ascii=False, indent=2))
            return 1

    # Mode 3.5: Run Storage Query (--list-runs / --show-run)
    if args.list_runs:
        try:
            from core.storage import AttendanceStorage
            from core.excel_safety import DiagnosticMasker
            storage = AttendanceStorage()
            runs = storage.list_runs(limit=30)
            masked_runs = []
            for r in runs:
                item = dict(r)
                for pkey in ['detail_path', 'output_path', 'summary_path', 'dictionary_path']:
                    if pkey in item and item[pkey]:
                        item[pkey] = DiagnosticMasker.mask_path(item[pkey])
                masked_runs.append(item)
            out_data = {
                "action": "list_runs",
                "status": "ok",
                "code": 0,
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
                "total_runs": len(masked_runs),
                "runs": masked_runs,
            }
            out_str = json.dumps(out_data, ensure_ascii=False, indent=2)
            assert_privacy_compliance(out_str)
            print(out_str)
            return 0
        except Exception as exc:
            err_data = {
                "action": "list_runs",
                "status": "error",
                "code": 1,
                "error": str(exc),
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
            }
            print(json.dumps(err_data, ensure_ascii=False, indent=2))
            return 1

    if args.show_run is not None:
        try:
            from core.storage import AttendanceStorage
            from core.excel_safety import DiagnosticMasker
            storage = AttendanceStorage()
            run_item = storage.get_run(args.show_run)
            if not run_item:
                err_data = {
                    "action": "show_run",
                    "status": "not_found",
                    "code": 1,
                    "error": f"Run ID {args.show_run} not found",
                    "duration_ms": int((time.perf_counter() - t_start) * 1000),
                }
                print(json.dumps(err_data, ensure_ascii=False, indent=2))
                return 1
            masked_item = dict(run_item)
            for pkey in ['detail_path', 'output_path', 'summary_path', 'dictionary_path']:
                if pkey in masked_item and masked_item[pkey]:
                    masked_item[pkey] = DiagnosticMasker.mask_path(masked_item[pkey])
            out_data = {
                "action": "show_run",
                "status": "ok",
                "code": 0,
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
                "run": masked_item,
            }
            out_str = json.dumps(out_data, ensure_ascii=False, indent=2)
            assert_privacy_compliance(out_str)
            print(out_str)
            return 0
        except Exception as exc:
            err_data = {
                "action": "show_run",
                "status": "error",
                "code": 1,
                "error": str(exc),
                "duration_ms": int((time.perf_counter() - t_start) * 1000),
            }
            print(json.dumps(err_data, ensure_ascii=False, indent=2))
            return 1

    # Mode 4: Full Calculation Pipeline
    if not args.detail:
        err_data = {
            "action": "run",
            "status": "error",
            "code": 1,
            "error": "The --detail argument is required for full calculation run.",
            "duration_ms": int((time.perf_counter() - t_start) * 1000),
        }
        print(json.dumps(err_data, ensure_ascii=False, indent=2))
        return 1

    try:
        week = args.week or infer_workbook_week(args.detail) or 7
    except Exception as exc:
        err_data = {
            "action": "run",
            "status": "error",
            "code": 1,
            "error": f"Failed inferring week number: {exc}",
            "duration_ms": int((time.perf_counter() - t_start) * 1000),
        }
        print(json.dumps(err_data, ensure_ascii=False, indent=2))
        return 1

    complete_mode_map = {'tolerant': True, 'strict': False, 'auto': None}
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

    result = run_pipeline(config)

    if args.diagnostic_log:
        try:
            from core.excel_safety import generate_diagnostic_log
            diag_text = generate_diagnostic_log(config, result=result)
            diag_dir = os.path.dirname(os.path.abspath(args.diagnostic_log))
            if diag_dir:
                os.makedirs(diag_dir, exist_ok=True)
            with open(args.diagnostic_log, "w", encoding="utf-8") as f:
                f.write(diag_text)
        except Exception:
            pass

    if not result.ok:
        err_data = {
            "action": "run",
            "status": "error",
            "code": 1,
            "error": result.error,
            "duration_ms": result.duration_ms or int((time.perf_counter() - t_start) * 1000),
        }
        if args.diagnostic_log:
            err_data["diagnostic_log"] = args.diagnostic_log
        print(json.dumps(err_data, ensure_ascii=False, indent=2))
        return 1

    # Optional standalone exempt report generation
    exempt_report_path = args.exempt_report
    if exempt_report_path:
        try:
            from core.exempt_report import export_exempt_report
            export_exempt_report(exempt_report_path, result)
        except Exception as exc:
            err_data = {
                "action": "run",
                "status": "error",
                "code": 1,
                "error": f"Detail output generated ({result.output_path}), but exemption report failed: {exc}",
                "output_file": result.output_path,
                "duration_ms": result.duration_ms or int((time.perf_counter() - t_start) * 1000),
            }
            print(json.dumps(err_data, ensure_ascii=False, indent=2))
            return 1

    college_rates_list = []
    for rec in result.college_rates:
        college_rates_list.append({
            "sheet": rec.sheet,
            "college": rec.college,
            "day": rec.day,
            "expected": rec.expected,
            "actual": rec.actual,
            "exempt": rec.exempt,
            "rate": f"{rec.rate * 100:.2f}%" if rec.rate is not None else None,
            "status": rec.status,
        })

    weekly_rates_list = []
    for rec in result.weekly_rates:
        weekly_rates_list.append({
            "sheet": rec.sheet,
            "college": rec.college,
            "rate": f"{rec.rate * 100:.2f}%" if rec.rate is not None else None,
            "calculated_days": rec.calculated_days,
            "scheduled_days": rec.scheduled_days,
        })

    anomaly_agg = sanitize_anomalies(result.anomalies)

    out_data = {
        "action": "run",
        "status": "success",
        "code": 0,
        "duration_ms": result.duration_ms or int((time.perf_counter() - t_start) * 1000),
        "stage_durations_ms": result.phase_durations,
        "output_file": result.output_path,
        "summary_output_file": result.summary_output_path,
        "exempt_report": exempt_report_path,
        "summary": result.summary,
        "exempt_stats": result.exempt_stats,
        "college_rates": college_rates_list,
        "weekly_rates": weekly_rates_list,
        "anomalies_count": anomaly_agg["total_anomalies"],
        "anomaly_summary": anomaly_agg,
    }
    if args.diagnostic_log:
        out_data["diagnostic_log"] = args.diagnostic_log

    out_str = json.dumps(out_data, ensure_ascii=False, indent=2)
    assert_privacy_compliance(out_str)
    print(out_str)
    return 0


def build_arg_parser() -> argparse.ArgumentParser:
    """Build command line argument parser."""
    parser = argparse.ArgumentParser(
        description="Privacy-enforcing calculation runner wrapper for attendance-stats-calc"
    )
    parser.add_argument("--detail", help="Path to execution detail workbook (.xlsx)")
    parser.add_argument("--exempt", action="extend", nargs="+", default=[], help="One or more exemption workbooks (.xlsx)")
    parser.add_argument("--out", help="Path to output calculation workbook (.xlsx)")
    parser.add_argument("--exempt-report", help="Path to standalone exemption report workbook (.xlsx)")
    parser.add_argument("--week", type=int, choices=range(1, 31), help="Week number (1-30)")
    parser.add_argument("--dictionary", help="Path to custom dictionary workbook (.xlsx)")
    parser.add_argument("--campus", default="雁山", help="Campus name (default: 雁山)")
    parser.add_argument("--source", default="未分类", choices=["全免勤大表", "在线存档", "未分类"], help="Exemption source category")
    parser.add_argument("--match-classes", action="store_true", help="Match classroom by class schedule if room is missing")
    parser.add_argument("--college-exempt", action="store_true", help="Backfill exemptions by college")
    parser.add_argument("--include-exempt-summary", action="store_true", help="Include exemption summary in detail workbook")
    parser.add_argument("--complete-mode", choices=["auto", "strict", "tolerant"], default="auto", help="Complete calculation mode (auto, strict, tolerant)")
    parser.add_argument("--summary-out", default=None, help="Path to export summary final workbook (*_汇总终稿.xlsx)")
    parser.add_argument("--no-storage", action="store_true", help="Disable SQLite run history persistence")
    parser.add_argument("--list-runs", action="store_true", help="List recent calculation runs from SQLite storage")
    parser.add_argument("--show-run", type=int, default=None, help="Show details of a specific calculation run ID")
    parser.add_argument("--diagnostic-log", default=None, help="Path to export diagnostic markdown report")
    parser.add_argument("--preflight", nargs="+", default=None, metavar="FILE", help="Execute preflight health check on workbooks")
    parser.add_argument("--verify-tamper", default=None, help="Verify anti-tamper signature of result workbook")
    parser.add_argument("--against-detail", default=None, help="Compare actual headcount cell-by-cell against original detail")
    parser.add_argument("--app-root", default=None, help="Explicit path to application root directory")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    parser = build_arg_parser()
    args = parser.parse_args(argv)
    return execute_runner(args)


if __name__ == "__main__":
    sys.exit(main())
