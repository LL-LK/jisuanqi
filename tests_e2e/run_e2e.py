# -*- coding: utf-8 -*-
"""Autonomous E2E Test Suite Runner.
Runs the 4-Tier requirement-driven E2E test suite for the Attendance Tool (草莓小熊 / jisuanqi).

Usage:
    python tests_e2e/run_e2e.py [--tier {1,2,3,4,all}] [-v] [--report-json <path>]

Exit Code:
    0 = All executed tests succeeded (or known pending milestone defects caught cleanly)
    1 = Test failures / unexpected errors encountered
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import time
from pathlib import Path

# Add project root to sys.path
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest

TIER_FILES = {
    1: [
        "tests_e2e/test_tier1_r1_storage.py",
        "tests_e2e/test_tier1_r2_tamper.py",
        "tests_e2e/test_tier1_r3_audit.py",
        "tests_e2e/test_tier1_r4_tray.py",
        "tests_e2e/test_tier1_r5_mobile_excel.py",
        "tests_e2e/test_tier1_r6_packaging.py",
    ],
    2: [
        "tests_e2e/test_tier2_boundary_corner.py",
    ],
    3: [
        "tests_e2e/test_tier3_cross_feature.py",
    ],
    4: [
        "tests_e2e/test_tier4_real_world.py",
    ],
}


class E2ETestPlugin:
    """Pytest plugin to capture structured test results per tier."""

    def __init__(self):
        self.results = []
        self.start_time = 0.0
        self.duration = 0.0

    def pytest_sessionstart(self, session):
        self.start_time = time.time()

    def pytest_sessionfinish(self, session, exitstatus):
        self.duration = time.time() - self.start_time

    def pytest_runtest_logreport(self, report):
        if report.when == "call" or (report.when in ("setup", "teardown") and report.failed):
            nodeid = report.nodeid
            # Determine Tier
            tier = 0
            if "tier1" in nodeid:
                tier = 1
            elif "tier2" in nodeid:
                tier = 2
            elif "tier3" in nodeid:
                tier = 3
            elif "tier4" in nodeid:
                tier = 4

            # Determine Requirement if Tier 1
            req = "N/A"
            for r in ("r1", "r2", "r3", "r4", "r5", "r6"):
                if f"_{r}_" in nodeid or f"/{r}_" in nodeid:
                    req = r.upper()
                    break

            msg = ""
            if report.failed and report.longrepr:
                msg = str(report.longrepr).splitlines()[-1] if hasattr(report.longrepr, '__iter__') else str(report.longrepr)

            self.results.append({
                "nodeid": nodeid,
                "tier": tier,
                "requirement": req,
                "outcome": report.outcome,
                "duration": getattr(report, "duration", 0.0),
                "error_message": msg,
            })


def print_banner():
    print("=" * 80)
    print("  草莓小熊 / jisuanqi · 4-Tier E2E Test Suite Autonomous Runner")
    print("  Authoritative Requirements: ORIGINAL_REQUEST.md & PROJECT.md (R1 - R6)")
    print("=" * 80)


def print_summary_table(plugin: E2ETestPlugin):
    total = len(plugin.results)
    passed = sum(1 for r in plugin.results if r["outcome"] == "passed")
    failed = sum(1 for r in plugin.results if r["outcome"] == "failed")
    skipped = sum(1 for r in plugin.results if r["outcome"] == "skipped")

    print("\n" + "-" * 80)
    print("  E2E TEST EXECUTION SUMMARY (4-TIER BREAKDOWN)")
    print("-" * 80)
    print(f"  {'Tier / Category':<35} | {'Total':<6} | {'Passed':<6} | {'Failed':<6} | {'Skipped':<7}")
    print("  " + "-" * 76)

    tier_names = {
        1: "Tier 1: Feature Coverage (R1-R6)",
        2: "Tier 2: Boundary & Corner Cases",
        3: "Tier 3: Cross-Feature Combinations",
        4: "Tier 4: Real-World Scenarios",
    }

    for tier_id in (1, 2, 3, 4):
        items = [r for r in plugin.results if r["tier"] == tier_id]
        t_tot = len(items)
        t_pass = sum(1 for r in items if r["outcome"] == "passed")
        t_fail = sum(1 for r in items if r["outcome"] == "failed")
        t_skip = sum(1 for r in items if r["outcome"] == "skipped")
        t_name = tier_names.get(tier_id, f"Tier {tier_id}")
        print(f"  {t_name:<35} | {t_tot:<6} | {t_pass:<6} | {t_fail:<6} | {t_skip:<7}")

    print("  " + "-" * 76)
    print(f"  {'TOTAL':<35} | {total:<6} | {passed:<6} | {failed:<6} | {skipped:<7}")
    print(f"  Total Duration: {plugin.duration:.2f}s\n")

    # If there are failures, list them as implementation defect escalations
    failures = [r for r in plugin.results if r["outcome"] == "failed"]
    if failures:
        print("=" * 80)
        print("  DISCOVERED IMPLEMENTATION DEFECTS TO ESCALATE (QA AUDIT)")
        print("=" * 80)
        for idx, f in enumerate(failures, 1):
            short_id = f["nodeid"].split("::")[-1]
            err = f["error_message"].strip()
            # Truncate long trace
            if len(err) > 120:
                err = err[:117] + "..."
            print(f"  [{idx:02d}] {short_id}")
            print(f"       -> {err}")
        print("=" * 80)


def main():
    parser = argparse.ArgumentParser(description="Autonomous E2E Test Suite Runner")
    parser.add_argument("--tier", choices=["1", "2", "3", "4", "all"], default="all",
                        help="Select tier to run (default: all)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose pytest output")
    parser.add_argument("--report-json", type=str, default="", help="Write JSON summary report")
    parser.add_argument("--junit-xml", type=str, default="", help="Write JUnit XML report")

    args = parser.parse_args()

    print_banner()

    target_files = []
    if args.tier == "all":
        for files in TIER_FILES.values():
            target_files.extend(files)
    else:
        target_files = TIER_FILES[int(args.tier)]

    # Filter to existing files
    existing_targets = []
    for f in target_files:
        full_p = ROOT / f
        if full_p.exists():
            existing_targets.append(str(full_p))
        else:
            print(f"[WARN] Target test file not found: {f}")

    if not existing_targets:
        print("[ERROR] No test files found to run!")
        return 1

    pytest_args = [
        "-q" if not args.verbose else "-v",
        "--tb=short",
        "-o", "asyncio_default_fixture_loop_scope=session",
        *existing_targets,
    ]

    if args.junit_xml:
        pytest_args.extend(["--junitxml", args.junit_xml])

    plugin = E2ETestPlugin()
    exit_code = pytest.main(pytest_args, plugins=[plugin])

    print_summary_table(plugin)

    if args.report_json:
        report_data = {
            "summary": {
                "total": len(plugin.results),
                "passed": sum(1 for r in plugin.results if r["outcome"] == "passed"),
                "failed": sum(1 for r in plugin.results if r["outcome"] == "failed"),
                "skipped": sum(1 for r in plugin.results if r["outcome"] == "skipped"),
                "duration": plugin.duration,
                "exit_code": exit_code,
            },
            "results": plugin.results,
        }
        with open(args.report_json, "w", encoding="utf-8") as jf:
            json.dump(report_data, jf, ensure_ascii=False, indent=2)
        print(f"  JSON report written to: {args.report_json}")

    return exit_code


if __name__ == "__main__":
    sys.exit(main())
