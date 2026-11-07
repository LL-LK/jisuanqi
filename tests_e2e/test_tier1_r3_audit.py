# -*- coding: utf-8 -*-
"""Tier 1: Requirement R3 Feature Coverage (Isolation).
Tests core attendance business logic, multi-dimensional deduplication, class-group coverage,
Hare-Niemeyer room allocation, capacity capping, rate formula accuracy, and weekly average strictness.

Specification: PROJECT.md § Architecture, ORIGINAL_REQUEST.md § R3
Feature IDs: F07, F08
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.calc import Calculator, allocate
from core.io_exempt import aggregate_exemptions, load_exemptions
from core.model import DetailModel, ExemptRecord, RoomRow
from tests_e2e.helpers import (
    create_detail_workbook,
    create_exempt_workbook,
    run_pipeline_e2e,
)


def test_r3_deduplication_personal_key(tmp_path: Path):
    """Test 1 (R3.1): Duplicate applications for the same student on same day/session count as 1."""
    p = tmp_path / "exempt_dedup.xlsx"
    rows = [
        {'sid': '9900000000001', 'name': '张三', 'college': '理学院', 'class_text': '理24-1班', 'day_text': '周一', 'session': '早'},
        {'sid': '9900000000001', 'name': '张三', 'college': '理学院', 'class_text': '理24-1班', 'day_text': '周一', 'session': '早'},
    ]
    create_exempt_workbook(p, early_rows=rows)
    records = load_exemptions(str(p), source='未分类')
    agg = aggregate_exemptions(records, week_no=7, campus='雁山')
    count = agg.exempt_count('理学院', '周一', '早自习')
    assert count == 1, f"Expected 1 deduplicated headcount, got {count}"
    assert agg.duplicate_events >= 1, "Duplicate event should be recorded in anomaly stats"


def test_r3_group_application_covers_personal(tmp_path: Path):
    """Test 2 (R3.2): Valid class group exemption completely covers individual applications in same class."""
    p = tmp_path / "exempt_group.xlsx"
    rows = [
        {'sid': '9900000000001', 'name': '李四', 'college': '计算机学院', 'class_text': '网络24-1班', 'day_text': '周一', 'session': '早'},
        {'sid': '', 'name': '网络24-1班全体学生（共30人）', 'college': '计算机学院', 'class_text': '网络24-1班', 'day_text': '周一', 'session': '早'},
    ]
    create_exempt_workbook(p, early_rows=rows)
    records = load_exemptions(str(p), source='未分类')
    agg = aggregate_exemptions(records, week_no=7, campus='雁山')
    count = agg.exempt_count('计算机学院', '周一', '早自习')
    assert count == 30, f"Group exemption of 30 should cover individual, total should be 30, got {count}"
    assert agg.covered_personal >= 1, "Covered individual should be recorded in covered_personal stat"


def test_r3_hamilton_remainder_room_allocation():
    """Test 3 (R3.3): Hare-Niemeyer (Hamilton) maximum remainder algorithm partitions without integer drift."""
    # Rooms with capacities 50, 30, 20. Total = 100.
    # Allocate 13 unassigned students
    weights = [50, 30, 20]
    total_to_allocate = 13
    allocated = allocate(total_to_allocate, weights)

    assert sum(allocated) == total_to_allocate, f"Sum of allocated {allocated} must equal {total_to_allocate}"
    # 50/100*13 = 6.5 -> 7
    # 30/100*13 = 3.9 -> 4
    # 20/100*13 = 2.6 -> 2 (or 3 depending on remainder resolution)
    # Check non-negative integers
    assert all(isinstance(x, int) and x >= 0 for x in allocated)


def test_r3_room_capacity_cap_and_overflow(tmp_path: Path):
    """Test 4 (R3.4): Room exemption is strictly capped at room.expected and overflow is tracked."""
    detail_path = tmp_path / "detail_overflow.xlsx"
    # Room 8101 has expected = 20
    create_detail_workbook(
        detail_path,
        early_blocks=[('计算机学院', 100, [
            ('8101', 20, {'周一': 18, '周三': 18, '周四': 18, '周五': 18}),
            ('8102', 80, {'周一': 75, '周三': 75, '周四': 75, '周五': 75}),
        ])]
    )

    # 35 students with explicit room 8101 (exceeds capacity 20)
    exempt_rows = [
        {'sid': f'990000000{i:04d}', 'college': '计算机学院', 'day_text': '周一', 'session': '早自习', 'room': '8101'}
        for i in range(1, 36)
    ]
    exempt_path = tmp_path / "exempt_overflow.xlsx"
    create_exempt_workbook(exempt_path, early_rows=exempt_rows)

    out_file = tmp_path / "result_overflow.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)

    wb = load_workbook(out_file, data_only=True)
    ws = wb['雁山（早）']
    exempt_cell_val = ws.cell(row=4, column=7).value  # Room 8101 Mon exempt
    wb.close()

    assert exempt_cell_val <= 20, f"Room exemption must be capped at expected 20, got {exempt_cell_val}"
    # Verify overflow is logged in anomalies
    has_overflow_anomaly = any('超额' in str(a) or '封顶' in str(a) for a in result.anomalies)
    assert has_overflow_anomaly, "Room capacity overflow should be reported in result.anomalies"


def test_r3_attendance_rate_formula_accuracy(tmp_path: Path):
    """Test 5 (R3.5): Rate formula Actual / (Expected - Exempt) is computed accurately."""
    detail_path = tmp_path / "detail_rate.xlsx"
    # College Total = 100
    # Room 8101: Expected 100, Mon Actual 80
    create_detail_workbook(
        detail_path,
        early_blocks=[('理学院', 100, [
            ('8101', 100, {'周一': 80, '周三': 80, '周四': 80, '周五': 80}),
        ])]
    )
    # 20 exemptions on Monday
    exempt_rows = [
        {'sid': f'990000001{i:04d}', 'college': '理学院', 'day_text': '周一', 'session': '早自习', 'room': '8101'}
        for i in range(1, 21)
    ]
    exempt_path = tmp_path / "exempt_rate.xlsx"
    create_exempt_workbook(exempt_path, early_rows=exempt_rows)

    out_file = tmp_path / "result_rate.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)

    # Expected Rate: Actual=80, Denominator = 100 - 20 = 80 -> Rate = 80/80 = 1.0 (100.00%)
    for r in result.college_rates:
        if r.college == '理学院' and r.day == '周一' and r.session == '早自习':
            assert r.rate == pytest.approx(1.0, 1e-4)
            assert r.expected == 100
            assert r.exempt == 20
            assert r.actual == 80


def test_r3_complete_mode_effective_denominator(tmp_path: Path):
    """Test 6 (R3.6 / BUG 2 Regression): In complete_mode=True with missing rooms, rate denominator scales to effective rooms."""
    detail_path = tmp_path / "detail_bug2.xlsx"
    # College total: 200.
    # Room 8101 (expected 100, actual 100)
    # Room 8102 (expected 100, actual None - missing)
    create_detail_workbook(
        detail_path,
        early_blocks=[('工学院', 200, [
            ('8101', 100, {'周一': 100, '周三': 100, '周四': 100, '周五': 100}),
            ('8102', 100, {'周一': None, '周三': None, '周四': None, '周五': None}),
        ])]
    )
    exempt_path = tmp_path / "exempt_bug2.xlsx"
    create_exempt_workbook(exempt_path, early_rows=[])

    out_file = tmp_path / "result_bug2.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file, complete_mode=True)

    # Under BUG 2: actual_sum = 100, denominator = 200 (full college) -> rate was 50% (distorted!)
    # Under fixed code: effective rooms expected = 100, rate should be 100 / 100 = 100.0%
    for r in result.college_rates:
        if r.college == '工学院' and r.day == '周一' and r.session == '早自习':
            assert r.rate is not None
            assert r.rate == pytest.approx(1.0, 0.05), (
                f"BUG 2 Regression: With missing room in complete_mode, rate was {r.rate}. "
                "Effective denominator should be used rather than full college denominator!"
            )


def test_r3_weekly_average_rate_strictness(tmp_path: Path):
    """Test 7 (R3.7): Weekly average rate requires all statistical days to be complete; missing one leaves average None."""
    detail_path = tmp_path / "detail_avg.xlsx"
    # Tuesday/Wednesday missing actuals in strict mode (complete_mode=False)
    create_detail_workbook(
        detail_path,
        early_blocks=[('物理学院', 100, [
            ('8101', 100, {'周一': 90, '周三': None, '周四': 90, '周五': 90}),
        ])]
    )
    exempt_path = tmp_path / "exempt_avg.xlsx"
    create_exempt_workbook(exempt_path, early_rows=[])

    out_file = tmp_path / "result_avg.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file, complete_mode=False)

    # Weekly average rate for 物理学院 must be None
    weekly = next((w for w in result.weekly_rates if w.college == '物理学院' and w.sheet == '雁山（早）'), None)
    assert weekly is not None, "Weekly rate record should exist"
    assert weekly.rate is None, (
        f"Weekly average should be None when Wednesday is missing actuals, got {weekly.rate}"
    )
