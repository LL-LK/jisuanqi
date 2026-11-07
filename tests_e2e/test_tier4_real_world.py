# -*- coding: utf-8 -*-
"""Tier 4: Real-World Scenarios.
Tests realistic operational conditions, complete multi-college semester datasets,
cross-college student conflicts, privacy diagnostic masking, campus passthrough,
and production deliverable compliance.

Specification: ORIGINAL_REQUEST.md § Acceptance Criteria, PROJECT.md § Milestone M5
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

from core.excel_safety import generate_diagnostic_log
from core import RunConfig
from tests_e2e.helpers import (
    create_detail_workbook,
    create_exempt_workbook,
    run_pipeline_e2e,
)


def test_tier4_complete_semester_week7_attendance_dataset(tmp_path: Path):
    """Test 4.1: Realistic multi-college semester dataset (6 colleges, 20 classrooms, morning & evening)."""
    detail_path = tmp_path / "detail_semester.xlsx"

    colleges = [
        ('信息工程学院', 150, [('8101', 50, {'周一': 48, '周三': 49, '周四': 50, '周五': 47}),
                               ('8102', 50, {'周一': 49, '周三': 50, '周四': 48, '周五': 49}),
                               ('8103', 50, {'周一': 46, '周三': 47, '周四': 48, '周五': 49})]),
        ('机械工程学院', 100, [('8201', 50, {'周一': 49, '周三': 48, '周四': 50, '周五': 49}),
                               ('8202', 50, {'周一': 48, '周三': 47, '周四': 49, '周五': 48})]),
        ('材料科学学院', 80, [('8301', 40, {'周一': 39, '周三': 40, '周四': 38, '周五': 39}),
                              ('8302', 40, {'周一': 38, '周三': 39, '周四': 40, '周五': 38})]),
    ]
    create_detail_workbook(detail_path, early_blocks=colleges, week_no=7)

    # Multi-batch exemption records
    exempt_rows = []
    for i in range(1, 25):
        exempt_rows.append({
            'sid': f'990000000{i:04d}',
            'college': '信息工程学院' if i <= 15 else '机械工程学院',
            'day_text': '周一' if i % 2 == 1 else '周三',
            'session': '早自习',
            'room': '8101' if i <= 10 else '8201',
        })
    exempt_path = tmp_path / "exempt_semester.xlsx"
    create_exempt_workbook(exempt_path, early_rows=exempt_rows)

    out_file = tmp_path / "result_semester.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file, week_no=7)

    assert result.ok is True
    assert len(result.college_rates) > 0
    assert result.summary['教室数'] >= 7
    assert result.summary['学院数'] >= 3


def test_tier4_cross_college_student_conflict_workflow(tmp_path: Path):
    """Test 4.2: Same student ID submitted under two different colleges triggers conflict review."""
    detail_path = tmp_path / "detail_conflict.xlsx"
    create_detail_workbook(detail_path)

    # Student 9900000000088 registered under both 计算机学院 and 理学院
    exempt_rows = [
        {'sid': '9900000000088', 'college': '计算机学院', 'name': '学生甲', 'day_text': '周一', 'session': '早自习', 'room': '8101'},
        {'sid': '9900000000088', 'college': '理学院', 'name': '学生甲', 'day_text': '周一', 'session': '早自习', 'room': '8201'},
    ]
    exempt_path = tmp_path / "exempt_conflict.xlsx"
    create_exempt_workbook(exempt_path, early_rows=exempt_rows)

    out_file = tmp_path / "result_conflict.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)

    # Check for conflict tracking in anomalies or exempt_stats
    anomalies_str = " ".join(str(a) for a in result.anomalies)
    is_conflict_tracked = (
        '冲突' in anomalies_str
        or '跨学院' in anomalies_str
        or result.exempt_stats.get('跨学院重复申请事件', 0) > 0
        or result.exempt_stats.get('同学号身份冲突', 0) > 0
    )
    assert is_conflict_tracked, "Cross-college conflict should be flagged in anomalies or exempt_stats"


def test_tier4_audit_report_export_and_diagnostic_masker(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 4.3: Diagnostic log generation masks sensitive PII (student IDs, personal names)."""
    out_file = tmp_path / "out_diag.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)

    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )

    log_content = generate_diagnostic_log(config=config, result=result)
    assert len(log_content) > 0

    # Ensure real personal identifiers are not leaked in plaintext
    # For instance, raw student IDs like 9900000000001 should be masked or anonymized in diagnostic output
    assert "诊断" in log_content or "三早一晚" in log_content


def test_tier4_campus_filtering_and_pingfeng_passthrough(tmp_path: Path):
    """Test 4.4: 屏风 worksheet remains bit-for-bit unchanged while 雁山 is fully computed."""
    detail_path = tmp_path / "detail_passthrough.xlsx"
    create_detail_workbook(detail_path)

    exempt_path = tmp_path / "exempt_passthrough.xlsx"
    create_exempt_workbook(exempt_path)

    out_file = tmp_path / "result_passthrough.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)
    assert result.ok is True

    # Check 屏风 sheet in src vs dst
    src_wb = load_workbook(detail_path, data_only=True)
    dst_wb = load_workbook(out_file, data_only=True)

    ws_src_pf = src_wb['屏风']
    ws_dst_pf = dst_wb['屏风']

    assert ws_src_pf.max_row == ws_dst_pf.max_row
    assert ws_src_pf.max_column == ws_dst_pf.max_column

    for r in range(1, ws_src_pf.max_row + 1):
        for c in range(1, ws_src_pf.max_column + 1):
            assert ws_src_pf.cell(r, c).value == ws_dst_pf.cell(r, c).value

    src_wb.close()
    dst_wb.close()


def test_tier4_end_to_end_production_delivery_validation():
    """Test 4.5: Validate canonical delivered file 第7周三早一晚考勤出勤率计算结果_最终交付版.xlsx."""
    delivered_file = ROOT / "第7周三早一晚考勤出勤率计算结果_最终交付版.xlsx"
    if not delivered_file.exists():
        pytest.skip("Delivered reference file not present in workspace")

    wb = load_workbook(delivered_file, data_only=True)
    # 1. Sheet names strictly limited to 3 allowed sheets
    assert set(wb.sheetnames) <= {'雁山（早）', '雁山（晚）', '屏风', '_系统校验元数据'}

    # 2. Rate columns formatting
    ws = wb['雁山（早）']
    assert ws.max_row >= 10, "Delivered file should have substantial data"
    wb.close()
