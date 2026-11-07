# -*- coding: utf-8 -*-
"""Tier 2: Boundary & Corner Cases.
Tests extreme inputs, empty tables, division-by-zero, corrupted files,
quota saturation (100+ runs), missing weekdays, and invalid schemas.

Specification: PROJECT.md § Architecture, ORIGINAL_REQUEST.md § Acceptance Criteria
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import RunConfig
from core.excel_safety import verify_excel_health
from core.io_exempt import aggregate_exemptions, normalize_sid
from core.pipeline import run as pipeline_run
from tests_e2e.helpers import (
    create_detail_workbook,
    create_exempt_workbook,
    get_storage_class,
    run_pipeline_e2e,
)


def test_tier2_empty_detail_and_exempt_tables(tmp_path: Path):
    """Test 2.1: Empty detail table (headers only) and empty exemption table raise clean ValueError, not unhandled crash."""
    empty_detail = tmp_path / "empty_detail.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = '雁山（早）'
    ws.append(['序号', '学院（人数）', '教室', '专业班级', '应到人数', '周一实到', '周一免勤人数'])
    wb.save(empty_detail)
    wb.close()

    empty_exempt = tmp_path / "empty_exempt.xlsx"
    create_exempt_workbook(empty_exempt, early_rows=[])

    out_file = tmp_path / "out_empty.xlsx"

    # Running pipeline with zero classrooms should be safely caught in result.error
    result = run_pipeline_e2e(empty_detail, [empty_exempt], out_file)
    assert result.ok is False, "Empty detail sheet should cause pipeline to return ok=False"
    assert "未找到" in result.error or "没有可识别" in result.error or "ValueError" in result.error


def test_tier2_zero_expected_headcounts_and_division_by_zero(tmp_path: Path):
    """Test 2.2: Expected - Exempt <= 0 handles division by zero safely (None in strict mode, 1.0 in complete mode)."""
    detail_path = tmp_path / "detail_zero_denom.xlsx"
    # College with 10 people total, 10 actual
    create_detail_workbook(
        detail_path,
        early_blocks=[('数学学院', 10, [
            ('8101', 10, {'周一': 10, '周三': 10, '周四': 10, '周五': 10}),
        ])]
    )
    # 10 exemptions on Monday -> Denominator = 10 - 10 = 0!
    exempt_rows = [
        {'sid': f'990000000{i:04d}', 'college': '数学学院', 'day_text': '周一', 'session': '早自习', 'room': '8101'}
        for i in range(1, 11)
    ]
    exempt_path = tmp_path / "exempt_zero_denom.xlsx"
    create_exempt_workbook(exempt_path, early_rows=exempt_rows)

    # 1. Strict mode: Rate should be None, status indicates non-positive denominator
    out_strict = tmp_path / "out_strict.xlsx"
    result_strict = run_pipeline_e2e(detail_path, [exempt_path], out_strict, complete_mode=False)
    assert result_strict.ok is True
    rate_strict = next(r for r in result_strict.college_rates if r.college == '数学学院' and r.day == '周一')
    assert rate_strict.rate is None, "Strict mode must output None when denominator <= 0"
    assert "分母非正" in rate_strict.status or "无法计算" in rate_strict.status

    # 2. Complete mode: Rate should be capped at 1.0 (100% full attendance fallback)
    out_complete = tmp_path / "out_complete.xlsx"
    result_complete = run_pipeline_e2e(detail_path, [exempt_path], out_complete, complete_mode=True)
    assert result_complete.ok is True
    rate_complete = next(r for r in result_complete.college_rates if r.college == '数学学院' and r.day == '周一')
    assert rate_complete.rate == 1.0, f"Complete mode should fall back to 1.0 (100%), got {rate_complete.rate}"


def test_tier2_corrupted_excel_and_magic_number(tmp_path: Path):
    """Test 2.3: Corrupted Excel files (invalid magic bytes, truncated zip) are intercepted by safety checks."""
    corrupted_file = tmp_path / "corrupted.xlsx"
    # Write garbage non-zip bytes
    corrupted_file.write_bytes(b"INVALID_MAGIC_HEADER_GARBAGE_BYTES_1234567890")

    check_res = verify_excel_health(str(corrupted_file))
    assert check_res.is_valid is False, "verify_excel_health should reject corrupted file"
    assert check_res.error_code in ('CORRUPTED', 'NOT_OOXML', 'BAD_ZIP'), (
        f"Expected corruption error code, got {check_res.error_code}"
    )


def test_tier2_extreme_exemption_headcount_exceeding_total(tmp_path: Path):
    """Test 2.4: Extreme exemptions (e.g. 500 exemptions for 50 people) handled without integer overflow or crash."""
    detail_path = tmp_path / "detail_extreme.xlsx"
    create_detail_workbook(
        detail_path,
        early_blocks=[('微电子学院', 50, [
            ('8101', 50, {'周一': 50, '周三': 50, '周四': 50, '周五': 50}),
        ])]
    )
    # 500 exemptions on Monday
    exempt_rows = [
        {'sid': f'99000000{i:05d}', 'college': '微电子学院', 'day_text': '周一', 'session': '早自习', 'room': '8101'}
        for i in range(1, 501)
    ]
    exempt_path = tmp_path / "exempt_extreme.xlsx"
    create_exempt_workbook(exempt_path, early_rows=exempt_rows)

    out_file = tmp_path / "out_extreme.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)
    assert result.ok is True

    rate_rec = next(r for r in result.college_rates if r.college == '微电子学院' and r.day == '周一')
    # Denominator = 50 - 500 = -450 <= 0 -> rate is None
    assert rate_rec.rate is None


def test_tier2_database_quota_saturation_100_runs(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 2.5: High volume simulation (105 runs) confirms database prunes old runs and stays within 50MB quota."""
    cls = get_storage_class()
    if cls is None:
        pytest.fail("Implementation Defect (R1): AttendanceStorage not implemented")

    db_path = tmp_path / "saturation_test.db"
    storage = cls(db_path=str(db_path), max_size_mb=50.0, max_runs=100)

    out_file = tmp_path / "out_sat.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )

    # Insert 105 runs
    for i in range(105):
        storage.record_run(config, result)

    runs = storage.query_runs(limit=200)
    assert len(runs) <= 100, f"Total persisted runs {len(runs)} must be capped at max_runs=100"

    file_size_mb = os.path.getsize(db_path) / (1024 * 1024)
    assert file_size_mb < 50.0, f"Database file size {file_size_mb:.2f}MB must be strictly under 50MB"


def test_tier2_missing_weekdays_and_partial_attendance(tmp_path: Path):
    """Test 2.6: Detail sheet missing Friday or Tuesday columns computes available days without indexing errors."""
    detail_path = tmp_path / "detail_missing_weekday.xlsx"
    # Detail with only Monday and Wednesday
    wb = Workbook()
    ws = wb.active
    ws.title = '雁山（早）'
    ws.append(['序号', '学院（人数）', '教室', '专业班级', '应到人数', '周一实到', '周一免勤人数', '周三实到', '周三免勤人数', '学生出勤率'])
    ws.append(['', '', '', '', '', '', '', '', '', '周一', '周三', '平均'])
    ws.append([1, '软件学院（50人）', '8101', '软24-1', 50, 48, None, 49, None])
    wb.save(detail_path)
    wb.close()

    exempt_path = tmp_path / "exempt_dummy.xlsx"
    create_exempt_workbook(exempt_path, early_rows=[])

    out_file = tmp_path / "out_missing_weekday.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)
    assert result.ok is True


def test_tier2_invalid_student_id_formats():
    """Test 2.7: Student IDs with short integers, float suffixes, spaces, or letters are cleaned safely."""
    # Integer less than 13 digits should be left-padded with zeros
    sid_int, ok, notes_int = normalize_sid(123456789)
    assert sid_int == '0000123456789', f"Expected zero-padded 13 digits, got {sid_int}"
    assert ok is True

    # Float format from Excel (e.g. 9900000000001.0)
    sid_flt, ok_flt, _ = normalize_sid(9900000000001.0)
    assert sid_flt == '9900000000001', f"Float suffix .0 should be stripped, got {sid_flt}"
    assert ok_flt is True

    # String with spaces
    sid_sp, _, _ = normalize_sid('  9900000000002  ')
    assert sid_sp == '9900000000002'


def test_tier2_placeholder_group_exemptions_without_count(tmp_path: Path):
    """Test 2.8: '全班' without headcount held as incomplete without hallucinating numbers."""
    from core.io_exempt import load_exemptions
    p = tmp_path / "exempt_placeholder.xlsx"
    rows = [{'sid': '全班', 'name': '全体同学', 'college': '艺术学院', 'class_text': '音乐24-1班', 'day_text': '周一', 'session': '早'}]
    create_exempt_workbook(p, early_rows=rows)
    records = load_exemptions(str(p), source='未分类')
    agg = aggregate_exemptions(records, week_no=7, campus='雁山')
    assert agg.incomplete_records >= 1, "Record without headcount should be counted as incomplete"
    count = agg.exempt_count('艺术学院', '周一', '早自习')
    assert count == 0, f"Incomplete group record should not contribute fake headcount, got {count}"
