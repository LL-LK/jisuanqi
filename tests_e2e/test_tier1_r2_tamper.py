# -*- coding: utf-8 -*-
"""Tier 1: Requirement R2 Feature Coverage (Isolation).
Tests dual-sheet backfill, single-exempt identity verification, sheet protection,
cell-level protection, hidden metadata, and cryptographic tamper detection.

Specification: PROJECT.md § Interface Contracts (`core/export.py`), ORIGINAL_REQUEST.md § R2
Feature IDs: F06, F07, F08, F09, F10
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

from core import RunConfig
from core.pipeline import run as pipeline_run
from tests_e2e.helpers import (
    create_detail_workbook,
    create_exempt_workbook,
    get_sheet_protection_applier,
    get_tamper_verifier,
    run_pipeline_e2e,
)


def test_r2_dual_sheet_complete_backfill(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 1 (R2.1): Importing 1 detail + 1 exemption sheet completely backfills exemptions and computes rates."""
    out_file = tmp_path / "result_r2_backfill.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)

    assert result.ok is True, f"Pipeline execution failed: {result.error}"
    assert os.path.exists(out_file), "Exported Excel file must exist"

    wb = load_workbook(out_file, data_only=True)
    ws_early = wb['雁山（早）']

    # Row 4 is 计算机学院 8101 (expected 50)
    # Exemption for student 9900000000001 (room 8101) on Monday
    mon_exempt_cell = ws_early.cell(row=4, column=7)  # G4 is Mon exempt
    assert mon_exempt_cell.value is not None, "Monday exemption should be filled"
    assert isinstance(mon_exempt_cell.value, int)
    assert mon_exempt_cell.value >= 1, "Expected at least 1 exemption for 8101 on Monday"

    # Verify rate column (N4: early Monday rate)
    rate_cell = ws_early.cell(row=4, column=14)
    assert rate_cell.value is not None, "Monday attendance rate should be filled"
    assert isinstance(rate_cell.value, float)
    assert 0.0 < rate_cell.value <= 1.0, f"Rate {rate_cell.value} out of expected range (0, 1]"

    wb.close()


def test_r2_single_exempt_identity_check_enabled(tmp_path: Path):
    """Test 2 (R2.2 / BUG 1 Regression): Single exemption sheet activates identity check & conflict review."""
    # Construct a student conflict in a single exemption file (same student in two different colleges)
    detail_path = tmp_path / "detail_single_check.xlsx"
    create_detail_workbook(detail_path)

    exempt_path = tmp_path / "exempt_single_conflict.xlsx"
    conflict_rows = [
        {'sid': '9900000000999', 'college': '计算机学院', 'name': '学生冲突', 'day_text': '周一', 'session': '早自习', 'room': '8101'},
        {'sid': '9900000000999', 'college': '理学院', 'name': '学生冲突', 'day_text': '周一', 'session': '早自习', 'room': '8201'},
    ]
    create_exempt_workbook(exempt_path, early_rows=conflict_rows)

    out_file = tmp_path / "result_single_check.xlsx"
    result = run_pipeline_e2e(detail_path, [exempt_path], out_file)

    # BUG 1 check: Under single exemption import, identity_check should be True, detecting the cross-college conflict
    # In buggy version, identity_check was False, bypassing conflict detection
    anomalies_text = " ".join(str(a) for a in result.anomalies)
    has_conflict_flag = (
        '跨学院' in anomalies_text
        or '冲突' in anomalies_text
        or result.exempt_stats.get('跨学院重复申请事件', 0) > 0
        or result.exempt_stats.get('同学号身份冲突', 0) > 0
    )
    assert has_conflict_flag, (
        "GAP-R2-02 / BUG 1: Single exemption sheet failed to trigger identity verification or report cross-college conflict! "
        "`core/pipeline.py:126` identity_check must be activated for single exempt file."
    )


def test_r2_worksheet_password_protection(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 3 (R2.3 / F06): Exported workbook has sheet protection enabled with password."""
    out_file = tmp_path / "result_protected.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    for sheet_name in ('雁山（早）', '雁山（晚）'):
        ws = wb[sheet_name]
        assert ws.protection.sheet is True, (
            f"GAP-R2-01: Sheet '{sheet_name}' does not have sheet protection enabled (ws.protection.sheet is False). "
            "Milestone M2 tamper protection required."
        )
        assert ws.protection.password is not None or ws.protection.hashValue is not None, (
            f"Sheet '{sheet_name}' must have a protection password set."
        )
    wb.close()


def test_r2_critical_cells_locked(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 4 (R2.4 / F07): Critical data cells (actuals, expected, rates, exemptions) are locked."""
    out_file = tmp_path / "result_locked.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    ws = wb['雁山（早）']

    # Row 4: 8101
    # Col 5 (E): Expected headcount
    # Col 6 (F): Monday actual
    # Col 7 (G): Monday exempt
    # Col 14 (N): Monday rate
    for col_idx, col_name in [(5, "应到人数"), (6, "实到人数"), (7, "免勤人数"), (14, "出勤率")]:
        cell = ws.cell(row=4, column=col_idx)
        assert cell.protection is not None and cell.protection.locked is True, (
            f"GAP-R2-01: Critical cell at row 4, col {col_idx} ({col_name}) must be locked (cell.protection.locked=True)."
        )

    wb.close()


def test_r2_hidden_tamper_metadata_sheet(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 5 (R2.5 / F08): Exported workbook contains veryHidden _系统校验元数据 sheet with HMAC signatures."""
    out_file = tmp_path / "result_metadata.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    wb = load_workbook(out_file)
    meta_sheet_name = '_系统校验元数据'
    assert meta_sheet_name in wb.sheetnames, (
        f"GAP-R2-01: Hidden metadata sheet '{meta_sheet_name}' is missing from exported workbook."
    )
    ws_meta = wb[meta_sheet_name]
    assert ws_meta.sheet_state in ('hidden', 'veryHidden'), (
        f"Metadata sheet must be hidden/veryHidden, got {ws_meta.sheet_state}"
    )
    # Check that it contains signature entries
    assert ws_meta.max_row >= 2, "Metadata sheet should contain row HMAC entries"
    wb.close()


def test_r2_tamper_verifier_detects_tampered_actuals(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 6 (R2.6 / F09): verify_workbook_integrity accurately detects when actual headcount is modified."""
    verifier = get_tamper_verifier()
    if verifier is None:
        pytest.fail(
            "Implementation Defect (R2 / F09): `verify_workbook_integrity` function is not implemented. "
            "Milestone M2 tamper verifier is missing."
        )

    out_file = tmp_path / "result_for_tamper.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    # 1. Untampered workbook passes verification
    is_valid, violations = verifier(str(out_file))
    assert is_valid is True, f"Untampered workbook failed verification: {violations}"

    # 2. Deliberately tamper with an actual headcount (change from 48 to 99 in row 4, col 6)
    tampered_file = tmp_path / "result_tampered.xlsx"
    wb = load_workbook(out_file)
    ws = wb['雁山（早）']
    original_val = ws.cell(row=4, column=6).value
    ws.cell(row=4, column=6, value=99)  # Illegally modified!
    wb.save(tampered_file)
    wb.close()

    # 3. Verifier must catch the tamper
    is_valid_tampered, violations = verifier(str(tampered_file))
    assert is_valid_tampered is False, "Verifier failed to detect illegal headcount modification!"
    assert len(violations) > 0, "Verifier should report violation details"
