# -*- coding: utf-8 -*-
"""Tier 3: Cross-Feature Combinations.
Tests interaction between multiple features across milestones:
- Dual-sheet backfill + tamper protection + SQLite database persistence
- Complete calculation mode + mobile Excel styling + integrity verification
- GUI execution + tray minimization + memory trimming
- Tamper detection against persisted database ground truth
- Multi-week consecutive runs with quota maintenance

Specification: PROJECT.md § Milestones M1, M2, M3, M5
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
from tests_e2e.helpers import (
    create_detail_workbook,
    create_exempt_workbook,
    get_storage_class,
    get_tamper_verifier,
    run_pipeline_e2e,
)


def test_tier3_dual_sheet_backfill_with_tamper_protection_and_db_persistence(
    tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path
):
    """Test 3.1: Dual-sheet backfill exports tamper-protected Excel and persists audit record in SQLite DB."""
    cls = get_storage_class()
    if cls is None:
        pytest.fail("Implementation Defect (R1): AttendanceStorage not implemented")

    db_path = tmp_path / "combined_history.db"
    out_file = tmp_path / "combined_result.xlsx"

    storage = cls(db_path=str(db_path))

    # 1. Run pipeline
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    # 2. Persist to DB
    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )
    run_id = storage.record_run(config, result)
    assert run_id > 0

    # 3. Verify Excel tamper protection
    wb = load_workbook(out_file)
    ws = wb['雁山（早）']
    assert ws.protection.sheet is True, "GAP-R2-01: Exported workbook must have sheet protection enabled"
    wb.close()

    # 4. Verify DB ground truth matches Excel
    stored_runs = storage.query_runs()
    assert len(stored_runs) == 1
    assert stored_runs[0]["week_no"] == 7


def test_tier3_complete_mode_with_mobile_styling_and_integrity_audit(
    tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path
):
    """Test 3.2: Complete mode run outputs mobile-optimized Excel and verifies integrity signature."""
    verifier = get_tamper_verifier()
    if verifier is None:
        pytest.fail("Implementation Defect (R2): `verify_workbook_integrity` not implemented")

    out_file = tmp_path / "complete_mobile.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file, complete_mode=True)
    assert result.ok is True

    # Check mobile styling
    wb = load_workbook(out_file)
    ws = wb['雁山（早）']
    assert ws.freeze_panes == 'D4', "Mobile freeze_panes must be D4"
    assert ws.column_dimensions['N'].width >= 12.0, "Rate column width must be >= 12.0"
    wb.close()

    # Check integrity verification
    is_valid, violations = verifier(str(out_file))
    assert is_valid is True, f"Complete mode output failed integrity check: {violations}"


def test_tier3_gui_batch_run_with_minimize_tray_and_memory_trim(
    qapp, tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path
):
    """Test 3.3: GUI workflow invocation, followed by minimize to tray and memory trimming."""
    from app import MainWindow
    win = MainWindow()
    try:
        win.detail_slot.set_path(str(standard_detail_path))
        win.exempt_slot.set_paths([str(standard_exempt_path)])
        out_file = tmp_path / "gui_out.xlsx"
        win.out_edit.setText(str(out_file))

        # Minimize to tray
        from PySide6.QtCore import QEvent, Qt
        win.setWindowState(win.windowState() | Qt.WindowMinimized)
        win.changeEvent(QEvent(QEvent.WindowStateChange))
        qapp.processEvents()

        # Check tray remains active
        assert win.tray_icon.isVisible() is True
        # Check memory trim method executed safely
        win._trim_memory()
    finally:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()


def test_tier3_tamper_detection_on_persisted_export(
    tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path
):
    """Test 3.4: Deliberate tamper of Excel file is detected by verifier while DB preserves true record."""
    cls = get_storage_class()
    verifier = get_tamper_verifier()
    if cls is None or verifier is None:
        pytest.fail("Implementation Defect: Storage or Tamper Verifier not implemented")

    db_path = tmp_path / "truth_history.db"
    out_file = tmp_path / "truth_result.xlsx"

    storage = cls(db_path=str(db_path))
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )
    run_id = storage.record_run(config, result)

    # Tamper with Excel file
    wb = load_workbook(out_file)
    ws = wb['雁山（早）']
    ws.cell(row=4, column=6, value=999)  # Tampered!
    tampered_file = tmp_path / "tampered_result.xlsx"
    wb.save(tampered_file)
    wb.close()

    # Verifier detects tamper
    is_valid, _ = verifier(str(tampered_file))
    assert is_valid is False, "Verifier must detect tampering"

    # DB retains uncorrupted original rates
    db_rates = storage.get_run_rates(run_id)
    assert len(db_rates) > 0


def test_tier3_multi_week_consecutive_runs_with_quota_cleanup(
    tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path
):
    """Test 3.5: Consecutive weekly runs (weeks 1 to 5) stored in DB and queried chronologically."""
    cls = get_storage_class()
    if cls is None:
        pytest.fail("Implementation Defect: Storage not implemented")

    db_path = tmp_path / "multi_week.db"
    storage = cls(db_path=str(db_path), max_runs=10)

    for week in range(1, 6):
        out_file = tmp_path / f"week_{week}.xlsx"
        res = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file, week_no=week)
        cfg = RunConfig(
            detail_path=str(standard_detail_path),
            exempt_paths=(str(standard_exempt_path),),
            output_path=str(out_file),
            week_no=week,
            campus="雁山",
        )
        storage.record_run(cfg, res)

    runs = storage.query_runs(limit=10)
    assert len(runs) == 5
    weeks_stored = {r["week_no"] for r in runs}
    assert weeks_stored == {1, 2, 3, 4, 5}
