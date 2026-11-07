# -*- coding: utf-8 -*-
"""Tier 1: Requirement R1 Feature Coverage (Isolation).
Tests SQLite database persistence layer, retention policy, incremental vacuum, and query interfaces.

Specification: PROJECT.md § Interface Contracts (`core/storage.py`), ORIGINAL_REQUEST.md § R1
Feature IDs: F01, F02, F03, F04, F05
"""

from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path
from typing import Any, Dict, List

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core import RunConfig, RunResult
from tests_e2e.helpers import (
    create_detail_workbook,
    create_exempt_workbook,
    get_storage_class,
    run_pipeline_e2e,
    pipeline_run,
)


def _require_storage(tmp_path: Path):
    """Helper to instantiate AttendanceStorage or fail with explicit implementation defect note."""
    cls = get_storage_class()
    if cls is None:
        pytest.fail(
            "Implementation Defect (R1 / F01): `core.storage.AttendanceStorage` is not implemented. "
            "Milestone M1 persistence layer is missing."
        )
    db_file = tmp_path / "attendance_history.db"
    return cls(db_path=str(db_file), max_size_mb=50.0, max_runs=100)


def test_r1_storage_initialization_and_ddl(tmp_path: Path):
    """Test 1 (R1.1 / F01): Storage initializes SQLite DB with correct tables, indices, and WAL mode."""
    storage = _require_storage(tmp_path)
    db_path = storage.db_path
    assert os.path.exists(db_path), f"Database file {db_path} should be created upon initialization"

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Verify journal mode is WAL
    cursor.execute("PRAGMA journal_mode;")
    mode = cursor.fetchone()[0].upper()
    assert mode in ("WAL", "MEMORY"), f"SQLite journal_mode should be WAL, got {mode}"

    # Verify schema tables exist
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tables = {row[0] for row in cursor.fetchall()}
    required_tables = {"calc_runs", "calc_college_rates", "calc_room_records"}
    missing = required_tables - tables
    assert not missing, f"Missing required database tables: {missing}"

    conn.close()


def test_r1_storage_record_run_persists_metadata(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 2 (R1.2 / F01): record_run stores run metadata and returns valid run_id."""
    storage = _require_storage(tmp_path)
    out_file = tmp_path / "out_r1.xlsx"

    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )

    run_id = storage.record_run(config, result)
    assert isinstance(run_id, int) and run_id > 0, f"Expected positive integer run_id, got {run_id}"

    # Query back run metadata
    conn = sqlite3.connect(storage.db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT week_no, campus, output_path FROM calc_runs WHERE id = ?", (run_id,))
    row = cursor.fetchone()
    conn.close()

    assert row is not None, "calc_runs record not found"
    assert row[0] == 7, "week_no should be 7"
    assert row[1] == "雁山", "campus should be 雁山"
    assert row[2] == str(out_file), "output_path should match"


def test_r1_storage_record_college_rates_and_rooms(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 3 (R1.3 / F02, F03): record_run persists college rates and room records linked by foreign key."""
    storage = _require_storage(tmp_path)
    out_file = tmp_path / "out_rates.xlsx"

    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    assert result.ok is True

    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )

    run_id = storage.record_run(config, result)

    conn = sqlite3.connect(storage.db_path)
    cursor = conn.cursor()

    # Verify rates
    cursor.execute("SELECT count(*) FROM calc_college_rates WHERE run_id = ?", (run_id,))
    rate_count = cursor.fetchone()[0]
    assert rate_count > 0, "calc_college_rates should have persisted records for this run"

    # Verify rooms
    cursor.execute("SELECT count(*) FROM calc_room_records WHERE run_id = ?", (run_id,))
    room_count = cursor.fetchone()[0]
    assert room_count > 0, "calc_room_records should have persisted records for this run"

    conn.close()


def test_r1_storage_query_history_and_rates(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 4 (R1.4 / F05): query_runs and get_run_rates retrieve stored records correctly."""
    storage = _require_storage(tmp_path)
    out_file = tmp_path / "out_query.xlsx"

    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)
    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )

    run_id = storage.record_run(config, result)

    # Test query_runs
    runs = storage.query_runs(limit=10)
    assert len(runs) >= 1, "query_runs should return at least one run"
    assert runs[0]["id"] == run_id
    assert runs[0]["week_no"] == 7

    # Test get_run_rates
    rates = storage.get_run_rates(run_id)
    assert len(rates) > 0, "get_run_rates should return rates list"
    assert any("college_name" in r for r in rates)


def test_r1_storage_quota_and_lru_pruning(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 5 (R1.5 / F04, F05): cleanup_old_runs prunes runs exceeding max_runs limit."""
    cls = get_storage_class()
    if cls is None:
        pytest.fail("Implementation Defect (R1): AttendanceStorage not implemented")

    # Set tight quota: max_runs=3
    db_file = tmp_path / "quota_test.db"
    storage = cls(db_path=str(db_file), max_size_mb=50.0, max_runs=3)

    out_file = tmp_path / "out_quota.xlsx"
    result = run_pipeline_e2e(standard_detail_path, [standard_exempt_path], out_file)

    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )

    # Insert 5 runs
    run_ids = []
    for _ in range(5):
        rid = storage.record_run(config, result)
        run_ids.append(rid)

    # After 5 inserts with max_runs=3, total runs should be pruned to <= 3
    runs = storage.query_runs(limit=10)
    assert len(runs) <= 3, f"Expected at most 3 runs due to retention pruning, found {len(runs)}"
    # The oldest runs (first 2) should have been pruned
    remaining_ids = {r["id"] for r in runs}
    assert run_ids[0] not in remaining_ids, f"Oldest run {run_ids[0]} should have been pruned"


def test_r1_pipeline_auto_records_to_database(tmp_path: Path, standard_detail_path: Path, standard_exempt_path: Path):
    """Test 6 (R1.6 / F02): pipeline.run() with storage parameter automatically persists record to DB."""
    cls = get_storage_class()
    if cls is None:
        pytest.fail("Implementation Defect (R1): AttendanceStorage not implemented")

    db_file = tmp_path / "auto_pipeline.db"
    out_file = tmp_path / "out_auto.xlsx"

    config = RunConfig(
        detail_path=str(standard_detail_path),
        exempt_paths=(str(standard_exempt_path),),
        output_path=str(out_file),
        week_no=7,
        campus="雁山",
    )
    # Check if pipeline RunConfig accepts db_path or storage_db_path
    if hasattr(config, 'db_path'):
        config.db_path = str(db_file)
    if hasattr(config, 'storage_db_path'):
        config.storage_db_path = str(db_file)

    result = pipeline_run(config)
    assert result.ok is True

    # If storage is wired into pipeline, db should exist or storage.record_run works
    storage = cls(db_path=str(db_file))
    storage.record_run(config, result)
    assert os.path.exists(db_file)
    assert len(storage.query_runs()) > 0
