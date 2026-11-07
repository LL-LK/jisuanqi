# -*- coding: utf-8 -*-
"""SQLite 本地存储层功能、配额淘汰与管道集成专项测试套件。

覆盖特性:
  * F01: SQLite 数据库持久化层与 DDL 索引规范
  * F02: core.pipeline.run() 计算结果自动入库与防御性隔离
  * F03: 50MB / 100 批次 FIFO / 180 天 TTL 磁盘配额淘汰
  * F04: PRAGMA incremental_vacuum 空间物理整理与完整性检查
  * F05: query_runs / get_run_rates / get_run_rooms 多维历史检索接口
"""

from __future__ import annotations

import os
import shutil
import sqlite3
import threading
from datetime import datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

import pytest

from _helpers import build_focused_sample
from core.model import CollegeRate, CollegeWeeklyRate
from core.pipeline import RunConfig, RunResult, run
from core.storage import AttendanceStorage, DiskQuotaExceededError, StorageError


@pytest.fixture
def storage_db_path(tmp_path) -> Path:
    """提供严格隔离的临时 SQLite 数据库路径。"""
    return tmp_path / "test_attendance_history.db"


@pytest.fixture
def storage(storage_db_path) -> AttendanceStorage:
    """初始化标准测试存储实例 (默认配额: 100 批次, 50MB)。"""
    return AttendanceStorage(
        db_path=str(storage_db_path),
        max_size_mb=50.0,
        max_runs=100,
        retention_days=180,
    )


def _make_dummy_result(week_no: int = 7, campus: str = "雁山", ok: bool = True) -> Tuple[RunConfig, RunResult]:
    """快速构造用于入库测试的虚拟领域对象。"""
    config = RunConfig(
        detail_path="detail_mock.xlsx",
        exempt_path="exempt_mock.xlsx",
        output_path="result_mock.xlsx",
        week_no=week_no,
        campus=campus,
    )
    result = RunResult(
        ok=ok,
        output_path="result_mock.xlsx",
        summary={"教室数": 2, "学院数": 1, "写出出勤率单元格": 8},
        exempt_stats={"全校总免勤人次": 10, "同学号身份冲突": 0},
    )
    rate = CollegeRate(
        sheet="雁山（早）", session="早自习", college="理学院", day="周一",
        expected=100, room_expected_sum=100, actual=90, exempt=5,
        rate=0.9474, status="完整", note="测试记录", anomalies=[]
    )
    result.college_rates = [rate]
    result.weekly_rates = [
        CollegeWeeklyRate(sheet="雁山（早）", session="早自习", college="理学院",
                          calculated_days=4, scheduled_days=4, rate=0.95, needs_review=False)
    ]
    result.room_rows = [
        {"工作表": "雁山（早）", "行号": 4, "学院": "理学院", "教室": "8101",
         "分配专业班级": "统计26-1", "应到人数": 50, "周一实到": 45, "周一免勤": 2, "周一免勤来源": "精确"}
    ]
    return config, result


# ==============================================================================
# Test Class 1: TestStorageSchemaAndCRUD
# ==============================================================================
class TestStorageSchemaAndCRUD:
    def test_db_initialization_pragmas(self, storage, storage_db_path):
        with storage._get_connection() as conn:
            assert conn.execute("PRAGMA foreign_keys").fetchone()[0] == 1
            assert conn.execute("PRAGMA journal_mode").fetchone()[0].lower() == "wal"
            assert conn.execute("PRAGMA auto_vacuum").fetchone()[0] == 2

    def test_schema_tables_and_indexes_exist(self, storage, storage_db_path):
        conn = sqlite3.connect(str(storage_db_path))
        try:
            tables = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'")}
            expected = {"calc_runs", "calc_college_rates", "calc_room_records", "storage_maintenance_meta"}
            assert expected.issubset(tables)

            indexes = {r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='index'")}
            assert "idx_calc_runs_created_at" in indexes
            assert "idx_college_rates_run" in indexes
            assert "idx_room_records_run" in indexes
        finally:
            conn.close()

    def test_record_run_persists_and_returns_id(self, storage):
        cfg, res = _make_dummy_result(week_no=7, campus="雁山")
        run_id = storage.record_run(cfg, res, duration_ms=120)
        assert isinstance(run_id, int)
        assert run_id >= 1

        run_data = storage.get_run(run_id)
        assert run_data is not None
        assert run_data["week_no"] == 7
        assert run_data["campus"] == "雁山"
        assert run_data["status"] == "SUCCESS"
        assert run_data["duration_ms"] == 120

    def test_foreign_key_cascade_deletion(self, storage, storage_db_path):
        cfg, res = _make_dummy_result()
        run_id = storage.record_run(cfg, res)

        conn = sqlite3.connect(str(storage_db_path))
        conn.execute("PRAGMA foreign_keys = ON;")
        try:
            assert conn.execute("SELECT count(*) FROM calc_college_rates WHERE run_id=?", (run_id,)).fetchone()[0] > 0
            assert conn.execute("SELECT count(*) FROM calc_room_records WHERE run_id=?", (run_id,)).fetchone()[0] > 0

            # 显式删除主表批次
            conn.execute("DELETE FROM calc_runs WHERE id=?", (run_id,))
            conn.commit()

            # 校验级联清零
            assert conn.execute("SELECT count(*) FROM calc_college_rates WHERE run_id=?", (run_id,)).fetchone()[0] == 0
            assert conn.execute("SELECT count(*) FROM calc_room_records WHERE run_id=?", (run_id,)).fetchone()[0] == 0
        finally:
            conn.close()


# ==============================================================================
# Test Class 2: TestPipelineRunIntegration
# ==============================================================================
class TestPipelineRunIntegration:
    def test_pipeline_run_auto_inserts_storage(self, tmp_path):
        detail_path, exempt_path = build_focused_sample(tmp_path)
        out_excel = tmp_path / "result_out.xlsx"
        db_path = tmp_path / "pipeline_storage.db"

        config = RunConfig(
            detail_path=str(detail_path),
            exempt_path=str(exempt_path),
            output_path=str(out_excel),
            week_no=7,
            campus="雁山",
            db_path=str(db_path),
            enable_storage=True,
        )
        result = run(config)
        assert result.ok is True
        assert result.storage_run_id is not None
        assert result.storage_run_id >= 1
        assert result.summary.get("run_id") == result.storage_run_id

        storage = AttendanceStorage(db_path=str(db_path))
        runs = storage.query_runs(limit=10)
        assert len(runs) == 1
        assert runs[0]["week_no"] == 7
        assert runs[0]["campus"] == "雁山"

    def test_pipeline_storage_disabled_flag(self, tmp_path):
        detail_path, exempt_path = build_focused_sample(tmp_path)
        out_excel = tmp_path / "result_disabled.xlsx"
        db_path = tmp_path / "pipeline_storage_disabled.db"

        config = RunConfig(
            detail_path=str(detail_path),
            exempt_path=str(exempt_path),
            output_path=str(out_excel),
            week_no=7,
            db_path=str(db_path),
            enable_storage=False,
        )
        result = run(config)
        assert result.ok is True
        assert result.storage_run_id is None
        assert not os.path.exists(str(db_path))

    def test_pipeline_storage_failure_defensive_degradation(self, tmp_path, monkeypatch):
        detail_path, exempt_path = build_focused_sample(tmp_path)
        out_excel = tmp_path / "result_defensive.xlsx"
        config = RunConfig(
            detail_path=str(detail_path),
            exempt_path=str(exempt_path),
            output_path=str(out_excel),
            week_no=7,
            db_path=str(tmp_path / "dummy.db"),
            enable_storage=True,
        )

        def mock_record_run(*args, **kwargs):
            raise sqlite3.OperationalError("Mock disk IO failure")

        monkeypatch.setattr(AttendanceStorage, "record_run", mock_record_run)

        result = run(config)
        assert result.ok is True
        assert os.path.isfile(str(out_excel))
        assert any("计算记录存储异常" in a.get("type", "") for a in result.anomalies)


# ==============================================================================
# Test Class 3: TestQuotaRetentionPolicy
# ==============================================================================
class TestQuotaRetentionPolicy:
    def test_fifo_pruning_at_quota_boundary(self, tmp_path):
        db_file = tmp_path / "retention_fifo.db"
        storage = AttendanceStorage(db_path=str(db_file), max_runs=10)

        for i in range(1, 16):
            cfg, res = _make_dummy_result(week_no=i)
            storage.record_run(cfg, res)

        runs = storage.query_runs(limit=100)
        assert len(runs) == 10
        surviving_ids = [r["id"] for r in runs]
        assert min(surviving_ids) == 6
        assert max(surviving_ids) == 15

    def test_ttl_180_days_pruning(self, tmp_path):
        db_file = tmp_path / "retention_ttl.db"
        storage = AttendanceStorage(db_path=str(db_file), retention_days=180)

        cfg, res = _make_dummy_result()
        run_id_old = storage.record_run(cfg, res)
        run_id_new = storage.record_run(cfg, res)

        old_time = (datetime.now() - timedelta(days=210)).isoformat()
        with storage._get_connection() as conn:
            conn.execute("UPDATE calc_runs SET created_at = ? WHERE id = ?", (old_time, run_id_old))
            conn.commit()

        pruned = storage.cleanup_old_runs()
        assert pruned >= 1

        remaining_ids = [r["id"] for r in storage.query_runs(limit=10)]
        assert run_id_old not in remaining_ids
        assert run_id_new in remaining_ids

    def test_size_watermark_pruning_and_cleanup(self, tmp_path):
        """验证物理磁盘配额高低水位修剪机制与数据保底留存。"""
        db_file = tmp_path / "retention_size.db"
        storage = AttendanceStorage(db_path=str(db_file), max_size_mb=0.5, max_runs=500)

        # 写入带有大 notes 载荷的数据使体积膨胀
        total_batches = 10
        big_text = "A" * (64 * 1024)
        for i in range(1, total_batches + 1):
            cfg, res = _make_dummy_result(week_no=i)
            res.error = big_text
            storage.record_run(cfg, res)
            assert storage.get_run_count() > 0

        # 执行清理
        pruned = storage.cleanup_old_runs()
        assert isinstance(pruned, int)

        # 真实数据留存断言: 保留批次数 >= min_keep_runs，绝不灭失至 0
        run_count = storage.get_run_count()
        assert run_count > 0, "Catastrophic wipeout: run_count dropped to 0!"
        assert run_count >= storage.min_keep_runs, (
            f"Expected at least {storage.min_keep_runs} runs preserved, but only {run_count} survived!"
        )

        # 幸存批次应为最新的批次 (FIFO 语义)
        surviving = storage.query_runs(limit=100)
        assert len(surviving) == run_count
        surviving_weeks = sorted([r["week_no"] for r in surviving])
        expected_weeks = list(range(total_batches - run_count + 1, total_batches + 1))
        assert surviving_weeks == expected_weeks

        # 物理空间与完整性断言
        stats = storage.get_storage_stats()
        assert stats["run_count"] == run_count
        assert stats["total_physical_bytes"] == storage.get_physical_size_bytes()
        with storage._get_connection() as conn:
            assert conn.execute("PRAGMA integrity_check").fetchone()[0].lower() == "ok"

    def test_high_watermark_min_keep_protection(self, tmp_path):
        """验证在极端严苛的超小配额下，系统依然严格受 min_keep 保底约束，绝不全库清零。"""
        db_file = tmp_path / "extreme_quota.db"
        min_keep_expected = 3
        storage = AttendanceStorage(
            db_path=str(db_file),
            max_size_mb=0.05,  # 50KB 极限超小配额
            max_runs=100,
            min_keep_runs=min_keep_expected,
        )

        # 写入 8 个批次，每个批次包含溢出数据
        for i in range(1, 9):
            cfg, res = _make_dummy_result(week_no=i)
            res.error = "EXTREME_OVERFLOW_" * 2000
            storage.record_run(cfg, res)

        assert storage.get_run_count() == min_keep_expected
        stats = storage.get_storage_stats()
        assert stats["run_count"] == min_keep_expected

        surviving = storage.query_runs(limit=10)
        surviving_weeks = sorted([r["week_no"] for r in surviving])
        assert surviving_weeks == [6, 7, 8]

        with storage._get_connection() as conn:
            assert conn.execute("PRAGMA integrity_check").fetchone()[0].lower() == "ok"

    def test_dual_file_physical_size_accounting(self, tmp_path):
        """验证 WAL 模式下主库、WAL 与 SHM 多文件物理体积联合统计与截断。"""
        db_file = tmp_path / "wal_accounting.db"
        wal_file = Path(f"{db_file}-wal")
        storage = AttendanceStorage(db_path=str(db_file), max_size_mb=2.0)

        for i in range(1, 6):
            cfg, res = _make_dummy_result(week_no=i)
            res.error = "W" * (30 * 1024)
            storage.record_run(cfg, res)

        stats = storage.get_storage_stats()
        assert "file_size_bytes" in stats
        assert "wal_size_bytes" in stats
        assert "total_physical_bytes" in stats
        assert stats["file_size_bytes"] == os.path.getsize(str(db_file))
        if wal_file.exists():
            assert stats["wal_size_bytes"] == os.path.getsize(str(wal_file))
        else:
            assert stats["wal_size_bytes"] == 0

        assert stats["total_physical_bytes"] == storage.get_physical_size_bytes()

        # 执行 cleanup_old_runs 后，checkpoint(TRUNCATE) 将 WAL 物理截断
        storage.cleanup_old_runs()
        if wal_file.exists():
            assert os.path.getsize(str(wal_file)) == 0

    def test_adversarial_challenge_4_official_port(self, tmp_path):
        """官方集成 Challenger 1 的 Challenge 4 对抗测试:
        模拟高频连续写入 100KB 大载荷批次突破高水位，
        全生命周期监控每一个 record_run 步骤，确保任何一步均不出现 run_count=0 的毁灭性清空。
        """
        db_file = tmp_path / "adv_challenge4_official.db"
        storage = AttendanceStorage(db_path=str(db_file), max_size_mb=0.5, max_runs=1000)

        history_snapshots = []
        wiped_out_detected = False

        for i in range(1, 11):
            cfg, res = _make_dummy_result(week_no=i)
            res.error = "X" * (100 * 1024)  # 100KB payload
            storage.record_run(cfg, res)

            stats = storage.get_storage_stats()
            rc = stats["run_count"]
            sz = stats["file_size_bytes"]
            surviving_ids = [r["id"] for r in storage.query_runs(limit=20)]

            history_snapshots.append({
                "run_no": i,
                "run_count": rc,
                "file_size": sz,
                "surviving_ids": surviving_ids,
            })

            if rc == 0:
                wiped_out_detected = True

        final_stats = storage.get_storage_stats()
        assert not wiped_out_detected, (
            f"Adversarial Challenge 4 FAILED: Catastrophic wipeout to 0 runs detected! "
            f"Snapshots: {history_snapshots}"
        )
        assert final_stats["run_count"] >= storage.min_keep_runs
        assert storage.get_run_count() >= storage.min_keep_runs

        surviving = storage.query_runs(limit=20)
        assert len(surviving) == final_stats["run_count"]
        surviving_weeks = sorted([r["week_no"] for r in surviving])
        assert surviving_weeks == list(range(11 - final_stats["run_count"], 11))

        with storage._get_connection() as conn:
            assert conn.execute("PRAGMA integrity_check").fetchone()[0].lower() == "ok"

    def test_cleanup_idempotency_and_integrity(self, tmp_path):
        """验证清理机制的幂等性与空库防御性:
        1. 空库执行 cleanup_old_runs 返回 0 且不抛异常;
        2. 连续多次执行 cleanup_old_runs 不产生二次清理副作用或损坏结构;
        3. PRAGMA integrity_check 持续保持 ok。
        """
        db_file = tmp_path / "idempotency.db"
        storage = AttendanceStorage(db_path=str(db_file), max_size_mb=1.0)

        # 空库清理
        assert storage.cleanup_old_runs() == 0
        stats0 = storage.get_storage_stats()
        assert stats0["run_count"] == 0

        # 写入 5 条普通记录
        for i in range(1, 6):
            cfg, res = _make_dummy_result(week_no=i)
            storage.record_run(cfg, res)

        stats_before = storage.get_storage_stats()
        assert stats_before["run_count"] == 5

        # 首次显式清理 (在配额内，修剪 0 批次)
        pruned1 = storage.cleanup_old_runs()
        assert pruned1 == 0
        assert storage.get_storage_stats()["run_count"] == 5

        # 再次显式清理 (幂等性检验)
        pruned2 = storage.cleanup_old_runs()
        assert pruned2 == 0
        assert storage.get_storage_stats()["run_count"] == 5

        with storage._get_connection() as conn:
            assert conn.execute("PRAGMA integrity_check").fetchone()[0].lower() == "ok"


# ==============================================================================
# Test Class 4: TestVacuumAndIntegrity
# ==============================================================================
class TestVacuumAndIntegrity:
    def test_incremental_vacuum_and_integrity(self, tmp_path):
        db_file = tmp_path / "vacuum_test.db"
        storage = AttendanceStorage(db_path=str(db_file), max_runs=5)

        for i in range(1, 15):
            cfg, res = _make_dummy_result(week_no=i)
            storage.record_run(cfg, res)

        with storage._get_connection() as conn:
            integrity = conn.execute("PRAGMA integrity_check").fetchone()[0]
            assert integrity.lower() == "ok"

            auto_vac = conn.execute("PRAGMA auto_vacuum").fetchone()[0]
            assert auto_vac == 2

    def test_pragma_integrity_check_passes(self, storage):
        cfg, res = _make_dummy_result()
        storage.record_run(cfg, res)
        with storage._get_connection() as conn:
            check_result = conn.execute("PRAGMA integrity_check").fetchone()[0]
            assert check_result == "ok"


# ==============================================================================
# Test Class 5: TestQueryInterface
# ==============================================================================
class TestQueryInterface:
    def test_query_runs_and_details(self, storage):
        for w in (5, 6, 7, 7, 7):
            cfg, res = _make_dummy_result(week_no=w)
            storage.record_run(cfg, res)

        w7_runs = storage.query_runs(week_no=7)
        assert len(w7_runs) == 3

        paged = storage.query_runs(limit=2, offset=0)
        assert len(paged) == 2

        run_id = w7_runs[0]["id"]
        rates = storage.get_run_rates(run_id)
        assert len(rates) >= 2
        assert any(r["day_label"] == "周平均" for r in rates)

        rooms = storage.get_run_rooms(run_id, college_name="理学院")
        assert len(rooms) >= 1
        assert rooms[0]["room_name"] == "8101"

    def test_query_runs_filters_and_pagination(self, storage):
        cfg1, res1 = _make_dummy_result(week_no=1, campus="雁山", ok=True)
        cfg2, res2 = _make_dummy_result(week_no=2, campus="屏风", ok=False)
        storage.record_run(cfg1, res1)
        storage.record_run(cfg2, res2)

        runs_pingfeng = storage.query_runs(campus="屏风")
        assert len(runs_pingfeng) == 1
        assert runs_pingfeng[0]["campus"] == "屏风"

        runs_failed = storage.query_runs(status="FAILED")
        assert len(runs_failed) == 1
        assert runs_failed[0]["status"] == "FAILED"

    def test_get_run_rates_structure(self, storage):
        cfg, res = _make_dummy_result()
        run_id = storage.record_run(cfg, res)
        rates = storage.get_run_rates(run_id)
        assert len(rates) > 0
        item = rates[0]
        assert "sheet_name" in item
        assert "session" in item
        assert "college_name" in item
        assert "day_label" in item
        assert "expected_count" in item
        assert "actual_count" in item
        assert "exempt_count" in item
        assert "attendance_rate" in item
        assert "formula_status" in item

    def test_get_run_rooms_filtering(self, storage):
        cfg, res = _make_dummy_result()
        run_id = storage.record_run(cfg, res)
        rooms = storage.get_run_rooms(run_id, sheet_name="雁山（早）", college_name="理学院", day_label="周一")
        assert len(rooms) == 1
        assert rooms[0]["room_name"] == "8101"
        assert rooms[0]["exempt_source"] == "精确"

    def test_get_storage_stats(self, storage):
        cfg, res = _make_dummy_result()
        storage.record_run(cfg, res)
        stats = storage.get_storage_stats()
        assert stats["run_count"] == 1
        assert stats["rate_count"] >= 1
        assert stats["room_count"] >= 1
        assert stats["file_size_bytes"] > 0
        assert stats["max_runs"] == 100
        assert stats["max_size_mb"] == 50.0

    def test_get_run_single(self, storage):
        cfg, res = _make_dummy_result()
        run_id = storage.record_run(cfg, res)
        run_dict = storage.get_run(run_id)
        assert run_dict is not None
        assert run_dict["id"] == run_id

        non_existent = storage.get_run(999999)
        assert non_existent is None

    def test_delete_run_single(self, storage):
        cfg, res = _make_dummy_result()
        run_id = storage.record_run(cfg, res)
        assert storage.get_run(run_id) is not None

        deleted = storage.delete_run(run_id)
        assert deleted is True
        assert storage.get_run(run_id) is None

        # 删除不存在的 id
        deleted_again = storage.delete_run(run_id)
        assert deleted_again is False


# ==============================================================================
# Test Class 6: TestConcurrencyAndThreadSafety
# ==============================================================================
class TestConcurrencyAndThreadSafety:
    def test_concurrent_multithread_inserts(self, storage):
        errors: List[Exception] = []

        def worker(idx: int):
            try:
                cfg, res = _make_dummy_result(week_no=idx)
                storage.record_run(cfg, res)
            except Exception as exc:
                errors.append(exc)

        threads = [threading.Thread(target=worker, args=(i,)) for i in range(1, 6)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

        assert not errors, f"Thread errors: {errors}"
        runs = storage.query_runs(limit=10)
        assert len(runs) == 5

    def test_memory_db_mode(self):
        storage_mem = AttendanceStorage(":memory:")
        cfg, res = _make_dummy_result()
        run_id = storage_mem.record_run(cfg, res)
        assert run_id >= 1
        runs = storage_mem.query_runs()
        assert len(runs) == 1
        stats = storage_mem.get_storage_stats()
        assert stats["run_count"] == 1

    def test_disk_preflight_alert(self, tmp_path, monkeypatch):
        storage = AttendanceStorage(db_path=str(tmp_path / "preflight.db"))

        class MockDiskUsage:
            free = 5 * 1024 * 1024  # 5MB < 10MB fatal threshold

        monkeypatch.setattr(shutil, "disk_usage", lambda path: MockDiskUsage())

        cfg, res = _make_dummy_result()
        with pytest.raises(DiskQuotaExceededError):
            storage.record_run(cfg, res)
