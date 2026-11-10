# -*- coding: utf-8 -*-
"""SQLite 本地存储与历史记录管理层。

负责考勤计算批次、学院出勤率与教室抽查明细的持久化存储，
并实施 50MB / 100 次双水位配额淘汰与物理空间整理。
"""

from __future__ import annotations

import os
import sys
import json
import uuid
import shutil
import hashlib
import sqlite3
import threading
from datetime import datetime, timedelta
from contextlib import contextmanager
from typing import Any, Dict, List, Optional, Tuple, TYPE_CHECKING

if TYPE_CHECKING:
    from .pipeline import RunConfig, RunResult


class StorageError(Exception):
    """存储层通用异常基类。"""
    pass


class DiskQuotaExceededError(StorageError):
    """物理磁盘剩余空间不足告警异常。"""
    pass


class AttendanceStorage:
    """考勤计算历史存储主控制器。"""

    DEFAULT_DB_FILENAME = "attendance_history.db"
    DEFAULT_MAX_SIZE_MB = 50.0
    DEFAULT_LOW_WATERMARK_MB = 35.0
    DEFAULT_MAX_RUNS = 100
    DEFAULT_RETENTION_DAYS = 180
    DEFAULT_MIN_KEEP_RUNS = 3
    MIN_KEEP_RUNS = DEFAULT_MIN_KEEP_RUNS
    DISK_ALERT_FREE_MB = 20.0
    DISK_FATAL_FREE_MB = 10.0

    def __init__(
        self,
        db_path: Optional[str] = None,
        max_size_mb: float = DEFAULT_MAX_SIZE_MB,
        max_runs: int = DEFAULT_MAX_RUNS,
        retention_days: int = DEFAULT_RETENTION_DAYS,
        min_keep_runs: int = DEFAULT_MIN_KEEP_RUNS,
    ):
        """初始化存储控制器。若 db_path 为空，自动解析到平台默认位置。"""
        raw_path = db_path or self.resolve_default_db_path()
        if raw_path == ":memory:":
            self._is_memory = True
            self.db_path = ":memory:"
            self._mem_conn = sqlite3.connect(":memory:", check_same_thread=False)
            self._mem_conn.row_factory = sqlite3.Row
        else:
            self._is_memory = False
            self.db_path = os.path.abspath(raw_path)
            os.makedirs(os.path.dirname(self.db_path), exist_ok=True)

        self.max_size_mb = float(max_size_mb)
        self.low_watermark_mb = min(self.DEFAULT_LOW_WATERMARK_MB, self.max_size_mb * 0.7)
        self.max_runs = int(max_runs)
        self.retention_days = int(retention_days)
        self.min_keep_runs = max(1, min(int(min_keep_runs), self.max_runs))
        self._lock = threading.RLock()

        self._init_database()

    @classmethod
    def resolve_default_db_path(cls) -> str:
        """解析平台标准用户数据目录。"""
        # 跨平台标准应用数据路径检测 (严格避免在软件目录产生 _work 残留)
        if sys.platform == "win32":
            base = os.environ.get("LOCALAPPDATA", os.path.expanduser("~\\AppData\\Local"))
            return os.path.join(base, "SanzaoWan", cls.DEFAULT_DB_FILENAME)
        elif sys.platform == "darwin":
            return os.path.expanduser(f"~/Library/Application Support/SanzaoWan/{cls.DEFAULT_DB_FILENAME}")
        else:
            base = os.environ.get("XDG_DATA_HOME", os.path.expanduser("~/.local/share"))
            return os.path.join(base, "sanzaowan", cls.DEFAULT_DB_FILENAME)

    def get_physical_size_bytes(self) -> int:
        """获取主数据库文件与 WAL、SHM 文件的真实物理磁盘占用字节总和。"""
        if self._is_memory or not os.path.exists(self.db_path):
            return 0
        total = 0
        for suffix in ("", "-wal", "-shm"):
            p = f"{self.db_path}{suffix}"
            if os.path.exists(p):
                try:
                    total += os.path.getsize(p)
                except OSError:
                    pass
        return total

    def get_run_count(self) -> int:
        """获取当前历史批次总数。"""
        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM calc_runs")
            return cur.fetchone()[0]

    def _init_database(self) -> None:
        """初始化数据库表与索引，确保启用 WAL 与 INCREMENTAL auto_vacuum。"""
        with self._lock:
            if self._is_memory:
                conn = self._mem_conn
                conn.execute("PRAGMA foreign_keys = ON;")
                self._create_schema(conn)
                conn.commit()
                return

            is_new = not os.path.exists(self.db_path) or os.path.getsize(self.db_path) == 0
            conn = sqlite3.connect(self.db_path, timeout=10.0)
            try:
                if is_new:
                    conn.execute("PRAGMA auto_vacuum = INCREMENTAL;")
                conn.execute("PRAGMA foreign_keys = ON;")
                conn.execute("PRAGMA journal_mode = WAL;")
                conn.execute("PRAGMA synchronous = NORMAL;")
                conn.execute("PRAGMA busy_timeout = 5000;")

                self._create_schema(conn)
                conn.commit()
            finally:
                conn.close()

    def _create_schema(self, conn: sqlite3.Connection) -> None:
        """执行 Schema DDL。"""
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS calc_runs (
                id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                run_uuid                TEXT NOT NULL UNIQUE,
                created_at              TEXT NOT NULL,
                week_no                 INTEGER NOT NULL,
                campus                  TEXT NOT NULL,
                status                  TEXT NOT NULL,
                detail_file_name        TEXT NOT NULL,
                detail_file_size        INTEGER NOT NULL DEFAULT 0,
                detail_sha256           TEXT NOT NULL DEFAULT '',
                exempt_files_json       TEXT NOT NULL DEFAULT '[]',
                dictionary_path         TEXT,
                dictionary_sha256       TEXT,
                output_path             TEXT NOT NULL DEFAULT '',
                output_sha256           TEXT DEFAULT '',
                total_expected          INTEGER DEFAULT 0,
                total_actual            INTEGER DEFAULT 0,
                total_exempt            INTEGER DEFAULT 0,
                overall_rate            REAL,
                anomaly_count           INTEGER DEFAULT 0,
                conflict_count          INTEGER DEFAULT 0,
                duration_ms             INTEGER DEFAULT 0,
                config_flags_json       TEXT DEFAULT '{}',
                summary_json            TEXT DEFAULT '{}',
                exempt_stats_json       TEXT DEFAULT '{}',
                notes                   TEXT DEFAULT ''
            );

            CREATE TABLE IF NOT EXISTS calc_college_rates (
                id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id                  INTEGER NOT NULL REFERENCES calc_runs(id) ON DELETE CASCADE,
                sheet_name              TEXT NOT NULL,
                session                 TEXT NOT NULL,
                college_name            TEXT NOT NULL,
                day_label               TEXT NOT NULL,
                expected_count          INTEGER,
                room_expected_sum       INTEGER DEFAULT 0,
                actual_count            INTEGER,
                exempt_count            INTEGER DEFAULT 0,
                attendance_rate         REAL,
                formula_status          TEXT NOT NULL,
                formula_note            TEXT DEFAULT '',
                anomalies_json          TEXT DEFAULT '[]'
            );

            CREATE TABLE IF NOT EXISTS calc_room_records (
                id                      INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id                  INTEGER NOT NULL REFERENCES calc_runs(id) ON DELETE CASCADE,
                sheet_name              TEXT NOT NULL,
                excel_row               INTEGER NOT NULL,
                college_name            TEXT NOT NULL,
                room_name               TEXT NOT NULL,
                classes_text            TEXT DEFAULT '',
                expected_count          INTEGER,
                day_label               TEXT NOT NULL,
                actual_count            INTEGER,
                exempt_count            INTEGER DEFAULT 0,
                exempt_source           TEXT NOT NULL DEFAULT '精确'
            );

            CREATE TABLE IF NOT EXISTS storage_maintenance_meta (
                meta_key                TEXT PRIMARY KEY,
                meta_value              TEXT NOT NULL,
                updated_at              TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS idx_calc_runs_created_at ON calc_runs(created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_calc_runs_week_campus ON calc_runs(week_no, campus);
            CREATE INDEX IF NOT EXISTS idx_calc_runs_status ON calc_runs(status);
            CREATE INDEX IF NOT EXISTS idx_college_rates_run ON calc_college_rates(run_id);
            CREATE INDEX IF NOT EXISTS idx_college_rates_search ON calc_college_rates(college_name, day_label);
            CREATE INDEX IF NOT EXISTS idx_room_records_run ON calc_room_records(run_id);
            CREATE INDEX IF NOT EXISTS idx_room_records_search ON calc_room_records(college_name, room_name, day_label);
        """)

    @contextmanager
    def _get_connection(self):
        """上下文连接工厂。"""
        if self._is_memory:
            yield self._mem_conn
            return

        conn = sqlite3.connect(self.db_path, timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA busy_timeout = 5000;")
        try:
            yield conn
        finally:
            conn.close()

    def record_run(self, config: Any, result: Any, duration_ms: int = 0) -> int:
        """持久化保存单次完整计算结果，并触发配额淘汰。"""
        self._check_disk_preflight()

        run_uuid = str(uuid.uuid4())
        created_at = datetime.now().isoformat()
        detail_path = getattr(config, "detail_path", "") or ""
        detail_name = os.path.basename(detail_path) if detail_path else ""
        detail_size = os.path.getsize(detail_path) if (detail_path and os.path.isfile(detail_path)) else 0
        detail_hash = self._calc_sha256(detail_path) if (detail_path and os.path.isfile(detail_path)) else ""

        exempt_files_meta = []
        raw_exempt_files = getattr(config, "exemption_files", None)
        if raw_exempt_files is None:
            raw_exempt_files = getattr(config, "exempt_paths", ())
            if getattr(config, "exempt_path", None):
                raw_exempt_files = (getattr(config, "exempt_path"), *raw_exempt_files)

        for p in raw_exempt_files:
            if p and os.path.isfile(p):
                exempt_files_meta.append({
                    "name": os.path.basename(p),
                    "size": os.path.getsize(p),
                    "sha256": self._calc_sha256(p)
                })
            elif p:
                exempt_files_meta.append({
                    "name": os.path.basename(p),
                    "size": 0,
                    "sha256": ""
                })

        dict_path = getattr(config, "dictionary_path", None)
        dict_hash = self._calc_sha256(dict_path) if (dict_path and os.path.isfile(dict_path)) else ""
        out_path = getattr(result, "output_path", None) or getattr(config, "output_path", "") or ""
        out_hash = self._calc_sha256(out_path) if (out_path and os.path.isfile(out_path)) else ""

        valid_rates = [
            r for r in getattr(result, "college_rates", [])
            if getattr(r, "actual", None) is not None and getattr(r, "expected", None) is not None
        ]
        total_exp = sum(r.expected for r in valid_rates)
        total_act = sum(r.actual for r in valid_rates)
        total_exm = sum(getattr(r, "exempt", 0) for r in valid_rates)
        if (total_exp - total_exm) > 0:
            overall_rate = total_act / (total_exp - total_exm)
        elif total_exp > 0:
            overall_rate = 1.0
        else:
            overall_rate = None

        anomalies = getattr(result, "anomalies", [])
        status = "FAILED" if not getattr(result, "ok", False) else ("WARNING" if anomalies else "SUCCESS")

        config_flags = {
            "complete_mode": getattr(config, "complete_mode", None),
            "match_classes": getattr(config, "match_classes", False),
            "college_exempt": getattr(result, "college_exempt", False),
            "include_exempt_summary": getattr(config, "include_exempt_summary", False),
            "exempt_source": getattr(config, "exempt_source", "未分类"),
        }

        exempt_stats = getattr(result, "exempt_stats", {})
        conflict_count = exempt_stats.get("同学号身份冲突", 0) if isinstance(exempt_stats, dict) else 0

        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO calc_runs (
                    run_uuid, created_at, week_no, campus, status,
                    detail_file_name, detail_file_size, detail_sha256,
                    exempt_files_json, dictionary_path, dictionary_sha256,
                    output_path, output_sha256, total_expected, total_actual,
                    total_exempt, overall_rate, anomaly_count, conflict_count,
                    duration_ms, config_flags_json, summary_json, exempt_stats_json, notes
                ) VALUES (
                    ?, ?, ?, ?, ?,
                    ?, ?, ?,
                    ?, ?, ?,
                    ?, ?, ?, ?,
                    ?, ?, ?, ?,
                    ?, ?, ?, ?, ?
                )
            """, (
                run_uuid, created_at, getattr(config, "week_no", 0), getattr(config, "campus", ""), status,
                detail_name, detail_size, detail_hash,
                json.dumps(exempt_files_meta, ensure_ascii=False), dict_path or "", dict_hash,
                out_path, out_hash, total_exp, total_act,
                total_exm, overall_rate, len(anomalies),
                conflict_count,
                duration_ms, json.dumps(config_flags, ensure_ascii=False),
                json.dumps(getattr(result, "summary", {}), ensure_ascii=False),
                json.dumps(exempt_stats if isinstance(exempt_stats, dict) else {}, ensure_ascii=False),
                getattr(result, "error", "") or ""
            ))
            run_id = cur.lastrowid

            # 插入学院出勤率
            rate_rows = []
            for r in getattr(result, "college_rates", []):
                anom_json = json.dumps(getattr(r, "anomalies", []), ensure_ascii=False)
                rate_rows.append((
                    run_id, getattr(r, "sheet", ""), getattr(r, "session", ""),
                    getattr(r, "college", ""), getattr(r, "day", ""),
                    getattr(r, "expected", None), getattr(r, "room_expected_sum", 0),
                    getattr(r, "actual", None), getattr(r, "exempt", 0),
                    getattr(r, "rate", None), getattr(r, "status", ""),
                    getattr(r, "note", "") or "", anom_json
                ))
            # 插入周平均汇总
            for wr in getattr(result, "weekly_rates", []):
                rate_rows.append((
                    run_id, getattr(wr, "sheet", ""), getattr(wr, "session", ""),
                    getattr(wr, "college", ""), "周平均",
                    None, 0, None, 0,
                    getattr(wr, "rate", None), "需复核" if getattr(wr, "needs_review", False) else "完整",
                    f"统计天数: {getattr(wr, 'calculated_days', 0)}/{getattr(wr, 'scheduled_days', 0)}", "[]"
                ))

            if rate_rows:
                cur.executemany("""
                    INSERT INTO calc_college_rates (
                        run_id, sheet_name, session, college_name, day_label,
                        expected_count, room_expected_sum, actual_count, exempt_count,
                        attendance_rate, formula_status, formula_note, anomalies_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, rate_rows)

            # 插入教室明细
            room_records = []
            days = ("周一", "周三", "周四", "周五")
            for room in getattr(result, "room_rows", []):
                if isinstance(room, dict):
                    sheet = room.get("工作表", "")
                    row_idx = room.get("行号", 0)
                    college = room.get("学院", "")
                    room_no = room.get("教室", "")
                    classes = room.get("分配专业班级", "")
                    expected = room.get("应到人数")
                    for d in days:
                        act_key = f"{d}实到"
                        exm_key = f"{d}免勤"
                        src_key = f"{d}免勤来源"
                        if act_key in room or exm_key in room:
                            room_records.append((
                                run_id, sheet, row_idx, college, room_no, classes, expected,
                                d, room.get(act_key), room.get(exm_key, 0) or 0, room.get(src_key, "精确") or "精确"
                            ))

            if room_records:
                cur.executemany("""
                    INSERT INTO calc_room_records (
                        run_id, sheet_name, excel_row, college_name, room_name,
                        classes_text, expected_count, day_label, actual_count,
                        exempt_count, exempt_source
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, room_records)

            conn.commit()

        # 触发配额淘汰与空间整理
        self.cleanup_old_runs()
        return run_id

    def query_runs(
        self,
        limit: int = 20,
        offset: int = 0,
        campus: Optional[str] = None,
        week_no: Optional[int] = None,
        status: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """条件检索历史计算批次记录。"""
        query = "SELECT * FROM calc_runs WHERE 1=1"
        params: List[Any] = []
        if campus:
            query += " AND campus = ?"
            params.append(campus)
        if week_no is not None:
            query += " AND week_no = ?"
            params.append(week_no)
        if status:
            query += " AND status = ?"
            params.append(status)

        query += " ORDER BY created_at DESC LIMIT ? OFFSET ?"
        params.extend([limit, offset])

        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(row) for row in cur.fetchall()]

    def get_run(self, run_id: int) -> Optional[Dict[str, Any]]:
        """获取单个计算批次的详细信息。"""
        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT * FROM calc_runs WHERE id = ?", (run_id,))
            row = cur.fetchone()
            return dict(row) if row else None

    def get_run_rates(self, run_id: int, sheet_name: Optional[str] = None) -> List[Dict[str, Any]]:
        """获取指定批次的学院出勤率明细列表。"""
        query = "SELECT * FROM calc_college_rates WHERE run_id = ?"
        params: List[Any] = [run_id]
        if sheet_name:
            query += " AND sheet_name = ?"
            params.append(sheet_name)
        query += " ORDER BY id ASC"

        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(row) for row in cur.fetchall()]

    def get_run_rooms(
        self,
        run_id: int,
        sheet_name: Optional[str] = None,
        college_name: Optional[str] = None,
        day_label: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """获取指定批次的教室级抽查与分摊明细列表。"""
        query = "SELECT * FROM calc_room_records WHERE run_id = ?"
        params: List[Any] = [run_id]
        if sheet_name:
            query += " AND sheet_name = ?"
            params.append(sheet_name)
        if college_name:
            query += " AND college_name = ?"
            params.append(college_name)
        if day_label:
            query += " AND day_label = ?"
            params.append(day_label)
        query += " ORDER BY id ASC"

        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(query, params)
            return [dict(row) for row in cur.fetchall()]

    def delete_run(self, run_id: int) -> bool:
        """显式物理删除单个批次及其下所有级联数据。"""
        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM calc_runs WHERE id = ?", (run_id,))
            conn.commit()
            deleted = cur.rowcount > 0
            if deleted and not self._is_memory:
                cur.execute("PRAGMA incremental_vacuum;").fetchall()
                conn.commit()
                try:
                    cur.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                    conn.commit()
                except sqlite3.OperationalError:
                    pass
            return deleted

    def cleanup_old_runs(self) -> int:
        """执行四阶段配额检查、FIFO淘汰与磁盘空间物理释放。

        返回被删除的批次总数。
        """
        total_pruned = 0
        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()

            # 阶段 1: TTL 期限淘汰 (retention_days)
            cutoff_date = (datetime.now() - timedelta(days=self.retention_days)).isoformat()
            cur.execute("SELECT id FROM calc_runs WHERE created_at < ?", (cutoff_date,))
            expired_ids = [row[0] for row in cur.fetchall()]
            if expired_ids:
                cur.execute("DELETE FROM calc_runs WHERE created_at < ?", (cutoff_date,))
                conn.commit()
                total_pruned += len(expired_ids)

            # 阶段 2: 批次数量淘汰 (上限 max_runs)
            cur.execute("SELECT COUNT(*) FROM calc_runs")
            total_runs = cur.fetchone()[0]
            if total_runs > self.max_runs:
                excess = total_runs - self.max_runs
                cur.execute("""
                    DELETE FROM calc_runs WHERE id IN (
                        SELECT id FROM calc_runs ORDER BY created_at ASC, id ASC LIMIT ?
                    )
                """, (excess,))
                conn.commit()
                total_pruned += excess

            # 阶段 3: 物理文件体积检查与回退淘汰 (高水位 max_size_mb -> 低水位 low_watermark_mb)
            if not self._is_memory and os.path.exists(self.db_path):
                high_watermark_bytes = int(self.max_size_mb * 1024 * 1024)
                low_watermark_bytes = int(self.low_watermark_mb * 1024 * 1024)
                current_size = self.get_physical_size_bytes()

                if current_size > high_watermark_bytes:
                    while current_size > low_watermark_bytes:
                        # 安全保底底线检查：确保至少保留 min_keep_runs 批次
                        cur.execute("SELECT COUNT(*) FROM calc_runs")
                        current_runs = cur.fetchone()[0]
                        if current_runs <= self.min_keep_runs:
                            break

                        # 获取最老的一条记录
                        cur.execute("SELECT id FROM calc_runs ORDER BY created_at ASC, id ASC LIMIT 1")
                        row = cur.fetchone()
                        if not row:
                            break

                        cur.execute("DELETE FROM calc_runs WHERE id = ?", (row[0],))
                        conn.commit()
                        total_pruned += 1

                        # 执行增量回收与 WAL 截断推进，并更新真实物理体积
                        cur.execute("PRAGMA incremental_vacuum;").fetchall()
                        conn.commit()
                        try:
                            cur.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                            conn.commit()
                        except sqlite3.OperationalError:
                            pass

                        current_size = self.get_physical_size_bytes()

            # 阶段 4: 磁盘物理空间完全收敛
            if not self._is_memory:
                cur.execute("PRAGMA incremental_vacuum;").fetchall()
                cur.execute("PRAGMA freelist_count;")
                row = cur.fetchone()
                freelist_pages = row[0] if row else 0
                if total_pruned > 0 and freelist_pages > 50:
                    conn.commit()
                    try:
                        cur.execute("VACUUM;")
                        conn.commit()
                    except sqlite3.OperationalError:
                        pass

                try:
                    cur.execute("PRAGMA wal_checkpoint(TRUNCATE);")
                    conn.commit()
                except sqlite3.OperationalError:
                    pass

            # 记录本次维护时间戳
            cur.execute("""
                INSERT INTO storage_maintenance_meta (meta_key, meta_value, updated_at)
                VALUES ('last_cleanup', ?, datetime('now'))
                ON CONFLICT(meta_key) DO UPDATE SET
                    meta_value = excluded.meta_value,
                    updated_at = excluded.updated_at
            """, (f"pruned_{total_pruned}_runs",))
            conn.commit()

        return total_pruned

    def get_storage_stats(self) -> Dict[str, Any]:
        """获取当前本地数据库的存储体量与统计指标。"""
        with self._lock, self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM calc_runs")
            run_count = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM calc_college_rates")
            rate_count = cur.fetchone()[0]
            cur.execute("SELECT COUNT(*) FROM calc_room_records")
            room_count = cur.fetchone()[0]

            freelist_count = 0
            auto_vacuum_mode = 0
            journal_mode = "memory"
            if not self._is_memory:
                cur.execute("PRAGMA freelist_count")
                r = cur.fetchone()
                if r:
                    freelist_count = r[0]
                cur.execute("PRAGMA auto_vacuum")
                r = cur.fetchone()
                if r:
                    auto_vacuum_mode = r[0]
                cur.execute("PRAGMA journal_mode")
                r = cur.fetchone()
                if r:
                    journal_mode = str(r[0])

        file_size_bytes = os.path.getsize(self.db_path) if (not self._is_memory and os.path.exists(self.db_path)) else 0
        wal_path = f"{self.db_path}-wal"
        wal_size_bytes = os.path.getsize(wal_path) if (not self._is_memory and os.path.exists(wal_path)) else 0
        shm_path = f"{self.db_path}-shm"
        shm_size_bytes = os.path.getsize(shm_path) if (not self._is_memory and os.path.exists(shm_path)) else 0
        total_physical_bytes = file_size_bytes + wal_size_bytes + shm_size_bytes

        return {
            "db_path": self.db_path,
            "file_size_bytes": file_size_bytes,
            "file_size_mb": round(file_size_bytes / (1024 * 1024), 3),
            "wal_size_bytes": wal_size_bytes,
            "shm_size_bytes": shm_size_bytes,
            "total_physical_bytes": total_physical_bytes,
            "total_physical_mb": round(total_physical_bytes / (1024 * 1024), 3),
            "run_count": run_count,
            "total_runs": run_count,
            "rate_count": rate_count,
            "room_count": room_count,
            "freelist_pages": freelist_count,
            "auto_vacuum_mode": auto_vacuum_mode,
            "journal_mode": journal_mode,
            "max_runs": self.max_runs,
            "max_size_mb": self.max_size_mb,
            "min_keep_runs": getattr(self, "min_keep_runs", self.DEFAULT_MIN_KEEP_RUNS),
            "db_size_bytes": file_size_bytes,
            "db_size_mb": round(file_size_bytes / (1024 * 1024), 3),
        }

    def _check_disk_preflight(self) -> None:
        """前置磁盘可用空间探测。"""
        if self._is_memory:
            return
        try:
            target_dir = os.path.dirname(self.db_path) or "."
            usage = shutil.disk_usage(target_dir)
            free_mb = usage.free / (1024 * 1024)
            if free_mb < self.DISK_FATAL_FREE_MB:
                raise DiskQuotaExceededError(
                    f"所在磁盘剩余空间仅 {free_mb:.1f}MB (低于致命阈值 {self.DISK_FATAL_FREE_MB}MB)，跳过历史写入。"
                )
            elif free_mb < self.DISK_ALERT_FREE_MB:
                self.cleanup_old_runs()
        except OSError:
            pass

    @staticmethod
    def _calc_sha256(file_path: str) -> str:
        """分块流式计算文件 SHA-256。"""
        if not file_path or not os.path.isfile(file_path):
            return ""
        hasher = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                while chunk := f.read(65536):
                    hasher.update(chunk)
            return hasher.hexdigest()
        except Exception:
            return ""
