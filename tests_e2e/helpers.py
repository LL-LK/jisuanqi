# -*- coding: utf-8 -*-
"""E2E Test Helpers & Data Generators.
Provides isolated, repeatable, synthetic test data generation and opaque-box inspection utilities.
Uses tools/gen_synthetic.py for 100% authentic template structure.
"""

from __future__ import annotations

import hashlib
import importlib.util
import os
import sqlite3
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Load gen_synthetic
spec = importlib.util.spec_from_file_location('gen_synthetic', ROOT / 'tools' / 'gen_synthetic.py')
gen_synthetic = importlib.util.module_from_spec(spec)
sys.modules['gen_synthetic'] = gen_synthetic
spec.loader.exec_module(gen_synthetic)

from core import RunConfig, RunResult
from core.pipeline import run as pipeline_run

# ---------------------------------------------------------------------------
# Dynamic feature loaders (Opaque-box inspection for planned & implemented features)
# ---------------------------------------------------------------------------

def get_storage_class():
    """Dynamically load AttendanceStorage from core.storage if implemented."""
    try:
        from core import storage
        if hasattr(storage, 'AttendanceStorage'):
            return storage.AttendanceStorage
        if hasattr(storage, 'Storage'):
            return storage.Storage
    except ImportError:
        pass
    try:
        import core.storage as storage_mod
        return getattr(storage_mod, 'AttendanceStorage', None)
    except ImportError:
        return None


def get_tamper_verifier():
    """Dynamically load verify_workbook_integrity from core.export or core.tamper."""
    for mod_name in ('core.export', 'core.tamper', 'core.excel_safety'):
        try:
            mod = __import__(mod_name, fromlist=['verify_workbook_integrity'])
            if hasattr(mod, 'verify_workbook_integrity'):
                return getattr(mod, 'verify_workbook_integrity')
        except ImportError:
            pass
    return None


def get_sheet_protection_applier():
    """Dynamically load apply_sheet_protection."""
    for mod_name in ('core.export', 'core.tamper'):
        try:
            mod = __import__(mod_name, fromlist=['apply_sheet_protection'])
            if hasattr(mod, 'apply_sheet_protection'):
                return getattr(mod, 'apply_sheet_protection')
        except ImportError:
            pass
    return None


# ---------------------------------------------------------------------------
# Synthetic Excel Generation Helpers
# ---------------------------------------------------------------------------

def _make_spec_sheet(
    name: str,
    blocks: Sequence[Tuple[str, int, Sequence[Tuple[str, int, Dict[str, Optional[int]]]]]],
    rate_labels=None,
    start_row: int = 4,
    week_no: int = 7,
):
    spec_sheet = gen_synthetic.SheetSpec(
        name=name,
        title=f'第{week_no}周三早一晚考勤统计（{name}）'
    )
    if rate_labels:
        spec_sheet.rate_labels = rate_labels
    elif '晚' in name:
        spec_sheet.rate_labels = gen_synthetic.RATE_LABELS_LATE
    else:
        spec_sheet.rate_labels = gen_synthetic.RATE_LABELS_EARLY

    row = start_row
    seq = 1
    for college, total, rooms in blocks:
        blk = gen_synthetic.BlockSpec(row, row + len(rooms) - 1, college, college_total=total)
        for room, expected, actuals in rooms:
            spec_sheet.rows.append(gen_synthetic.DetailRow(
                row=row,
                seq=seq,
                college=college,
                college_total=total,
                room=str(room),
                classes_text=f'{college}26-1班',
                expected=expected,
                actuals=dict(actuals)
            ))
            seq += 1
            row += 1
        spec_sheet.blocks.append(blk)
    return spec_sheet


def create_detail_workbook(
    path: Path,
    early_blocks: Optional[List[Tuple[str, int, List[Tuple[str, int, Dict[str, Optional[int]]]]]]] = None,
    late_blocks: Optional[List[Tuple[str, int, List[Tuple[str, int, Dict[str, Optional[int]]]]]]] = None,
    pingfeng_blocks: Optional[List[Tuple[str, int, List[Tuple[str, int, Dict[str, Optional[int]]]]]]] = None,
    week_no: int = 7,
) -> Path:
    """Create a synthetic Attendance Detail Workbook adhering to the exact template format."""
    early_blocks = early_blocks or [
        ('计算机学院', 120, [
            ('8101', 50, {'周一': 48, '周三': 47, '周四': 49, '周五': 50}),
            ('8102', 70, {'周一': 65, '周三': 66, '周四': 68, '周五': 67}),
        ]),
        ('理学院', 90, [
            ('8201', 40, {'周一': 38, '周三': 39, '周四': 40, '周五': 38}),
            ('8202', 50, {'周一': 48, '周三': 47, '周四': 49, '周五': 48}),
        ]),
    ]
    late_blocks = late_blocks or [
        ('计算机学院', 120, [
            ('8101', 50, {'周一': 49, '周三': 48, '周四': 50}),
            ('8102', 70, {'周一': 68, '周三': 67, '周四': 69}),
        ]),
    ]
    pingfeng_blocks = pingfeng_blocks or [
        ('化工学院', 60, [
            ('9101', 60, {'周一': 55, '周三': 56, '周四': 57, '周五': 58}),
        ]),
    ]

    sheets = [
        _make_spec_sheet('雁山（早）', early_blocks, week_no=week_no),
        _make_spec_sheet('雁山（晚）', late_blocks, week_no=week_no),
        _make_spec_sheet('屏风', pingfeng_blocks, week_no=week_no),
    ]

    path.parent.mkdir(parents=True, exist_ok=True)
    gen_synthetic.write_detail_workbook(path, sheets)
    return path


def _to_exempt_row(idx: int, r: Dict[str, Any]) -> gen_synthetic.ExemptRow:
    return gen_synthetic.ExemptRow(
        seq=r.get('seq', idx),
        sid=r.get('sid', f'990000000{idx:04d}'),
        college=r.get('college', '计算机学院'),
        name=r.get('name', f'测试生{idx:02d}'),
        class_text=r.get('class_text', '计算机24-1班'),
        week_text=r.get('week_text', '第7周'),
        day_text=r.get('day_text', '周一'),
        session=r.get('session', '早'),
        campus=r.get('campus', '雁山校区'),
        reason=r.get('reason', '病假'),
        room=r.get('room', ''),
    )


def create_exempt_workbook(
    path: Path,
    early_rows: Optional[List[Dict[str, Any]]] = None,
    late_rows: Optional[List[Dict[str, Any]]] = None,
) -> Path:
    """Create a synthetic Exemption Workbook with realistic schema."""
    early_objs = [_to_exempt_row(idx, r) for idx, r in enumerate(early_rows or [], 1)]
    late_objs = [_to_exempt_row(idx, r) for idx, r in enumerate(late_rows or [], 1)]

    sheets = {
        '早自习免勤人员统计表': early_objs,
        '晚自习免勤人员统计表': late_objs,
    }

    path.parent.mkdir(parents=True, exist_ok=True)
    gen_synthetic.write_exempt_workbook(path, sheets, with_room=True)
    return path


def run_pipeline_e2e(
    detail_path: Path,
    exempt_paths: Sequence[Path],
    output_path: Path,
    week_no: int = 7,
    campus: str = '雁山',
    complete_mode: bool = False,
    dictionary_path: Optional[Path] = None,
) -> RunResult:
    """Execute the core pipeline end-to-end and return the RunResult."""
    config = RunConfig(
        detail_path=str(detail_path),
        exempt_paths=tuple(str(p) for p in exempt_paths),
        output_path=str(output_path),
        week_no=week_no,
        campus=campus,
        dictionary_path=str(dictionary_path) if dictionary_path else None,
        complete_mode=complete_mode,
    )
    return pipeline_run(config)


def get_file_sha256(path: Path) -> str:
    """Compute standard SHA-256 hash of a file."""
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for chunk in iter(lambda: f.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()
