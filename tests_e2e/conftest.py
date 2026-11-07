# -*- coding: utf-8 -*-
"""E2E Test Fixtures & Environment Setup.
Ensures headless testing capability for Qt, deterministic test fixtures, and isolation.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from PySide6.QtWidgets import QApplication, QSystemTrayIcon

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests_e2e.helpers import create_detail_workbook, create_exempt_workbook

# Force headless Qt platform for non-interactive / CI / command line execution
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


@pytest.fixture(scope='session')
def qapp():
    """Session-scoped PySide6 application instance."""
    app = QApplication.instance()
    if app is None:
        app = QApplication(['attendance_e2e', '-platform', 'offscreen'])
    yield app


@pytest.fixture(autouse=True)
def ensure_tray_available(monkeypatch):
    """Ensure QSystemTrayIcon is marked available even in offscreen headless environments."""
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: True))


@pytest.fixture
def standard_detail_path(tmp_path: Path) -> Path:
    """Fixture providing a standard detail workbook."""
    p = tmp_path / 'detail_sample.xlsx'
    create_detail_workbook(p)
    return p


@pytest.fixture
def standard_exempt_path(tmp_path: Path) -> Path:
    """Fixture providing a standard exemption workbook."""
    p = tmp_path / 'exempt_sample.xlsx'
    early_rows = [
        {'sid': '9900000000001', 'college': '计算机学院', 'name': '学生甲01', 'class_text': '计算机24-1班', 'day_text': '周一', 'session': '早自习', 'room': '8101'},
        {'sid': '9900000000002', 'college': '计算机学院', 'name': '学生甲02', 'class_text': '计算机24-1班', 'day_text': '周一', 'session': '早自习', 'room': '8102'},
        {'sid': '9900000000003', 'college': '理学院', 'name': '学生乙01', 'class_text': '理学24-1班', 'day_text': '周三', 'session': '早自习', 'room': '8201'},
    ]
    late_rows = [
        {'sid': '9900000000001', 'college': '计算机学院', 'name': '学生甲01', 'class_text': '计算机24-1班', 'day_text': '周一', 'session': '晚自习', 'room': '8101'},
    ]
    create_exempt_workbook(p, early_rows=early_rows, late_rows=late_rows)
    return p
