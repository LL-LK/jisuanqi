# -*- coding: utf-8 -*-
"""Milestone M1 Adversarial Verification & Stress Challenge Test Suite.

Empirical verification of:
1. QSS syntax validity & Qt message warnings on app.setStyleSheet(APP_QSS).
2. Widget objectName consistency with APP_QSS selectors.
3. Rapid open/close churn and lifecycle safety of StatusDialog & StudentDrilldownDialog.
4. DropSlot drag/drop dynamic property transitions (dragOver, dragReject) and error resilience.
5. Multi-resolution icon generator resilience under missing/corrupted asset files and paths.
"""

from __future__ import annotations

import io
import os
import shutil
import struct
import sys
import tempfile
from pathlib import Path
from typing import List

import pytest
import shiboken6
from PySide6.QtCore import (
    QByteArray,
    QMimeData,
    QPointF,
    QPoint,
    Qt,
    QUrl,
    qInstallMessageHandler,
    QtMsgType,
)
from PySide6.QtGui import (
    QDragEnterEvent,
    QDragLeaveEvent,
    QDropEvent,
    QGuiApplication,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QLabel,
    QPushButton,
    QWidget,
)

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import app as gui
from app import (
    APP_QSS,
    AttendanceIcon,
    BearPainter,
    COLORS,
    DropSlot,
    MainWindow,
    ResultSlot,
    StatusDialog,
    StudentDrilldownDialog,
    TOKENS,
)
from tools.generate_icon import STANDARD_SIZES, assemble_ico, generate_ico, render_png_frames


@pytest.fixture(scope="session")
def qapp():
    """Ensure QApplication instance is initialized offscreen."""
    os.environ["QT_QPA_PLATFORM"] = "offscreen"
    app = QApplication.instance()
    if app is None:
        app = QApplication([sys.argv[0], "-platform", "offscreen"])
    return app


# =============================================================================
# 1. QSS Syntax & Warning Verification
# =============================================================================

class QtLogCatcher:
    """Captures Qt warning, critical, and syntax parser messages."""

    def __init__(self):
        self.messages: List[tuple[QtMsgType, str]] = []
        self._prev_handler = None

    def __enter__(self):
        self.messages.clear()

        def handler(msg_type, context, message):
            self.messages.append((msg_type, message))

        self._prev_handler = qInstallMessageHandler(handler)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        qInstallMessageHandler(self._prev_handler)


def test_qss_syntax_zero_warnings(qapp):
    """Verify that applying APP_QSS causes 0 Qt stylesheet parser warnings or errors."""
    with QtLogCatcher() as catcher:
        qapp.setStyleSheet("")
        qapp.setStyleSheet(APP_QSS)
        qapp.processEvents()

    # Filter messages for QSS parser / CSS / style related warnings
    qss_warnings = [
        msg for msg_type, msg in catcher.messages
        if any(keyword in msg.lower() for keyword in ("css", "qss", "stylesheet", "unknown property", "selector", "parse error"))
    ]
    assert len(qss_warnings) == 0, f"QSS parser warnings detected: {qss_warnings}"


def test_qss_tokens_consistency():
    """Verify that TOKENS definitions match design system specifications."""
    assert "colors" in TOKENS
    assert "spacing" in TOKENS
    assert "radii" in TOKENS
    assert "typography" in TOKENS
    assert "elevation" in TOKENS

    # Calculation invariant color lock
    assert TOKENS["colors"]["estimate_bg"] == "#FFFFF3CD"
    assert TOKENS["colors"]["warn_bg"] == "#FFF6DC"

    # Backward-compatible COLORS dictionary
    for k, v in TOKENS["colors"].items():
        assert COLORS[k] == v


def test_qss_selectors_match_declared_object_names(qapp):
    """Verify that every ID selector declared in APP_QSS corresponds to actual widgets in UI."""
    win = MainWindow()
    try:
        # Collect all object names present in MainWindow and its children
        existing_names = {w.objectName() for w in win.findChildren(QWidget) if w.objectName()}
        existing_names.add(win.objectName())

        # Test StatusDialog object names
        dlg = StatusDialog(win, title="Test", message="Msg", advice="Advice", technical_detail="Detail")
        dlg_names = {w.objectName() for w in dlg.findChildren(QWidget) if w.objectName()}
        dlg_names.add(dlg.objectName())
        dlg.reject()

        # Test StudentDrilldownDialog object names
        drill = StudentDrilldownDialog(win, title="Drill", scope_info="Scope", students=[{"name": "Alice"}])
        drill_names = {w.objectName() for w in drill.findChildren(QWidget) if w.objectName()}
        drill_names.add(drill.objectName())
        drill.reject()

        all_names = existing_names | dlg_names | drill_names

        # Critical ID selectors in APP_QSS
        expected_ids = [
            "dropSlot", "slotBadge", "slotTitle", "slotPath",
            "settingsPanel", "metricCard", "workspaceTitle",
            "sectionTitle", "settingLabel", "metricLabel", "metricValue",
            "statusDialog", "dialogTitle", "dialogMessage",
            "adviceCard", "adviceTitle", "adviceText",
            "toggleBtn", "copyBtn", "techDetailBox",
            "drilldownCard", "drilldownCountLabel", "drilldownFilterLabel", "drilldownExportBtn",
        ]
        for eid in expected_ids:
            assert eid in all_names, f"Expected widget objectName '{eid}' not found in UI component trees!"
    finally:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()


# =============================================================================
# 2. StatusDialog & StudentDrilldownDialog Rapid Lifecycle Stress
# =============================================================================

def test_status_dialog_rapid_churn_stress(qapp):
    """Stress test StatusDialog under 100 rapid creation/destruction cycles with edge case payloads."""
    win = MainWindow()
    try:
        payloads = [
            ("Normal Title", "Simple message", "", None),
            ("Warning Title", "Warning with advice", "Check your inputs carefully.", None),
            ("Error with Detail", "Crash occurred", "Report to admin", "Traceback (most recent call last):\n  File 'app.py', line 123"),
            ("Special Characters", "Unicode: 🎓📊 测试 & <xml> 'quote' \"double\"", "Advice: emoji 🎉", "Special: \n\t\r\\n"),
            ("Massive Text", "A" * 5000, "Advice: " + "B" * 2000, "Detail: " + "C" * 10000),
        ]

        for i in range(100):
            title, msg, advice, tech = payloads[i % len(payloads)]
            dlg = StatusDialog(win, title=title, message=msg, advice=advice, technical_detail=tech)

            if tech:
                # Show dialog so widget visibility state updates properly
                dlg.show()
                qapp.processEvents()
                assert dlg.tech_box.isVisible() is False
                dlg._toggle_detail()
                qapp.processEvents()
                assert dlg.tech_box.isVisible() is True
                assert dlg.toggle_btn.text() == "隐藏技术详情"
                dlg._toggle_detail()
                qapp.processEvents()
                assert dlg.tech_box.isVisible() is False
                assert dlg.toggle_btn.text() == "显示技术详情"
                dlg._copy_detail()
                assert dlg.copy_btn.text() == "已复制"

            if i % 3 == 0:
                dlg.accept()
            elif i % 3 == 1:
                dlg.reject()
            else:
                dlg.close()

            qapp.processEvents()
            # Verify Python wrapper or C++ object is safely scheduled for deletion
            assert not shiboken6.isValid(dlg) or dlg.parent() is None

        # Assert no dangling dialogs in MainWindow's child tree
        child_dialogs = [c for c in win.children() if isinstance(c, StatusDialog) and shiboken6.isValid(c)]
        assert len(child_dialogs) == 0, f"Dangling StatusDialogs found in MainWindow: {child_dialogs}"
    finally:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()


def test_student_drilldown_dialog_rapid_churn_stress(qapp):
    """Stress test StudentDrilldownDialog with rapid creation, filtering, and edge-case student rows."""
    win = MainWindow()
    try:
        # Edge case student datasets
        empty_dataset = []
        weird_dataset = [
            {"sid": "20230001", "name": "张三", "college": "信息学院", "class_text": "计科23-1班", "reason": "早操免勤"},
            {"sid": "99999999", "name": "=CMD|' /C calc'!A0", "college": "恶意注入", "class_text": "@SUM(1+1)", "reason": "-危险字符"},
            {"sid": "", "name": "", "college": "", "class_text": "", "reason": ""},
            {"sid": "10000000", "name": "买买提·吐尔逊", "college": "地学院", "class_text": "地质23-2班", "reason": "生病"},
        ]
        large_dataset = [
            {"sid": f"2023{idx:04d}", "name": f"学生_{idx}", "college": f"学院_{idx % 5}", "class_text": f"班级_{idx % 10}", "reason": "免勤"}
            for idx in range(300)
        ]

        datasets = [empty_dataset, weird_dataset, large_dataset]

        for i in range(60):
            data = datasets[i % len(datasets)]
            dlg = StudentDrilldownDialog(win, title=f"Drilldown {i}", scope_info=f"Scope {i}", students=data)

            # Rapid filter queries
            filters = ["", "张", "2023", "不存在的姓名", "=CMD", "买买提", "学院_1", ".*+?[]()"]
            for f in filters:
                dlg._apply_filter(f)

            if i % 2 == 0:
                dlg.accept()
            else:
                dlg.reject()

            qapp.processEvents()
            assert not shiboken6.isValid(dlg) or dlg.parent() is None

        child_dialogs = [c for c in win.children() if isinstance(c, StudentDrilldownDialog) and shiboken6.isValid(c)]
        assert len(child_dialogs) == 0, f"Dangling StudentDrilldownDialogs found: {child_dialogs}"
    finally:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()


# =============================================================================
# 3. DropSlot Drag & Drop Dynamic Properties & Error Resilience
# =============================================================================

def _create_drag_event(urls: List[str], event_type: str = "enter"):
    mime = QMimeData()
    if urls:
        mime.setUrls([QUrl.fromLocalFile(p) for p in urls])
    pt = QPoint(10, 10)
    if event_type == "enter":
        ev = QDragEnterEvent(pt, Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
    elif event_type == "drop":
        ev = QDropEvent(QPointF(10, 10), Qt.CopyAction, mime, Qt.LeftButton, Qt.NoModifier)
    ev._keep_mime_alive = mime
    return ev


def test_dropslot_drag_properties_and_color_transitions(qapp):
    """Test DropSlot dynamic property transitions (dragOver, dragReject, ready) and resilience."""
    slot = DropSlot("步骤 1", "明细表槽位", "请选择自习执行情况明细表", multiple=False)

    try:
        # Initial state
        assert slot.property("dragOver") is False
        assert slot.property("dragReject") is False
        assert slot.property("ready") is False

        # 1. DragEnter with valid Excel file
        valid_event = _create_drag_event(["C:/fake/detail.xlsx"], "enter")
        slot.dragEnterEvent(valid_event)
        assert valid_event.isAccepted() is True
        assert slot.property("dragOver") is True
        assert slot.property("dragReject") is False
        assert slot.path_label.text() == "松开鼠标导入文件"

        # 2. DragLeave resets state
        leave_event = QDragLeaveEvent()
        slot.dragLeaveEvent(leave_event)
        assert slot.property("dragOver") is False
        assert slot.property("dragReject") is False

        # 3. DragEnter with invalid files (.txt, .docx, .png)
        invalid_event = _create_drag_event(["C:/fake/readme.txt", "C:/fake/photo.png"], "enter")
        slot.dragEnterEvent(invalid_event)
        assert invalid_event.isAccepted() is False
        assert slot.property("dragOver") is False
        assert slot.property("dragReject") is True
        assert slot.path_label.text() == "请选择 .xlsx 或 .xlsm 文件。"

        # 4. DragEnter with empty MIME data
        empty_mime_event = _create_drag_event([], "enter")
        slot.dragEnterEvent(empty_mime_event)
        # Should gracefully return without exception

        # 5. Drop with invalid file
        invalid_drop = _create_drag_event(["C:/fake/bad.csv"], "drop")
        slot.dropEvent(invalid_drop)
        assert invalid_drop.isAccepted() is False
        assert slot.property("dragReject") is True
        assert slot.path_label.text() == "请选择 .xlsx 或 .xlsm 文件"

        # 6. Drop with uppercase .XLSX / .XLSM
        upper_drop = _create_drag_event(["C:/fake/UPPERCASE.XLSX"], "drop")
        slot.dropEvent(upper_drop)
        assert slot.property("dragReject") is False
        assert "UPPERCASE.XLSX" in slot.path_label.text()

        # 7. Multiple slots handling
        multi_slot = DropSlot("步骤 2", "免勤表槽位", "请选择免勤表", multiple=True)
        multi_drop = _create_drag_event(["C:/fake/exempt1.xlsx", "C:/fake/exempt2.xlsx"], "drop")
        multi_slot.dropEvent(multi_drop)
        assert multi_slot.property("ready") is True
    finally:
        slot.deleteLater()
        qapp.processEvents()


# =============================================================================
# 4. Multi-Resolution Icon Generator Resilience
# =============================================================================

def test_generate_icon_missing_or_corrupt_asset_fallback(tmp_path):
    """Test generate_icon tool when assets/strawberry-bear.png is missing or corrupt."""
    test_out = str(tmp_path / "deep" / "nested" / "dir" / "test_bear.ico")

    # 1. Output directory does not exist -> should auto-create
    ico_path = generate_ico(output_path=test_out, force=True)
    assert os.path.exists(ico_path)
    assert os.path.getsize(ico_path) > 10000

    # 2. Test AttendanceIcon.pixmap fallback when PNG asset is missing
    real_asset = os.path.join(PROJECT_DIR, "assets", "strawberry-bear.png")
    temp_backup = os.path.join(PROJECT_DIR, "assets", "strawberry-bear.png.bak")

    if os.path.exists(real_asset):
        try:
            # Temporarily rename asset to simulate missing file
            shutil.move(real_asset, temp_backup)
            AttendanceIcon.pixmap.cache_clear()

            # Pixmap generation must fall back to vector painter without raising exception
            px = AttendanceIcon.pixmap(48)
            assert not px.isNull()
            assert px.width() == 48 and px.height() == 48

            # Render frames and assemble ICO
            frames = render_png_frames(AttendanceIcon, STANDARD_SIZES)
            assert len(frames) == 7
            ico_bytes = assemble_ico(frames)
            assert len(ico_bytes) > 10000

            # Test corrupted asset (0 bytes)
            with open(real_asset, "wb") as f:
                f.write(b"")
            AttendanceIcon.pixmap.cache_clear()
            px_corrupt = AttendanceIcon.pixmap(32)
            assert not px_corrupt.isNull()
            assert px_corrupt.width() == 32

            # Test corrupted asset (garbage bytes)
            with open(real_asset, "wb") as f:
                f.write(b"GARBAGE_NOT_A_VALID_IMAGE_DATA_1234567890")
            AttendanceIcon.pixmap.cache_clear()
            px_garbage = AttendanceIcon.pixmap(64)
            assert not px_garbage.isNull()
            assert px_garbage.width() == 64
        finally:
            # Restore asset
            if os.path.exists(real_asset):
                os.remove(real_asset)
            if os.path.exists(temp_backup):
                shutil.move(temp_backup, real_asset)
            AttendanceIcon.pixmap.cache_clear()

    # 3. Verify final icon valid Windows ICO binary header and entries
    with open(ico_path, "rb") as f:
        data = f.read()

    # ICONDIR: reserved (0), type (1), count (7)
    reserved, ico_type, count = struct.unpack("<HHH", data[:6])
    assert reserved == 0
    assert ico_type == 1
    assert count == 7

    # Check each ICONDIRENTRY offset and PNG magic bytes
    for i in range(count):
        entry_offset = 6 + i * 16
        w, h, _, _, planes, bpp, size, img_offset = struct.unpack("<BBBBHHII", data[entry_offset:entry_offset + 16])
        expected_size = STANDARD_SIZES[i]
        assert (w == expected_size or (expected_size == 256 and w == 0)), f"Entry {i} width mismatch"
        png_header = data[img_offset:img_offset + 8]
        assert png_header == b"\x89PNG\r\n\x1a\n", f"Entry {i} is not a valid PNG stream!"
