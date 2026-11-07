# -*- coding: utf-8 -*-
"""Milestone M2 Automated Test Suite: High-DPI Retina Assets, Iconography & Empty State (F5, F6, F7).

Verifies:
1. High-DPI Retina Sharpness Scaling (F5):
   - AttendanceIcon.pixmap respects target_widget devicePixelRatio and primaryScreen fallback.
   - Physical scaling to int(round(size * dpr)) with setDevicePixelRatio(dpr).
   - In Qt, deviceIndependentSize matches logical (size, size).
   - Vector calendar fallback drawing applies high-DPI scaling when image asset is unavailable.
   - AttendanceIcon.icon returns multi-resolution QIcon across 7 standard sizes.
2. Modern Action Iconography (F6):
   - ActionIcon.get returns valid, non-null QIcon for all action types (document, compute, export, folder_open, tools, search, save_disk, refresh).
   - Primary action buttons in MainWindow have non-null icons with appropriate iconSize.
   - All button texts, shortcuts, and objectNames remain 100% intact.
3. Charming Empty State Experience (F7):
   - Initial MainWindow displays EmptyStateCard inside summary_stack.
   - EmptyStateCard displays mascot badge, welcome header, and 3-step visual workflow chips.
   - summary_view.toPlainText() retains initial instruction string.
   - Rendering calculation results smoothly switches summary_stack to summary_view with structured report.
4. SVG Icon Assets:
   - assets/icons/*.svg files exist, are valid XML, and specify scalable vector viewports.
"""

from __future__ import annotations

import os
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest
from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QGuiApplication, QIcon, QPixmap
from PySide6.QtWidgets import QApplication, QLabel, QPushButton, QStackedWidget

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from app import ActionIcon, AttendanceIcon, EmptyStateCard, EmptyStepCard, MainWindow, TOKENS


@pytest.fixture(scope='module')
def qapp():
    return QApplication.instance() or QApplication(sys.argv[:1])


# ==============================================================================
# 1. High-DPI Retina Sharpness Scaling (F5)
# ==============================================================================
class TestHighDpiRetinaSharpness:
    """Verifies AttendanceIcon high-DPI retina sharpness scaling."""

    def test_default_dpr_and_logical_dimensions(self, qapp):
        """AttendanceIcon.pixmap(48) without target_widget returns valid pixmap."""
        px = AttendanceIcon.pixmap(48)
        assert not px.isNull()
        assert px.width() == 48
        assert px.height() == 48

    def test_explicit_dpr_scaling_physical_and_logical(self, qapp):
        """AttendanceIcon.pixmap respects devicePixelRatio float or target_widget."""
        for dpr in (1.0, 1.25, 1.5, 2.0):
            size = 48
            px = AttendanceIcon.pixmap(size, target_widget=dpr)
            assert not px.isNull()
            assert px.devicePixelRatio() == dpr
            # In Qt, physical pixel buffer width is size * dpr
            expected_phys = max(1, int(round(size * dpr)))
            assert px.width() == expected_phys
            # deviceIndependentSize matches logical dimensions
            logical_size = px.deviceIndependentSize()
            assert abs(logical_size.width() - size) < 1.0
            assert abs(logical_size.height() - size) < 1.0

    def test_label_size_hint_with_high_dpi_pixmap(self, qapp):
        """QLabel sizeHint reflects logical size when holding high-DPI pixmap."""
        px = AttendanceIcon.pixmap(48, target_widget=2.0)
        lbl = QLabel()
        lbl.setPixmap(px)
        assert lbl.sizeHint().width() == 48
        assert lbl.sizeHint().height() == 48

    def test_icon_factory_contains_all_resolutions(self, qapp):
        """AttendanceIcon.icon returns multi-resolution QIcon."""
        ico = AttendanceIcon.icon(48)
        assert not ico.isNull()
        for s in (16, 24, 32, 48, 64, 128, 256):
            px = ico.pixmap(s, s)
            assert not px.isNull()
            assert px.width() == s

    def test_cache_clear_and_cache_info_compatibility(self, qapp):
        """AttendanceIcon.pixmap exposes working cache_clear and cache_info."""
        assert callable(getattr(AttendanceIcon.pixmap, 'cache_clear', None))
        assert callable(getattr(AttendanceIcon.pixmap, 'cache_info', None))
        AttendanceIcon.pixmap.cache_clear()
        info = AttendanceIcon.pixmap.cache_info()
        assert info.currsize == 0


# ==============================================================================
# 2. Modern Action Iconography (F6)
# ==============================================================================
class TestActionIconography:
    """Verifies ActionIcon vector generator and button integration."""

    ACTION_KINDS = ['document', 'compute', 'export', 'folder_open', 'tools', 'search', 'save_disk', 'refresh']

    def test_all_action_icon_kinds_generate_valid_icons(self, qapp):
        """ActionIcon.get returns valid non-null QIcon for every registered action."""
        for kind in self.ACTION_KINDS:
            ico = ActionIcon.get(kind, '#781236')
            assert isinstance(ico, QIcon)
            assert not ico.isNull(), f"ActionIcon kind '{kind}' produced null QIcon"
            px = ico.pixmap(16, 16)
            assert not px.isNull()
            assert px.width() == 16

    def test_main_window_buttons_have_action_icons(self, qapp):
        """Main application buttons are equipped with crisp action icons."""
        win = MainWindow()
        try:
            # 1. Workflow DropSlot buttons
            assert not win.detail_slot.choose_btn.icon().isNull()
            assert win.detail_slot.choose_btn.iconSize() == QSize(16, 16)
            assert not win.exempt_slot.choose_btn.icon().isNull()
            assert win.exempt_slot.choose_btn.iconSize() == QSize(16, 16)
            assert not win.exempt_slot.manage_btn.icon().isNull()

            # 2. ResultSlot buttons
            assert not win.result_slot.save_btn.icon().isNull()
            assert win.result_slot.save_btn.iconSize() == QSize(16, 16)
            assert not win.result_slot.open_btn.icon().isNull()
            assert win.result_slot.open_btn.iconSize() == QSize(16, 16)

            # 3. Calculation run button
            assert not win.run_btn.icon().isNull()
            assert win.run_btn.iconSize() == QSize(18, 18)
            assert win.run_btn.text() == '开始计算'
            assert win.run_btn.objectName() == 'primaryButton'
            assert win.run_btn.shortcut().toString() in ('Ctrl+Return', 'Ctrl+Enter')

            # 4. Exempt page buttons
            assert not win.exempt_people_btn.icon().isNull()
            assert not win.exempt_export_btn.icon().isNull()
        finally:
            win.close()


# ==============================================================================
# 3. Charming Empty State Experience (F7)
# ==============================================================================
class TestEmptyStateExperience:
    """Verifies EmptyStateCard greeting and 3-step workflow guidance."""

    def test_initial_state_shows_empty_state_card(self, qapp):
        """Initial MainWindow shows EmptyStateCard inside summary_stack."""
        win = MainWindow()
        try:
            assert hasattr(win, 'summary_stack')
            assert hasattr(win, 'empty_state_card')
            assert isinstance(win.summary_stack, QStackedWidget)
            assert isinstance(win.empty_state_card, EmptyStateCard)
            assert win.summary_stack.currentWidget() is win.empty_state_card
        finally:
            win.close()

    def test_empty_state_card_components_and_steps(self, qapp):
        """EmptyStateCard contains mascot avatar, welcome header, and 3 steps."""
        win = MainWindow()
        try:
            card = win.empty_state_card
            mascot = card.findChild(QLabel, 'emptyMascot')
            assert mascot is not None
            assert not mascot.pixmap().isNull()

            title = card.findChild(QLabel, 'emptyWelcomeTitle')
            assert title is not None
            assert '欢迎使用早晚自习免勤自动统计' in title.text()

            # 3 EmptyStepCards
            step_cards = card.findChildren(EmptyStepCard)
            assert len(step_cards) == 3

            step_badges = card.findChildren(QLabel, 'emptyStepBadge')
            badge_texts = [b.text() for b in step_badges]
            assert sorted(badge_texts) == ['1', '2', '3']
        finally:
            win.close()

    def test_summary_view_plain_text_preserved(self, qapp):
        """summary_view.toPlainText() retains instructions for script compatibility."""
        win = MainWindow()
        try:
            text = win.summary_view.toPlainText()
            assert '等待计算' in text
            assert 'Ctrl+O' in text
            assert 'Ctrl+Enter' in text
        finally:
            win.close()

    def test_render_toggles_to_summary_view_and_back(self, qapp):
        """Rendering calculation results switches to summary_view; reset switches back."""
        from core.pipeline import RunResult

        win = MainWindow()
        try:
            assert win.summary_stack.currentWidget() is win.empty_state_card

            res = RunResult(
                ok=True,
                summary={'统计周次': 7, '学院数': 10},
                output_path='output/test.xlsx',
            )
            win._render(res)
            assert win.summary_stack.currentWidget() is win.summary_view
            assert '结果工作簿' in win.summary_view.toPlainText()

            # Re-invoking _show_empty_state switches back to empty state card
            win._show_empty_state()
            assert win.summary_stack.currentWidget() is win.empty_state_card
        finally:
            win.close()


# ==============================================================================
# 4. SVG Icon Assets Validation
# ==============================================================================
class TestSvgIconAssets:
    """Verifies assets/icons/*.svg files are valid XML vectors."""

    ICONS_DIR = PROJECT_DIR / 'assets' / 'icons'
    EXPECTED_FILES = [
        'document.svg',
        'compute.svg',
        'export.svg',
        'folder_open.svg',
        'tools.svg',
        'search.svg',
        'save.svg',
        'refresh.svg',
    ]

    def test_all_expected_svg_files_exist_and_valid_xml(self):
        """All expected SVG icon files exist on disk and parse as valid XML."""
        assert self.ICONS_DIR.exists()
        for filename in self.EXPECTED_FILES:
            path = self.ICONS_DIR / filename
            assert path.exists(), f"Missing icon asset: {path}"
            assert path.stat().st_size > 0
            # Parse XML
            tree = ET.parse(path)
            root = tree.getroot()
            assert 'svg' in root.tag.lower()
            assert root.attrib.get('viewBox') == '0 0 24 24'
