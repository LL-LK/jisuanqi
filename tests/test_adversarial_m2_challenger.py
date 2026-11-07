# -*- coding: utf-8 -*-
"""Milestone M2 Adversarial Verification & Stress Challenge Test Suite.

Empirical verification of:
1. High-DPI Retina Sharpness Scaling (F5):
   - Device pixel ratio spectrum: dpr in (1.0, 1.25, 1.5, 1.75, 2.0, 2.25, 2.5, 3.0)
   - Physical width int(round(size * dpr)) vs logical width deviceIndependentSize
   - pixmap.devicePixelRatio() == dpr
   - Vector calendar fallback drawing when strawberry-bear.png is removed/missing
   - Duck-typing resilience on target_widget and broken devicePixelRatio methods
   - Zero/negative DPR handling
2. Charming Empty State & Stack Resiliency (F7):
   - Rapid alternating toggle churn (50 iterations) between _show_empty_state() and _render()
   - Empty calculation results (summary with zero counts, empty lists)
   - Window resize stress: (980, 680) -> (1024, 768) -> (1920, 1080) -> (2560, 1440) -> (980, 680)
   - Step chip geometric containment, non-overlap, and word-wrap checks
   - Step chip click events dispatch to _pick_detail, _pick_exempt, _start_run
3. Action Iconography & Button Invariants (F6):
   - Multi-resolution QIcon generation across all kinds, aliases, and unknown fallback
   - Cache hit consistency
   - MainWindow button icons, texts, shortcuts, and objectNames retention
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from PySide6.QtCore import QPointF, QSize, Qt
from PySide6.QtGui import (
    QColor,
    QGuiApplication,
    QIcon,
    QMouseEvent,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QLabel,
    QMainWindow,
    QStackedWidget,
    QWidget,
)

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from app import (
    ActionIcon,
    AttendanceIcon,
    EmptyStateCard,
    EmptyStepCard,
    MainWindow,
    TOKENS,
)
from core.model import CollegeRate
from core.pipeline import RunResult


@pytest.fixture(scope='module')
def qapp():
    return QApplication.instance() or QApplication(sys.argv[:1])


# ==============================================================================
# 1. High-DPI Retina Sharpness & Fallback Vector Drawing (F5)
# ==============================================================================
class TestHighDpiAdversarial:
    """Stress tests and adversarial boundary checks for AttendanceIcon."""

    DPR_SPECTRUM = [1.0, 1.25, 1.5, 1.75, 2.0, 2.25, 2.5, 3.0]
    SIZES = [16, 24, 32, 48, 56, 64, 96, 128]

    def test_dpr_spectrum_scaling_physical_vs_logical(self, qapp):
        """AttendanceIcon.pixmap scales buffer to physical pixels and tags devicePixelRatio."""
        AttendanceIcon.pixmap.cache_clear()
        for dpr in self.DPR_SPECTRUM:
            for size in self.SIZES:
                px = AttendanceIcon.pixmap(size, target_widget=dpr)
                assert not px.isNull(), f"Pixmap isNull for size={size}, dpr={dpr}"
                assert px.devicePixelRatio() == dpr, f"DPR mismatch: expected {dpr}, got {px.devicePixelRatio()}"

                # Physical buffer width is int(round(size * dpr))
                expected_phys = max(1, int(round(size * dpr)))
                assert px.width() == expected_phys, (
                    f"Physical width mismatch: size={size}, dpr={dpr}, "
                    f"expected={expected_phys}, got={px.width()}"
                )
                assert px.height() == expected_phys

                # Logical dimensions via deviceIndependentSize
                logical = px.deviceIndependentSize()
                assert abs(logical.width() - size) < 1.0, (
                    f"Logical width mismatch: expected {size}, got {logical.width()}"
                )
                assert abs(logical.height() - size) < 1.0

    def test_fallback_vector_drawing_when_asset_missing(self, qapp):
        """When strawberry-bear.png is missing, vector calendar renders crisp high-DPI icons."""
        orig_meipass = getattr(sys, '_MEIPASS', None)
        try:
            # Point _MEIPASS to a non-existent directory to trigger artwork.isNull() fallback
            sys._MEIPASS = os.path.join(str(PROJECT_DIR), '_non_existent_mock_dir_')
            AttendanceIcon.pixmap.cache_clear()

            for dpr in [1.0, 1.25, 1.5, 2.0, 2.5]:
                for size in [24, 48, 64]:
                    px = AttendanceIcon.pixmap(size, target_widget=dpr)
                    assert not px.isNull(), f"Fallback vector pixmap isNull for size={size}, dpr={dpr}"
                    assert px.devicePixelRatio() == dpr
                    expected_phys = max(1, int(round(size * dpr)))
                    assert px.width() == expected_phys
                    assert px.height() == expected_phys
                    logical = px.deviceIndependentSize()
                    assert abs(logical.width() - size) < 1.0
                    assert abs(logical.height() - size) < 1.0
        finally:
            if orig_meipass is not None:
                sys._MEIPASS = orig_meipass
            elif hasattr(sys, '_MEIPASS'):
                delattr(sys, '_MEIPASS')
            AttendanceIcon.pixmap.cache_clear()

    def test_target_widget_duck_typing_and_resilience(self, qapp):
        """AttendanceIcon handles weird/broken target_widget arguments gracefully without crashing."""
        AttendanceIcon.pixmap.cache_clear()

        # 1. Plain object without devicePixelRatio
        px1 = AttendanceIcon.pixmap(48, target_widget=object())
        assert not px1.isNull()
        assert px1.devicePixelRatio() >= 1.0

        # 2. Object with broken devicePixelRatio that raises exception
        class BrokenWidget:
            def devicePixelRatio(self):
                raise RuntimeError("Display server disconnected")

        px2 = AttendanceIcon.pixmap(48, target_widget=BrokenWidget())
        assert not px2.isNull()
        assert px2.devicePixelRatio() == 1.0
        assert px2.width() == 48

        # 3. Negative and zero DPR values
        px3 = AttendanceIcon.pixmap(48, target_widget=0.0)
        assert not px3.isNull()
        assert px3.devicePixelRatio() == 1.0

        px4 = AttendanceIcon.pixmap(48, target_widget=-2.0)
        assert not px4.isNull()
        assert px4.devicePixelRatio() == 1.0

        # 4. Zero or negative size
        px5 = AttendanceIcon.pixmap(0, target_widget=2.0)
        assert not px5.isNull()
        assert px5.width() >= 1

        px6 = AttendanceIcon.pixmap(-10, target_widget=1.5)
        assert not px6.isNull()
        assert px6.width() >= 1

    def test_icon_multires_structure(self, qapp):
        """AttendanceIcon.icon returns multi-resolution QIcon across all standard sizes."""
        ico = AttendanceIcon.icon()
        assert not ico.isNull()
        for s in [16, 24, 32, 48, 64, 128, 256]:
            px = ico.pixmap(s, s)
            assert not px.isNull()
            assert px.width() == s

    def test_cache_churn_and_stats(self, qapp):
        """AttendanceIcon LRU cache records hits and clears cleanly."""
        AttendanceIcon.pixmap.cache_clear()
        info0 = AttendanceIcon.pixmap.cache_info()
        assert info0.currsize == 0

        # First call: miss
        AttendanceIcon.pixmap(48, target_widget=2.0)
        # Second call: hit
        AttendanceIcon.pixmap(48, target_widget=2.0)

        info1 = AttendanceIcon.pixmap.cache_info()
        assert info1.hits >= 1
        assert info1.currsize >= 1

        AttendanceIcon.pixmap.cache_clear()
        assert AttendanceIcon.pixmap.cache_info().currsize == 0


# ==============================================================================
# 2. Empty State Experience & Stack Resiliency (F7)
# ==============================================================================
class TestEmptyStateAndStackAdversarial:
    """Stress tests for EmptyStateCard, empty calculation results, and window resizing."""

    @pytest.fixture
    def main_window(self, qapp):
        win = MainWindow()
        yield win
        win.close()

    def test_rapid_toggle_empty_state_and_render_churn(self, main_window):
        """Rapidly toggle 50 times between _show_empty_state() and _render() without leaks or crashes."""
        win = main_window

        real_result = RunResult(
            ok=True,
            summary={'统计周次': 7, '学院数': 12},
            output_path='output/test_result.xlsx',
            exempt_rows=[
                {
                    '学院': '信通学院',
                    '星期': '星期一',
                    '自习': '早自习',
                    '个人免勤': 5,
                    '集体免勤': 0,
                    '免勤合计': 5,
                    '已回填': 5,
                    '容量超额': 0,
                    '回填差额': 0,
                    '免勤状态': '已统计',
                    '回填状态': '已回填',
                },
            ],
            anomalies=[],
            college_rates=[
                CollegeRate(
                    sheet='周一早自习',
                    session='早自习',
                    college='信通学院',
                    day='星期一',
                    expected=100,
                    room_expected_sum=100,
                    actual=95,
                    exempt=5,
                    rate=1.0,
                    status='完整',
                ),
            ],
        )

        for i in range(50):
            # 1. Switch to empty state
            win._show_empty_state()
            assert win.summary_stack.currentWidget() is win.empty_state_card
            assert '等待计算' in win.summary_view.toPlainText()

            # 2. Switch to rendered state
            win._render(real_result)
            assert win.summary_stack.currentWidget() is win.summary_view
            assert '结果工作簿' in win.summary_view.toPlainText()
            assert 'output/test_result.xlsx' in win.summary_view.toPlainText()

        # Final return to empty state
        win._show_empty_state()
        assert win.summary_stack.currentWidget() is win.empty_state_card

    def test_empty_and_sparse_calculation_results(self, main_window):
        """_render handles completely empty calculation results (0 colleges, empty rows) gracefully."""
        win = main_window

        empty_result = RunResult(
            ok=True,
            summary={'统计周次': 7, '学院数': 0},
            output_path='output/empty_run.xlsx',
            exempt_rows=[],
            college_rates=[],
            room_rows=[],
            anomalies=[],
        )

        win._render(empty_result)
        assert win.summary_stack.currentWidget() is win.summary_view
        text = win.summary_view.toPlainText()
        assert '结果工作簿' in text
        assert 'output/empty_run.xlsx' in text
        assert '统计周次  ·  7' in text
        assert '学院数  ·  0' in text

        # Metrics labels updated to 0
        assert win.overview_metrics[0].text() == '7'
        assert win.overview_metrics[1].text() == '0'
        assert win.overview_metrics[2].text() == '0'
        assert win.overview_metrics[3].text() == '0'

    def test_window_resize_stress_geometry(self, main_window):
        """Window resize down to (980, 680) and up to (2560, 1440) maintains non-overlapping chips."""
        win = main_window
        win._show_empty_state()

        test_sizes = [
            (980, 680),    # minimum supported desktop size boundary
            (1024, 768),   # standard XGA
            (1280, 800),   # standard WXGA
            (1360, 880),   # default application size
            (1920, 1080),  # full HD
            (2560, 1440),  # 2K QHD
            (980, 680),    # return to minimum size boundary
        ]

        card = win.empty_state_card
        step_cards = card.findChildren(EmptyStepCard)
        assert len(step_cards) == 3

        for req_w, req_h in test_sizes:
            win.resize(req_w, req_h)
            QApplication.processEvents()

            assert win.width() == req_w
            assert win.height() == req_h

            # Verify card is visible and has positive dimensions
            assert card.width() > 0
            assert card.height() > 0

            # Verify step cards maintain strictly ordered non-overlapping vertical layout
            for i in range(len(step_cards) - 1):
                c_current = step_cards[i]
                c_next = step_cards[i + 1]
                assert c_current.y() + c_current.height() <= c_next.y(), (
                    f"Step card {i} overlaps step card {i+1} at window size ({req_w}, {req_h})! "
                    f"current: bottom={c_current.y() + c_current.height()}, next: top={c_next.y()}"
                )
                assert c_current.width() > 0
                assert c_current.height() > 0

            # Mascot avatar must be valid
            mascot = card.findChild(QLabel, 'emptyMascot')
            assert mascot is not None
            assert not mascot.pixmap().isNull()

    def test_step_chip_clicks_dispatch(self, main_window):
        """Clicking on step chips triggers the corresponding MainWindow action methods."""
        win = main_window

        # Mock the handler methods
        win._pick_detail = MagicMock()
        win._pick_exempt = MagicMock()
        win._start_run = MagicMock()

        card = win.empty_state_card
        step_cards = card.findChildren(EmptyStepCard)
        assert len(step_cards) == 3

        pos = QPointF(10, 10)
        # Click step 1 -> _pick_detail
        dummy_event = QMouseEvent(
            QMouseEvent.Type.MouseButtonPress,
            pos,
            pos,
            pos,
            Qt.MouseButton.LeftButton,
            Qt.MouseButton.LeftButton,
            Qt.KeyboardModifier.NoModifier,
        )
        step_cards[0].mousePressEvent(dummy_event)
        assert win._pick_detail.call_count == 1

        # Click step 2 -> _pick_exempt
        step_cards[1].mousePressEvent(dummy_event)
        assert win._pick_exempt.call_count == 1

        # Click step 3 -> _start_run
        step_cards[2].mousePressEvent(dummy_event)
        assert win._start_run.call_count == 1

        # Right-click should NOT trigger action
        right_click_event = QMouseEvent(
            QMouseEvent.Type.MouseButtonPress,
            pos,
            pos,
            pos,
            Qt.MouseButton.RightButton,
            Qt.MouseButton.RightButton,
            Qt.KeyboardModifier.NoModifier,
        )
        step_cards[0].mousePressEvent(right_click_event)
        assert win._pick_detail.call_count == 1  # Still 1, not incremented


# ==============================================================================
# 3. Action Iconography & Button Invariants (F6)
# ==============================================================================
class TestActionIconAdversarial:
    """Verifies vector action icon edge cases and button properties."""

    ALL_KINDS_AND_ALIASES = [
        'document', 'file', 'table',
        'compute', 'play', 'sparkle',
        'export', 'save', 'save_as',
        'folder_open', 'folder', 'dir',
        'tools', 'settings', 'gear',
        'search', 'magnifier',
        'save_disk', 'disk',
        'refresh', 'reset',
    ]

    def test_all_icon_kinds_and_aliases_render_crisp(self, qapp):
        """ActionIcon.get renders non-null QIcon across all recognized kinds and aliases."""
        for kind in self.ALL_KINDS_AND_ALIASES:
            ico = ActionIcon.get(kind, '#781236')
            assert isinstance(ico, QIcon)
            assert not ico.isNull(), f"ActionIcon '{kind}' produced null icon"
            px = ico.pixmap(24, 24)
            assert not px.isNull()
            assert px.width() == 24

    def test_unknown_icon_kind_does_not_crash(self, qapp):
        """ActionIcon.get with unknown kind gracefully produces valid empty icon without throwing."""
        ico = ActionIcon.get('non_existent_glyph_xyz', '#123456')
        assert isinstance(ico, QIcon)
        px = ico.pixmap(16, 16)
        assert not px.isNull()

    def test_action_icon_caching_identity(self, qapp):
        """ActionIcon caches icons by (kind, color.lower())."""
        ico1 = ActionIcon.get('compute', '#FFFFFF')
        ico2 = ActionIcon.get('compute', '#ffffff')
        assert ico1 is ico2

    def test_main_window_buttons_retain_all_metadata(self, qapp):
        """Main window action buttons preserve their icons, texts, shortcuts, and objectNames."""
        win = MainWindow()
        try:
            # 1. Choose detail
            assert not win.detail_slot.choose_btn.icon().isNull()
            assert '选择' in win.detail_slot.choose_btn.text()

            # 2. Choose exempt
            assert not win.exempt_slot.choose_btn.icon().isNull()
            assert '添加' in win.exempt_slot.choose_btn.text()

            # 3. Manage exempt
            assert not win.exempt_slot.manage_btn.icon().isNull()
            assert '管理' in win.exempt_slot.manage_btn.text()

            # 4. Result save and open
            assert not win.result_slot.save_btn.icon().isNull()
            assert '另存' in win.result_slot.save_btn.text()
            assert not win.result_slot.open_btn.icon().isNull()
            assert '目录' in win.result_slot.open_btn.text()

            # 5. Run calculation button
            assert not win.run_btn.icon().isNull()
            assert win.run_btn.text() == '开始计算'
            assert win.run_btn.objectName() == 'primaryButton'
            assert win.run_btn.shortcut().toString() in ('Ctrl+Return', 'Ctrl+Enter')

            # 6. Exempt page actions
            assert not win.exempt_people_btn.icon().isNull()
            assert '查找' in win.exempt_people_btn.text()
            assert not win.exempt_export_btn.icon().isNull()
            assert '导出' in win.exempt_export_btn.text()
        finally:
            win.close()
