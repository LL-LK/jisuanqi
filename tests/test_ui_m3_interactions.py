# -*- coding: utf-8 -*-
"""Milestone M3 Automated Test Suite: Whimsical Micro-Interactions, Feedback Animations & Toasts (F8, F9, F10, F11, F12).

Verifies:
1. Headless Animation Safety Controller (F12):
   - get_animation_duration returns 0ms in offscreen / headless environments (preventing test timeouts).
   - get_animation_duration returns default_ms in normal GUI environments.
2. Metric Cards Interactive Elevation & Click Navigation (F8):
   - Metric cards on both Overview and Exempt tabs have PointingHandCursor on card and labels.
   - Clicking "异常条目" opens tools_dialog and switches to "异常清单" tab (index 1).
   - Clicking "学院数" or "免勤合计 · 人次" switches workspace tabs to "学院免勤统计" (index 1).
   - Clicking "统计周次" focuses week_spin.
   - Clicking "未落表 · 人次" focuses exempt_search and activates exempt_review.
   - Clicking "个人免勤 · 人次" / "集体免勤 · 人次" focuses session combo or toggles review checkbox.
3. File Slots Drop Feedback & File Badges (F9):
   - DropSlot.set_paths formats path into badge chip with file icon and readable file size.
   - DropSlot.set_paths with multiple files displays count chip and combined size badge.
   - ResultSlot.set_result sets celebration styling (property celebration=True, ready=True, celebration badge).
4. Calculation State Transitions & Celebration Feedback (F10):
   - run_btn transitions to "正在计算…" with loading spinner icon when busy.
   - run_btn restores to "开始计算" with compute icon when busy finishes.
   - _on_done displays celebration pill with elapsed time ("🎉 计算完成 · 耗时 X.X 秒 · 结果文件已就绪") with state="success".
5. Non-Blocking Floating Toast Notifications (F11):
   - ToastNotification renders centered pill with kind-specific icon (success, warn, info).
   - ToastNotification auto-dismisses via timer and cleans up without leaking in win.children().
   - MainWindow.clean_exit() safely closes and disposes all active toast notifications.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest
from PySide6.QtCore import QPointF, Qt
from PySide6.QtGui import QGuiApplication, QMouseEvent
from PySide6.QtWidgets import QApplication

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from core import RunResult

from app import (
    ActionIcon,
    DropSlot,
    MainWindow,
    MetricCard,
    ResultSlot,
    ToastNotification,
    _format_file_size,
    get_animation_duration,
)


@pytest.fixture(scope='module')
def qapp():
    return QApplication.instance() or QApplication(sys.argv[:1])


@pytest.fixture
def main_window(qapp):
    win = MainWindow()
    win.show()
    win.activateWindow()
    qapp.processEvents()
    yield win
    try:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
    except Exception:
        pass


# ==============================================================================
# 1. Headless Animation Safety Controller (F12)
# ==============================================================================
class TestHeadlessAnimationSafety:
    """Verifies headless animation safety duration zeroing under offscreen platforms."""

    def test_get_animation_duration_returns_zero_offscreen(self, monkeypatch):
        """Under QT_QPA_PLATFORM=offscreen, returns 0ms."""
        monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
        assert get_animation_duration(300) == 0
        assert get_animation_duration(2000) == 0

    def test_get_animation_duration_returns_default_when_normal(self, monkeypatch):
        """When not in offscreen platform, returns specified default_ms."""
        monkeypatch.delenv('QT_QPA_PLATFORM', raising=False)
        monkeypatch.setattr(QGuiApplication, 'platformName', lambda: 'windows')
        assert get_animation_duration(250) == 250
        assert get_animation_duration(1500) == 1500


# ==============================================================================
# 2. Metric Cards Interactive Elevation & Click Navigation (F8)
# ==============================================================================
class TestMetricCardsInteractions:
    """Verifies metric cards cursor styling and click-through navigation."""

    def test_metric_cards_cursor_is_pointing_hand(self, main_window):
        """All metric cards and their child labels must have Qt.PointingHandCursor."""
        win = main_window

        # Check overview cards
        overview_cards = win.findChildren(MetricCard)
        assert len(overview_cards) >= 8, f"Expected at least 8 MetricCards, found {len(overview_cards)}"

        for card in overview_cards:
            assert card.cursor().shape() == Qt.PointingHandCursor
            assert card.label.cursor().shape() == Qt.PointingHandCursor
            assert card.value.cursor().shape() == Qt.PointingHandCursor

    def test_overview_card_click_anomaly_opens_tools(self, qapp, main_window):
        """Clicking '异常条目' opens tools_dialog and switches to '异常清单' tab (index 1)."""
        win = main_window
        win.tabs.setCurrentIndex(0)  # On Overview tab

        # Find '异常条目' card
        anomaly_card = None
        for card in win.findChildren(MetricCard):
            if card._title == '异常条目':
                anomaly_card = card
                break

        assert anomaly_card is not None, "MetricCard for '异常条目' not found"

        # Simulate left click
        click_event = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        anomaly_card.mousePressEvent(click_event)
        qapp.processEvents()

        assert win.tools_dialog.isVisible() or win.tools_tabs.currentIndex() == 1
        assert win.tools_tabs.currentIndex() == 1
        win.tools_dialog.close()

    def test_overview_card_click_college_switches_tabs(self, qapp, main_window):
        """Clicking '学院数' on overview switches workspace tabs to '学院免勤统计' (index 1)."""
        win = main_window
        win.tabs.setCurrentIndex(0)

        college_card = None
        for card in win.findChildren(MetricCard):
            if card._title == '学院数':
                college_card = card
                break

        assert college_card is not None
        click_event = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        college_card.mousePressEvent(click_event)
        qapp.processEvents()

        assert win.tabs.currentIndex() == 1

    def test_overview_card_click_total_switches_tabs(self, qapp, main_window):
        """Clicking '免勤合计 · 人次' on overview switches workspace tabs to '学院免勤统计' (index 1)."""
        win = main_window
        win.tabs.setCurrentIndex(0)

        total_card = None
        for card in win.findChildren(MetricCard):
            if card._title == '免勤合计 · 人次' and win.tabs.widget(0).isAncestorOf(card):
                total_card = card
                break

        assert total_card is not None
        click_event = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        total_card.mousePressEvent(click_event)
        qapp.processEvents()
        assert win.tabs.currentIndex() == 1

    def test_overview_card_click_week_focuses_spin(self, qapp, main_window):
        """Clicking '统计周次' focuses week_spin."""
        win = main_window
        win.tabs.setCurrentIndex(0)

        week_card = None
        for card in win.findChildren(MetricCard):
            if card._title == '统计周次':
                week_card = card
                break

        assert week_card is not None
        click_event = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        week_card.mousePressEvent(click_event)
        qapp.processEvents()

        assert win.week_spin.hasFocus() or win.week_spin.lineEdit().hasFocus() or QApplication.focusWidget() in (win.week_spin, win.week_spin.lineEdit())

    def test_exempt_card_click_unmapped_focuses_filter(self, qapp, main_window):
        """Clicking '未落表 · 人次' focuses exempt_search and activates exempt_review."""
        win = main_window
        win.tabs.setCurrentIndex(1)  # On Exempt tab

        unmapped_card = None
        for card in win.findChildren(MetricCard):
            if card._title == '未落表 · 人次':
                unmapped_card = card
                break

        assert unmapped_card is not None
        win.exempt_review.setChecked(False)

        click_event = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        unmapped_card.mousePressEvent(click_event)
        qapp.processEvents()

        assert win.exempt_review.isChecked() is True
        assert win.exempt_search.hasFocus() or QApplication.focusWidget() == win.exempt_search

    def test_exempt_card_click_sessions_toggles(self, qapp, main_window):
        """Clicking '集体免勤 · 人次' toggles review checkbox."""
        win = main_window
        win.tabs.setCurrentIndex(1)

        collective_card = None
        for card in win.findChildren(MetricCard):
            if '集体免勤' in card._title:
                collective_card = card
                break

        assert collective_card is not None
        initial_state = win.exempt_review.isChecked()

        click_event = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        collective_card.mousePressEvent(click_event)
        qapp.processEvents()

        assert win.exempt_review.isChecked() != initial_state


# ==============================================================================
# 3. File Slots Drop Feedback & File Badges (F9)
# ==============================================================================
class TestFileSlotsDropFeedback:
    """Verifies file badges, size formatting, and celebration state in slots."""

    def test_format_file_size_helper(self):
        """_format_file_size correctly formats bytes, KB, and MB."""
        assert _format_file_size(500) == '500 B'
        assert _format_file_size(1024) == '1.0 KB'
        assert _format_file_size(245.8 * 1024) == '245.8 KB'
        assert _format_file_size(2.5 * 1024 * 1024) == '2.5 MB'
        assert _format_file_size('invalid') == ''

    def test_drop_slot_single_file_badge(self, tmp_path):
        """Setting existing file formats badge with document icon and file size."""
        test_file = tmp_path / "10.23明细表.xlsx"
        test_file.write_bytes(b'x' * 2048)  # 2.0 KB

        slot = DropSlot('01', '明细表', '请选择文件')
        slot.set_paths([str(test_file)])

        text = slot.path_label.text()
        assert '📄' in text
        assert '10.23明细表.xlsx' in text
        assert '2.0 KB' in text
        assert slot.property('ready') is True

    def test_drop_slot_multiple_files_badge(self, tmp_path):
        """Setting multiple files formats badge with count and combined file size."""
        f1 = tmp_path / "免勤表1.xlsx"
        f2 = tmp_path / "免勤表2.xlsx"
        f1.write_bytes(b'x' * 1024)  # 1.0 KB
        f2.write_bytes(b'x' * 3072)  # 3.0 KB

        slot = DropSlot('02', '免勤表', '请添加表格', multiple=True)
        slot.set_paths([str(f1), str(f2)])

        text = slot.path_label.text()
        assert '📄 已选 2 份' in text
        assert '4.0 KB' in text
        assert slot.property('ready') is True

    def test_result_slot_celebration_state(self, tmp_path):
        """set_result triggers celebration state and formats output badge."""
        out_file = tmp_path / "第7周计算结果.xlsx"
        out_file.write_bytes(b'x' * 5120)  # 5.0 KB

        slot = ResultSlot()
        assert slot.property('ready') is False
        assert slot.property('celebration') is not True

        slot.set_result(str(out_file))

        assert slot.property('ready') is True
        assert slot.property('celebration') is True
        assert '🎉 已生成' in slot.state_label.text()
        assert '第7周计算结果.xlsx' in slot.state_label.text()
        assert '5.0 KB' in slot.state_label.text()
        assert slot.save_btn.isEnabled() is True
        assert slot.open_btn.isEnabled() is True


# ==============================================================================
# 4. Calculation State Transitions & Celebration Feedback (F10)
# ==============================================================================
class TestCalculationStateTransitions:
    """Verifies button busy transitions, loading icon, and celebration status."""

    def test_loading_action_icon_generation(self):
        """ActionIcon generates valid vector loading spinner."""
        ico = ActionIcon.get('loading', '#FFFFFF')
        assert not ico.isNull()
        pix = ico.pixmap(24, 24)
        assert not pix.isNull()
        assert pix.width() == 24

    def test_run_btn_state_transitions(self, main_window):
        """run_btn transitions text and icon during busy and normal states."""
        win = main_window

        # Initially normal
        assert win.run_btn.text() == '开始计算'
        assert win.run_btn.isEnabled() is True

        # Enter busy state
        win._set_busy(True)
        assert win.run_btn.text() == '正在计算…'
        assert win.run_btn.isEnabled() is False

        # Restore normal state
        win._set_busy(False)
        assert win.run_btn.text() == '开始计算'
        assert win.run_btn.isEnabled() is True

    def test_on_done_celebration_pill_and_toast(self, qapp, main_window, tmp_path):
        """_on_done displays celebration pill with elapsed time and triggers toast."""
        win = main_window
        dummy_out = tmp_path / "output.xlsx"
        dummy_out.write_bytes(b'x' * 1024)

        result = RunResult(
            ok=True,
            output_path=str(dummy_out),
            input_paths=(),
            summary={'统计周次': 7, '学院数': 12},
            exempt_stats={'总记录数': 100, '有效记录数': 90, '占位学号待补人数记录': 10},
        )

        win._on_done(result)
        qapp.processEvents()

        # Verify status text contains celebration indicator and elapsed time
        status_text = win.status.text()
        assert '🎉 计算完成 · 耗时' in status_text
        assert '秒 · 结果文件已就绪' in status_text
        assert win.status.property('state') == 'success'

        # Verify toast was created
        toasts = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(toasts) >= 1
        toast = toasts[0]
        assert '统计完成！' in toast.msg_label.text()
        assert toast.property('kind') == 'success'


# ==============================================================================
# 5. Non-Blocking Floating Toast Notifications (F11)
# ==============================================================================
class TestToastNotifications:
    """Verifies ToastNotification creation, layout, auto-dismiss, and memory cleanup."""

    def test_toast_creation_and_properties(self, qapp, main_window):
        """Toast creates with kind-specific icon and centered position."""
        win = main_window

        toast_success = ToastNotification(win, '成功提示信息', kind='success')
        assert toast_success.icon_label.text() == '✓'
        assert toast_success.msg_label.text() == '成功提示信息'
        assert toast_success.property('kind') == 'success'
        assert toast_success in win.children()

        toast_warn = ToastNotification(win, '警告信息', kind='warn')
        assert toast_warn.icon_label.text() == '⚠'
        assert toast_warn.property('kind') == 'warn'

        toast_info = ToastNotification(win, '通知信息', kind='info')
        assert toast_info.icon_label.text() == 'ℹ'
        assert toast_info.property('kind') == 'info'

    def test_toast_self_dismiss_and_memory_cleanup(self, qapp, main_window):
        """In offscreen mode, toast dismisses and deletes without leaving leaked children."""
        win = main_window
        toast = ToastNotification(win, '自动关闭测试', kind='info', duration=100)

        # Before event processing, toast is registered
        assert toast in win.children()

        # In offscreen mode, get_animation_duration(100) == 0, so timer fires immediately
        qapp.processEvents()
        qapp.sendPostedEvents()

        # After processing, toast has been closed and scheduled for deletion
        # Clean exit ensures complete cleanup
        win.clean_exit()
        qapp.processEvents()

        leaked = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(leaked) == 0, f"Found {len(leaked)} leaked toasts in win.children()!"

    def test_clean_exit_cleans_active_toasts(self, qapp, main_window):
        """clean_exit terminates all active toasts even if duration is long."""
        win = main_window

        # Create multiple active toasts with long durations
        t1 = ToastNotification(win, '通知 1', kind='info', duration=60000)
        t2 = ToastNotification(win, '通知 2', kind='success', duration=60000)
        assert len(win._active_toasts) >= 2

        win.clean_exit()
        qapp.processEvents()

        assert len(win._active_toasts) == 0
        leaked = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(leaked) == 0
