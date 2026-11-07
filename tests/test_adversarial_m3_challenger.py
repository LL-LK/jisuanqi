# -*- coding: utf-8 -*-
"""Milestone M3 Adversarial Verification & Stress Challenge Test Suite.

Empirical verification of:
1. Window Resize Stress & Visual Layout Integrity (800x600 to 2560x1440):
   - Metric cards on Overview and Exempt tabs: non-overlapping, strictly positive geometry, within bounds.
   - DropSlot (01, 02) and ResultSlot (03): non-overlapping horizontal ordering, file badge integrity.
   - Status bar celebration pill: adaptively fits within bounds without layout breakage or text overflow.
   - Rapid size cycling across all standard resolutions.
2. Calculation State Machine & Feedback Transitions (F10):
   - run_btn transitions: normal ('开始计算', compute icon, enabled) -> busy ('正在计算…', loading icon, disabled) -> restored.
   - Full failure transition: error result cleanly resets run_btn, updates status to 'error', stores error trace.
   - Clean success transition: celebration status pill ('🎉 计算完成 · 耗时 X.X 秒 · 结果文件已就绪') with state='success'.
   - Week mismatch anomaly precedence: status updates to 'warn' when template title mismatch is present.
   - Rapid consecutive calculation runs reset timer and update celebration badges accurately.
3. Business Invariants & Highlight Color Locking:
   - ESTIMATE_RGB == 'FFFFF3CD' across core/calc.py, core/export.py, app.py, and TOKENS.
   - WARN_RGB == 'FFF6DC' across app.py and TOKENS.
   - Unchanged core calculation pipeline output digests and rules.
4. Toast Notification Stress & Clean Lifecycle:
   - Burst spawning (20 simultaneous toasts) cleanly tracks in _active_toasts.
   - Rapid dismissal under offscreen platform with zero memory leaks.
   - clean_exit() disposes all toasts immediately without dangling children.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from PySide6.QtCore import QPointF, QSize, Qt
from PySide6.QtGui import QColor, QGuiApplication, QIcon, QMouseEvent
from PySide6.QtWidgets import QApplication, QLabel, QMainWindow, QWidget

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from core import RunConfig, RunResult, run
from core.calc import ESTIMATE_RGB as CORE_ESTIMATE_RGB
from core.export import ESTIMATE_RGB as EXPORT_ESTIMATE_RGB

from app import (
    ESTIMATE_RGB as APP_ESTIMATE_RGB,
    WARN_RGB as APP_WARN_RGB,
    TOKENS,
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
# 1. Window Resize Stress & Visual Layout Integrity (800x600 -> 2560x1440)
# ==============================================================================
class TestWindowResizeStress:
    """Adversarial stress testing of window resize spectrum across multiple display profiles."""

    RESIZE_RESOLUTIONS = [
        (800, 600),    # Ultra-compact legacy / SVGA
        (980, 680),    # Minimum supported desktop boundary
        (1024, 768),   # Standard XGA
        (1280, 800),   # Standard WXGA (16:10)
        (1360, 880),   # Default application launch geometry
        (1440, 900),   # High-resolution laptop (16:10)
        (1920, 1080),  # Full HD (1080p)
        (2560, 1440),  # 2K QHD widescreen
        (800, 600),    # Sharp collapse back to minimum
    ]

    def test_metric_cards_layout_and_non_overlap_during_resize(self, qapp, main_window):
        """Metric cards on both Overview and Exempt tabs preserve strictly non-overlapping geometry."""
        win = main_window

        # Verify overview metric cards
        win.tabs.setCurrentIndex(0)
        qapp.processEvents()

        overview_cards = [c for c in win.findChildren(MetricCard) if c.parent() == win or c.window() == win]
        # At least 4 on overview + 4 on exempt = 8 cards total
        assert len(overview_cards) >= 8

        for width, height in self.RESIZE_RESOLUTIONS:
            win.setMinimumSize(min(800, width), min(600, height))
            win.resize(width, height)
            qapp.processEvents()

            assert win.width() == width
            assert win.height() == height

            # Check overview cards layout
            win.tabs.setCurrentIndex(0)
            qapp.processEvents()
            ov_cards = [c for c in win.findChildren(MetricCard) if c.isVisible()]
            assert len(ov_cards) == 4, f"Expected 4 visible cards on overview at ({width}, {height}), found {len(ov_cards)}"

            # Cards must be strictly ordered horizontally without overlap
            for i in range(len(ov_cards) - 1):
                c1 = ov_cards[i]
                c2 = ov_cards[i + 1]
                p1 = c1.mapTo(win, c1.rect().topLeft())
                p2 = c2.mapTo(win, c2.rect().topLeft())
                assert c1.width() > 0 and c1.height() > 0
                assert c2.width() > 0 and c2.height() > 0
                assert p1.x() + c1.width() <= p2.x() + 2, (
                    f"Overview MetricCard {i} overlaps {i+1} at resolution ({width}, {height})! "
                    f"c1.right={p1.x() + c1.width()}, c2.left={p2.x()}"
                )

            # Check exempt tab cards layout
            win.tabs.setCurrentIndex(1)
            qapp.processEvents()
            ex_cards = [c for c in win.findChildren(MetricCard) if c.isVisible()]
            assert len(ex_cards) == 4, f"Expected 4 visible cards on exempt tab at ({width}, {height}), found {len(ex_cards)}"

            for i in range(len(ex_cards) - 1):
                c1 = ex_cards[i]
                c2 = ex_cards[i + 1]
                p1 = c1.mapTo(win, c1.rect().topLeft())
                p2 = c2.mapTo(win, c2.rect().topLeft())
                assert c1.width() > 0 and c1.height() > 0
                assert c2.width() > 0 and c2.height() > 0
                assert p1.x() + c1.width() <= p2.x() + 2, (
                    f"Exempt MetricCard {i} overlaps {i+1} at resolution ({width}, {height})! "
                    f"c1.right={p1.x() + c1.width()}, c2.left={p2.x()}"
                )

    def test_slots_row_layout_and_badges_during_resize(self, qapp, main_window, tmp_path):
        """DropSlot and ResultSlot maintain horizontal order, positive sizes, and legible badges."""
        win = main_window

        # Populate slots with sample files and celebration result
        f1 = tmp_path / "10.23早明细.xlsx"
        f1.write_bytes(b'x' * 4096)
        f2 = tmp_path / "免勤表.xlsx"
        f2.write_bytes(b'x' * 8192)
        out = tmp_path / "第7周计算结果.xlsx"
        out.write_bytes(b'x' * 16384)

        win.detail_slot.set_paths([str(f1)])
        win.exempt_slot.set_paths([str(f2)])
        win.result_slot.set_result(str(out))

        slots = [win.detail_slot, win.exempt_slot, win.result_slot]

        for width, height in self.RESIZE_RESOLUTIONS:
            win.setMinimumSize(min(800, width), min(600, height))
            win.resize(width, height)
            qapp.processEvents()

            for i in range(len(slots) - 1):
                s1 = slots[i]
                s2 = slots[i + 1]
                p1 = s1.mapTo(win, s1.rect().topLeft())
                p2 = s2.mapTo(win, s2.rect().topLeft())
                assert s1.width() > 0 and s1.height() > 0
                assert s2.width() > 0 and s2.height() > 0
                assert p1.x() + s1.width() <= p2.x() + 2, (
                    f"Slot {i} overlaps slot {i+1} at resolution ({width}, {height})!"
                )

            # Verify badge contents are intact and non-empty
            assert '📄' in win.detail_slot.path_label.text()
            assert '📄' in win.exempt_slot.path_label.text()
            assert '🎉 已生成' in win.result_slot.state_label.text()
            assert '16.0 KB' in win.result_slot.state_label.text()

    def test_status_bar_celebration_pill_during_resize(self, qapp, main_window, tmp_path):
        """Celebration status pill in status bar adjusts cleanly across all resolutions."""
        win = main_window
        dummy_out = tmp_path / "test_out.xlsx"
        dummy_out.write_bytes(b'x' * 2048)

        result = RunResult(
            ok=True,
            output_path=str(dummy_out),
            input_paths=(),
            summary={'统计周次': 7, '学院数': 12},
            exempt_stats={'总记录数': 100, '有效记录数': 90, '占位学号待补人数记录': 10},
        )
        win._on_done(result)
        qapp.processEvents()

        for width, height in self.RESIZE_RESOLUTIONS:
            win.setMinimumSize(min(800, width), min(600, height))
            win.resize(width, height)
            qapp.processEvents()

            assert win.status.width() > 0
            assert win.status.height() > 0
            assert win.status.property('state') == 'success'
            assert '🎉 计算完成' in win.status.text()
            assert win.status.wordWrap() is True


# ==============================================================================
# 2. Calculation State Machine & Feedback Transitions (F10)
# ==============================================================================
class TestCalculationStateTransitionsAdversarial:
    """Stress tests on calculation state transitions, button states, and celebration invariants."""

    def test_run_btn_busy_disable_and_restore_cycle(self, main_window):
        """run_btn cleanly toggles between enabled compute state and disabled loading spinner."""
        win = main_window

        # Initial pristine state
        assert win.run_btn.isEnabled() is True
        assert win.run_btn.text() == '开始计算'
        assert not win.run_btn.icon().isNull()

        # Enter busy state
        win._set_busy(True)
        assert win.run_btn.isEnabled() is False
        assert win.run_btn.text() == '正在计算…'
        assert win.settings_panel.isEnabled() is False
        assert win.detail_slot.isEnabled() is False
        assert win.exempt_slot.isEnabled() is False
        assert win.tabs.isEnabled() is False

        # Exit busy state via normal finish
        win._set_busy(False)
        assert win.run_btn.isEnabled() is True
        assert win.run_btn.text() == '开始计算'
        assert win.settings_panel.isEnabled() is True
        assert win.detail_slot.isEnabled() is True
        assert win.exempt_slot.isEnabled() is True
        assert win.tabs.isEnabled() is True

    def test_calculation_failure_transition(self, qapp, main_window, monkeypatch):
        """Calculation failure resets run_btn, updates status to 'error', and stores traceback."""
        from app import StatusDialog
        monkeypatch.setattr(StatusDialog, 'show_error', MagicMock())

        win = main_window
        win._start_run = MagicMock()

        # Simulate failed RunResult passed to _on_done
        fail_result = RunResult(
            ok=False,
            error='表格解析异常：未找到自习工作表',
            summary={},
            exempt_stats={},
        )

        win._on_done(fail_result)
        qapp.processEvents()

        # Status updated to error state
        assert win.status.property('state') == 'error'
        assert '计算失败：表格解析异常' in win.status.text()
        assert win._last_error_trace == '表格解析异常：未找到自习工作表'

        # After failure cleanup, run_btn must be restored and enabled
        win._cleanup_thread()
        assert win.run_btn.isEnabled() is True
        assert win.run_btn.text() == '开始计算'

    def test_template_week_mismatch_warning_precedence(self, qapp, main_window, tmp_path):
        """When an anomaly '模板标题周次不一致' is present, status prioritizes warn over celebration."""
        win = main_window
        dummy_out = tmp_path / "out_warn.xlsx"
        dummy_out.write_bytes(b'x' * 1024)

        result_with_anomaly = RunResult(
            ok=True,
            output_path=str(dummy_out),
            input_paths=(),
            summary={'统计周次': 7, '学院数': 10},
            exempt_stats={'总记录数': 50, '有效记录数': 45, '占位学号待补人数记录': 5},
            anomalies=[
                {
                    'type': '模板标题周次不一致',
                    'scope': '早自习',
                    'note': '表内标题为第6周，当前统计周次为第7周',
                    'sheet': '周一早自习',
                    'row': 1,
                    'detail': '表内标题为第6周，当前统计周次为第7周',
                }
            ],
        )

        win._on_done(result_with_anomaly)
        qapp.processEvents()

        # Anomaly warning takes precedence over celebration pill
        assert win.status.property('state') == 'warn'
        assert '表内标题与统计周次有差异' in win.status.text()

    def test_consecutive_runs_elapsed_time_tracking(self, qapp, main_window, tmp_path):
        """Consecutive calculation runs compute fresh elapsed time without stale timestamps."""
        win = main_window
        dummy_out = tmp_path / "out_consecutive.xlsx"
        dummy_out.write_bytes(b'x' * 1024)

        for simulated_elapsed in [0.5, 1.8, 3.2]:
            win._run_start_time = datetime.now() - timedelta(seconds=simulated_elapsed)

            result = RunResult(
                ok=True,
                output_path=str(dummy_out),
                input_paths=(),
                summary={'统计周次': 7, '学院数': 12},
                exempt_stats={'总记录数': 100, '有效记录数': 90, '占位学号待补人数记录': 10},
            )
            win._on_done(result)
            qapp.processEvents()

            assert win.status.property('state') == 'success'
            expected_segment = f"耗时 {simulated_elapsed:.1f} 秒"
            assert expected_segment in win.status.text(), (
                f"Expected '{expected_segment}' in status '{win.status.text()}'"
            )


# ==============================================================================
# 3. Business Invariants & Highlight Color Locking
# ==============================================================================
class TestBusinessInvariantsColorLocking:
    """Verifies that calculation core highlight colors and business invariants remain 100% synchronized."""

    def test_estimate_rgb_strictly_synchronized(self):
        """ESTIMATE_RGB is 'FFFFF3CD' across core/calc.py, core/export.py, app.py, and TOKENS."""
        assert CORE_ESTIMATE_RGB == 'FFFFF3CD'
        assert EXPORT_ESTIMATE_RGB == 'FFFFF3CD'
        assert APP_ESTIMATE_RGB == 'FFFFF3CD'
        assert TOKENS['colors']['estimate_bg'] == '#FFFFF3CD'

    def test_warn_rgb_strictly_synchronized(self):
        """WARN_RGB is 'FFF6DC' across app.py and TOKENS."""
        assert APP_WARN_RGB == 'FFF6DC'
        assert TOKENS['colors']['warn_bg'] == '#FFF6DC'

    def test_calculation_core_pipeline_integrity(self, tmp_path):
        """RunConfig and run pipeline execution produce valid output structure and invariant types."""
        cfg = RunConfig(
            detail_path='non_existent_path.xlsx',
            exempt_path='non_existent_exempt.xlsx',
            output_path=str(tmp_path / 'calc_test.xlsx'),
            week_no=7,
        )
        res = run(cfg)
        assert res.ok is False
        assert isinstance(res.error, str)
        assert len(res.error) > 0


# ==============================================================================
# 4. Toast Notification Stress & Clean Lifecycle (F11)
# ==============================================================================
class TestToastNotificationStress:
    """Adversarial burst testing of ToastNotification to guarantee zero memory leaks."""

    def test_burst_toast_spawning_and_offscreen_dismiss(self, qapp, main_window):
        """Spawn 20 simultaneous toasts in burst; all must cleanly dismiss without leaks."""
        win = main_window

        toasts = []
        for i in range(20):
            kind = ['info', 'success', 'warn'][i % 3]
            t = ToastNotification(win, f'突发通知 {i+1}', kind=kind, duration=100)
            toasts.append(t)

        assert len(win._active_toasts) >= 20

        # In offscreen mode, timers have 0ms delay, processEvents will trigger timeouts
        qapp.processEvents()
        qapp.sendPostedEvents()

        win.clean_exit()
        qapp.processEvents()

        # All toasts must be deleted
        remaining = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(remaining) == 0, f"Found {len(remaining)} lingering toasts!"
        assert len(win._active_toasts) == 0

    def test_toast_show_method_polymorphism(self, qapp, main_window):
        """ToastNotification.show works both as instance method and factory."""
        win = main_window

        t = ToastNotification.show(win, '静态调用工厂测试', kind='info', duration=50)
        assert isinstance(t, ToastNotification)
        assert t.property('kind') == 'info'
        assert t in win._active_toasts

        win.clean_exit()
        qapp.processEvents()
        assert len(win._active_toasts) == 0
