# -*- coding: utf-8 -*-
"""Consolidated UI Modernization & Interaction Regression Suite (Features F1–F12).

Verifies the complete end-to-end modernization of the desktop client:
  - F1/F2: Modern design tokens (TOKENS, COLORS), centralized APP_QSS stylesheet,
           and elimination of scattered inline setStyleSheet anti-patterns.
  - F3: Business invariant color locking (ESTIMATE_RGB = 'FFFFF3CD', WARN_RGB = 'FFF6DC').
  - F4: generate_icon.py multi-resolution Windows ICO generation compatibility.
  - F5: High-DPI retina sharpness scaling on AttendanceIcon (setDevicePixelRatio).
  - F6: Modern action iconography in ActionIcon (browse, compute, export, loading spinner).
  - F7: Charming EmptyStateCard guidance and state transitions in summary_stack.
  - F8: MetricCard pointing hand cursor, hover highlight, and click-through navigation.
  - F9: DropSlot / ResultSlot file badges, human-readable byte sizing, and celebration states.
  - F10: run_btn calculation state machine transitions and elapsed time celebration pill.
  - F11: Non-blocking floating ToastNotification, auto-dismiss, and zero-leak memory cleanup.
  - F12: get_animation_duration headless safety controller.
"""

from __future__ import annotations

import os
import struct
import sys
from datetime import datetime
from pathlib import Path
from unittest.mock import MagicMock

import pytest
from PySide6.QtCore import QPointF, QRectF, QSize, Qt
from PySide6.QtGui import (
    QColor,
    QGuiApplication,
    QIcon,
    QMouseEvent,
    QPixmap,
)
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QLabel,
    QPushButton,
    QStackedWidget,
)

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

from core import RunResult
from core.calc import ESTIMATE_RGB

from app import (
    APP_QSS,
    COLORS,
    TOKENS,
    WARN_RGB,
    ActionIcon,
    AttendanceIcon,
    DropSlot,
    EmptyStateCard,
    EmptyStepCard,
    MainWindow,
    MetricCard,
    ResultSlot,
    _GenericTableModel,
    StatusDialog,
    StudentDrilldownDialog,
    ToastNotification,
    _format_file_size,
    get_animation_duration,
)
import tools.generate_icon as generate_icon_module


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
# F1 & F2: Modern Design Tokens & Unified QSS Theme Architecture
# ==============================================================================
class TestF1F2DesignTokensAndUnifiedQSS:
    """Verifies centralized design tokens, backward compatibility, and QSS rules."""

    def test_tokens_structure_and_keys(self):
        """TOKENS dictionary conforms to Google DESIGN.md specification."""
        assert isinstance(TOKENS, dict)
        required_sections = {'colors', 'spacing', 'radii', 'typography', 'elevation'}
        assert required_sections.issubset(TOKENS.keys())

        colors = TOKENS['colors']
        required_colors = {
            'primary', 'primary_dark', 'primary_light', 'bg', 'surface',
            'card', 'border', 'text_strong', 'text', 'muted',
            'estimate_bg', 'warn_bg', 'warn_fg', 'green_fg', 'danger',
        }
        assert required_colors.issubset(colors.keys())

        spacing = TOKENS['spacing']
        assert all(k in spacing for k in ('xs', 'sm', 'md', 'lg', 'xl', 'xxl'))

        radii = TOKENS['radii']
        assert all(k in radii for k in ('sm', 'md', 'lg', 'pill'))

    def test_colors_backward_compatibility(self):
        """COLORS dictionary retains all legacy tokens for backward compatibility."""
        assert isinstance(COLORS, dict)
        for key, val in TOKENS['colors'].items():
            assert key in COLORS
            assert COLORS[key] == val
        assert 'control' in COLORS
        assert 'muted_fg' in COLORS

    def test_app_qss_centralized_selectors(self):
        """APP_QSS contains definitions for key modernization components."""
        assert '#metricCard' in APP_QSS
        assert '#dropSlot' in APP_QSS
        assert '#settingsPanel' in APP_QSS
        assert '#statusDialog' in APP_QSS
        assert '#adviceCard' in APP_QSS
        assert '#emptyStateCard' in APP_QSS
        assert '#emptyStepCard' in APP_QSS
        assert '#toastNotification' in APP_QSS
        assert 'QTableView' in APP_QSS
        assert 'QHeaderView::section' in APP_QSS

    def test_absence_of_inline_setstylesheet_in_dialogs(self, qapp, main_window):
        """StatusDialog and StudentDrilldownDialog do not use inline setStyleSheet."""
        win = main_window

        # Verify StatusDialog does not have hardcoded background/border stylesheet
        dlg = StatusDialog(win, '测试标题', '测试内容', kind='info')
        assert dlg.objectName() == 'statusDialog'
        assert dlg.styleSheet() == ''  # Managed via APP_QSS #statusDialog
        dlg.close()
        dlg.deleteLater()
        qapp.processEvents()

        # Verify StudentDrilldownDialog uses centralized QSS object names
        records = [{
            '姓名': '张三', '学号': '20260101', '学院': '地球科学学院',
            '班级': '地质26-1', '周次': 7, '日期': '2026-10-23',
            '自习类型': '早自习', '缺勤类型': '缺勤', '申请免勤': '否',
            '状态': '待核对', '来源行号': '12'
        }]
        drill = StudentDrilldownDialog(win, '地球科学学院', scope_info='测试范围', students=records)
        card = drill.findChild(QFrame, 'drilldownCard')
        c_lbl = drill.findChild(QLabel, 'drilldownCountLabel')
        filter_lbl = drill.findChild(QLabel, 'drilldownFilterLabel')
        export_btn = drill.findChild(QPushButton, 'drilldownExportBtn')

        assert card is not None
        assert c_lbl is not None
        assert filter_lbl is not None
        assert export_btn is not None

        # None of these sub-widgets should have inline setStyleSheet
        assert card.styleSheet() == ''
        assert c_lbl.styleSheet() == ''
        assert filter_lbl.styleSheet() == ''
        assert export_btn.styleSheet() == ''
        drill.close()
        drill.deleteLater()
        qapp.processEvents()

    def test_drop_slot_uses_dynamic_qss_properties(self):
        """DropSlot controls drag feedback via dynamic properties rather than inline styles."""
        slot = DropSlot('01', '明细表', '请选择文件')
        assert slot.path_label.styleSheet() == ''
        assert slot.property('dragOver') is None or slot.property('dragOver') is False
        assert slot.property('dragReject') is None or slot.property('dragReject') is False


# ==============================================================================
# F3: Business Invariant Color Locking
# ==============================================================================
class TestF3BusinessColorLockingInvariants:
    """Verifies calculations highlight colors remain strictly locked to business spec."""

    def test_color_constants_match_invariants(self):
        """ESTIMATE_RGB and WARN_RGB invariant color strings are preserved."""
        assert ESTIMATE_RGB == 'FFFFF3CD'
        assert WARN_RGB == 'FFF6DC'
        assert TOKENS['colors']['estimate_bg'] == '#FFFFF3CD'
        assert TOKENS['colors']['warn_bg'] == '#FFF6DC'

    def test_room_status_model_highlight_color_mapping(self, qapp):
        """_GenericTableModel binds row highlights accurately to calculation invariants."""
        model = _GenericTableModel(
            columns=['校区', '组别', '教室', '考勤状态'],
            rows=[{'校区': '花江', '组别': 'A组', '教室': '0101', '考勤状态': '待核对'}, {'校区': '屏风', '组别': 'B组', '教室': '0202', '考勤状态': '正常'}],
            row_colors={0: QColor('#' + ESTIMATE_RGB), 1: QColor('#' + WARN_RGB)}
        )
        c0 = model.data(model.index(0, 0), Qt.BackgroundRole)
        c1 = model.data(model.index(1, 0), Qt.BackgroundRole)
        assert c0 == QColor('#FFFFF3CD')
        assert c1 == QColor('#FFF6DC')


# ==============================================================================
# F4: Icon Generator Multi-Frame Compatibility
# ==============================================================================
class TestF4IconGeneratorCompatibility:
    """Verifies tools/generate_icon.py imports and multi-frame binary assembly."""

    def test_icon_provider_exports_and_aliases(self):
        """AttendanceIcon and BearPainter aliases are accessible in app.py."""
        from app import AttendanceIcon, BearPainter
        assert BearPainter is AttendanceIcon
        assert callable(getattr(AttendanceIcon, 'pixmap', None))
        assert callable(getattr(AttendanceIcon, 'create_bear_avatar', None))

    def test_generate_ico_creates_valid_7_frame_binary(self, tmp_path):
        """generate_ico produces a valid Windows ICO file with 7 resolution frames."""
        out_ico = tmp_path / 'test_bear.ico'
        res = generate_icon_module.generate_ico(str(out_ico), force=True)
        assert os.path.exists(res)
        data = out_ico.read_bytes()
        assert len(data) > 1000

        # Validate ICO binary header: idReserved=0 (WORD), idType=1 (WORD), idCount=7 (WORD)
        reserved, ico_type, count = struct.unpack('<HHH', data[:6])
        assert reserved == 0
        assert ico_type == 1  # 1 = ICO
        assert count == len(generate_icon_module.STANDARD_SIZES) == 7


# ==============================================================================
# F5: High-DPI Retina Sharpness Scaling
# ==============================================================================
class TestF5HighDpiRetinaSharpness:
    """Verifies AttendanceIcon high-DPI scaling, setDevicePixelRatio, and multi-res icon."""

    def test_attendance_icon_pixmap_default_dpr(self, qapp):
        """AttendanceIcon.pixmap(48) returns valid logical 48x48 pixmap."""
        px = AttendanceIcon.pixmap(48)
        assert not px.isNull()
        assert px.width() == 48
        assert px.height() == 48

    def test_attendance_icon_pixmap_high_dpi_scaling(self, qapp):
        """AttendanceIcon.pixmap scales physical buffer and sets devicePixelRatio."""
        for dpr in (1.0, 1.25, 1.5, 2.0):
            size = 48
            px = AttendanceIcon.pixmap(size, target_widget=dpr)
            assert not px.isNull()
            assert px.devicePixelRatio() == dpr
            expected_phys = max(1, int(round(size * dpr)))
            assert px.width() == expected_phys
            logical_size = px.deviceIndependentSize()
            assert abs(logical_size.width() - size) < 1.0
            assert abs(logical_size.height() - size) < 1.0

    def test_attendance_icon_vector_fallback_scaling(self, qapp, monkeypatch):
        """Vector calendar fallback draws with physical scale and sets devicePixelRatio."""
        # Force image to not load by pointing MEIPASS to non-existent directory
        monkeypatch.setattr(sys, '_MEIPASS', 'non_existent_folder_xyz', raising=False)
        AttendanceIcon.pixmap.cache_clear()

        px = AttendanceIcon.pixmap(48, target_widget=2.0)
        assert not px.isNull()
        assert px.devicePixelRatio() == 2.0
        assert px.width() == 96  # 48 * 2.0

    def test_attendance_icon_multi_res_qicon(self, qapp):
        """AttendanceIcon.icon returns multi-resolution QIcon covering standard sizes."""
        ico = AttendanceIcon.icon(48)
        assert not ico.isNull()
        for s in (16, 24, 32, 48, 64, 128, 256):
            px = ico.pixmap(s, s)
            assert not px.isNull()
            assert px.width() == s

    def test_attendance_icon_cache_lifecycle(self, qapp):
        """AttendanceIcon exposes functional cache_clear and cache_info."""
        assert callable(getattr(AttendanceIcon.pixmap, 'cache_clear', None))
        assert callable(getattr(AttendanceIcon.pixmap, 'cache_info', None))
        AttendanceIcon.pixmap.cache_clear()
        info = AttendanceIcon.pixmap.cache_info()
        assert info.hits == 0


# ==============================================================================
# F6: Modern Action Iconography
# ==============================================================================
class TestF6ModernActionIconography:
    """Verifies ActionIcon vector generator and button action icon bindings."""

    def test_action_icon_all_registered_names(self, qapp):
        """ActionIcon.get returns valid non-null QIcon for all action types."""
        expected_actions = (
            'document', 'compute', 'export', 'folder_open', 'tools',
            'search', 'save_disk', 'refresh', 'loading'
        )
        for name in expected_actions:
            ico = ActionIcon.get(name)
            assert not ico.isNull(), f"ActionIcon.get('{name}') returned null QIcon"
            px = ico.pixmap(16, 16)
            assert not px.isNull()

    def test_action_icon_custom_colors_and_sizes(self, qapp):
        """ActionIcon renders with custom colors and dimensions."""
        ico_white = ActionIcon.get('compute', color='#FFFFFF')
        assert not ico_white.isNull()
        px = ico_white.pixmap(18, 18)
        assert px.width() == 18
        assert px.height() == 18

    def test_mainwindow_buttons_action_icons(self, qapp, main_window):
        """Primary action buttons in MainWindow have valid icons, text, and shortcuts."""
        win = main_window

        buttons_to_check = [
            (win.detail_slot.choose_btn, '选择文件…'),
            (win.exempt_slot.choose_btn, '添加表格…'),
            (win.exempt_slot.manage_btn, '管理…'),
            (win.result_slot.save_btn, '另存结果…'),
            (win.result_slot.open_btn, '打开所在目录'),
            (win.run_btn, '开始计算'),
            (win.exempt_people_btn, '按姓名查找…'),
            (win.exempt_export_btn, '导出当前视图…'),
        ]

        for btn, text_hint in buttons_to_check:
            assert not btn.icon().isNull(), f"Button '{btn.text()}' has null icon!"
            assert text_hint in btn.text()

        # Check run_btn shortcut is preserved
        assert win.run_btn.shortcut().toString() in ('Ctrl+Return', 'Ctrl+Enter')


# ==============================================================================
# F7: Charming Empty State Experience
# ==============================================================================
class TestF7EmptyStateCardExperience:
    """Verifies EmptyStateCard welcoming layout, step chips, and stack transitions."""

    def test_initial_state_shows_empty_state_card(self, main_window):
        """MainWindow initially displays EmptyStateCard inside summary_stack."""
        win = main_window
        assert isinstance(win.summary_stack, QStackedWidget)
        assert win.summary_stack.currentWidget() == win.empty_state_card
        assert win.empty_state_card.findChild(QLabel, 'emptyMascot') is not None
        assert win.empty_state_card.findChild(QLabel, 'emptyWelcomeTitle') is not None
        assert len(win.empty_state_card.findChildren(EmptyStepCard)) == 3

    def test_empty_step_cards_interactive_click(self, qapp, main_window, monkeypatch):
        """Clicking each step card triggers the corresponding action."""
        win = main_window
        mock_pick_detail = MagicMock()
        mock_pick_exempt = MagicMock()
        mock_start_run = MagicMock()
        monkeypatch.setattr(win, '_pick_detail', mock_pick_detail)
        monkeypatch.setattr(win, '_pick_exempt', mock_pick_exempt)
        monkeypatch.setattr(win, '_start_run', mock_start_run)

        cards = win.empty_state_card.findChildren(EmptyStepCard)
        assert len(cards) == 3
        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)

        cards[0].mousePressEvent(ev)
        assert mock_pick_detail.call_count == 1

        cards[1].mousePressEvent(ev)
        assert mock_pick_exempt.call_count == 1

        cards[2].mousePressEvent(ev)
        assert mock_start_run.call_count == 1

    def test_state_transitions_between_empty_and_rendered(self, qapp, main_window):
        """summary_stack switches between empty state card and calculation summary report."""
        win = main_window
        assert win.summary_stack.currentWidget() == win.empty_state_card

        mock_result = RunResult(
            ok=True,
            output_path="test_out.xlsx",
            input_paths=(),
            summary={'统计周次': 7, '学院数': 10},
            exempt_stats={'总记录数': 100, '有效记录数': 90, '占位学号待补人数记录': 10},
        )

        win._render(mock_result)
        qapp.processEvents()

        # Stack should now show summary_view (QPlainTextEdit)
        assert win.summary_stack.currentWidget() == win.summary_view
        assert '结果工作簿' in win.summary_view.toPlainText()

        # Calling _show_empty_state restores the empty state card
        win._show_empty_state()
        qapp.processEvents()
        assert win.summary_stack.currentWidget() == win.empty_state_card


# ==============================================================================
# F8: MetricCard Interactive Elevation & Navigation
# ==============================================================================
class TestF8MetricCardInteractions:
    """Verifies MetricCard cursor, hover highlight, and click-through navigation."""

    def test_metric_card_cursor_and_elevation(self, main_window):
        """All MetricCards and labels have Qt.PointingHandCursor."""
        win = main_window
        cards = win.findChildren(MetricCard)
        assert len(cards) >= 8
        for card in cards:
            assert card.cursor().shape() == Qt.PointingHandCursor
            assert card.label.cursor().shape() == Qt.PointingHandCursor
            assert card.value.cursor().shape() == Qt.PointingHandCursor

    def test_overview_card_click_anomaly_opens_tools(self, qapp, main_window):
        """Clicking '异常条目' opens tools_dialog and switches to '异常清单' tab (index 1)."""
        win = main_window
        win.tabs.setCurrentIndex(0)
        qapp.processEvents()

        anomaly_card = next(c for c in win.findChildren(MetricCard) if c._title == '异常条目')
        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        anomaly_card.mousePressEvent(ev)
        qapp.processEvents()

        assert win.tools_dialog.isVisible() or win.tools_tabs.currentIndex() == 1
        assert win.tools_tabs.currentIndex() == 1
        win.tools_dialog.close()

    def test_overview_card_click_college_switches_tabs(self, qapp, main_window):
        """Clicking '学院数' switches workspace tabs to '学院免勤统计' (index 1)."""
        win = main_window
        win.tabs.setCurrentIndex(0)
        qapp.processEvents()

        college_card = next(c for c in win.findChildren(MetricCard) if c._title == '学院数')
        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        college_card.mousePressEvent(ev)
        qapp.processEvents()
        assert win.tabs.currentIndex() == 1

    def test_overview_card_click_total_switches_tabs(self, qapp, main_window):
        """Clicking '免勤合计 · 人次' on overview switches workspace tabs to index 1."""
        win = main_window
        win.tabs.setCurrentIndex(0)
        qapp.processEvents()

        total_card = next(c for c in win.findChildren(MetricCard) if c._title == '免勤合计 · 人次' and win.tabs.widget(0).isAncestorOf(c))
        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        total_card.mousePressEvent(ev)
        qapp.processEvents()
        assert win.tabs.currentIndex() == 1

    def test_overview_card_click_week_focuses_spin(self, qapp, main_window):
        """Clicking '统计周次' focuses week_spin."""
        win = main_window
        win.tabs.setCurrentIndex(0)
        qapp.processEvents()

        week_card = next(c for c in win.findChildren(MetricCard) if c._title == '统计周次')
        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        week_card.mousePressEvent(ev)
        qapp.processEvents()
        assert win.week_spin.hasFocus() or win.week_spin.lineEdit().hasFocus() or QApplication.focusWidget() in (win.week_spin, win.week_spin.lineEdit())

    def test_exempt_card_click_unmapped_focuses_filter(self, qapp, main_window):
        """Clicking '未落表 · 人次' focuses exempt_search and activates exempt_review."""
        win = main_window
        win.tabs.setCurrentIndex(1)
        qapp.processEvents()

        unmapped_card = next(c for c in win.findChildren(MetricCard) if c._title == '未落表 · 人次')
        win.exempt_review.setChecked(False)

        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        unmapped_card.mousePressEvent(ev)
        qapp.processEvents()

        assert win.exempt_review.isChecked() is True
        assert win.exempt_search.hasFocus() or QApplication.focusWidget() == win.exempt_search

    def test_exempt_card_click_collective_toggles_review(self, qapp, main_window):
        """Clicking '集体免勤 · 人次' toggles review checkbox."""
        win = main_window
        win.tabs.setCurrentIndex(1)
        qapp.processEvents()

        collective_card = next(c for c in win.findChildren(MetricCard) if '集体免勤' in c._title)
        initial_state = win.exempt_review.isChecked()

        ev = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.LeftButton, Qt.LeftButton, Qt.NoModifier)
        collective_card.mousePressEvent(ev)
        qapp.processEvents()

        assert win.exempt_review.isChecked() != initial_state

    def test_metric_card_non_left_click_ignored(self, qapp, main_window):
        """Right click does not trigger navigation."""
        win = main_window
        win.tabs.setCurrentIndex(0)
        qapp.processEvents()

        college_card = next(c for c in win.findChildren(MetricCard) if c._title == '学院数' and win.tabs.widget(0).isAncestorOf(c))
        right_click = QMouseEvent(QMouseEvent.Type.MouseButtonPress, QPointF(10, 10), QPointF(10, 10), Qt.RightButton, Qt.RightButton, Qt.NoModifier)
        college_card.mousePressEvent(right_click)
        qapp.processEvents()
        assert win.tabs.currentIndex() == 0  # Still on tab 0


# ==============================================================================
# F9: File Slots Drop Feedback & File Badges
# ==============================================================================
class TestF9FileSlotsDropFeedback:
    """Verifies file badges, size formatting, and celebration state in slots."""

    def test_file_size_formatter(self):
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

    def test_result_slot_celebration_badge(self, tmp_path):
        """ResultSlot.set_result sets celebration state, ready flag, and badge."""
        res_file = tmp_path / "第7周结果.xlsx"
        res_file.write_bytes(b'x' * 5120)  # 5.0 KB

        slot = ResultSlot()
        assert slot.property('ready') is False
        assert slot.property('celebration') is not True

        slot.set_result(str(res_file))
        assert slot.property('ready') is True
        assert slot.property('celebration') is True
        text = slot.state_label.text()
        assert '🎉 已生成' in text
        assert '第7周结果.xlsx' in text
        assert '5.0 KB' in text
        assert slot.save_btn.isEnabled() is True
        assert slot.open_btn.isEnabled() is True


# ==============================================================================
# F10: Calculation State Transitions & Celebration Feedback
# ==============================================================================
class TestF10CalculationStateTransitions:
    """Verifies run_btn busy state transitions and completion celebration."""

    def test_run_btn_state_transitions(self, qapp, main_window):
        """run_btn transitions between active and busy states."""
        win = main_window

        win._set_busy(True)
        qapp.processEvents()
        assert win.run_btn.isEnabled() is False
        assert win.run_btn.text() == '正在计算…'

        win._set_busy(False)
        qapp.processEvents()
        assert win.run_btn.isEnabled() is True
        assert win.run_btn.text() == '开始计算'

    def test_on_done_celebration_status_and_toast(self, qapp, main_window, tmp_path):
        """_on_done displays celebration pill with elapsed time and creates toast."""
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

        win._run_start_time = datetime.now()
        win._on_done(result)
        qapp.processEvents()

        status_text = win.status.text()
        assert '🎉 计算完成 · 耗时' in status_text
        assert '秒 · 结果文件已就绪' in status_text
        assert win.status.property('state') == 'success'

        toasts = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(toasts) >= 1
        assert '统计完成！' in toasts[0].msg_label.text()
        assert toasts[0].property('kind') == 'success'


# ==============================================================================
# F11: Non-Blocking Floating Toast Notifications
# ==============================================================================
class TestF11ToastNotifications:
    """Verifies ToastNotification creation, auto-dismiss, and memory cleanup."""

    def test_toast_kinds_and_properties(self, qapp, main_window):
        """Toast formats with kind-specific icon and registers in parent's active set."""
        win = main_window

        t_success = ToastNotification(win, '操作成功', kind='success')
        assert t_success.icon_label.text() == '✓'
        assert t_success.msg_label.text() == '操作成功'
        assert t_success.property('kind') == 'success'
        assert t_success in win._active_toasts

        t_warn = ToastNotification(win, '注意警告', kind='warn')
        assert t_warn.icon_label.text() == '⚠'
        assert t_warn.property('kind') == 'warn'

        t_info = ToastNotification(win, '普通通知', kind='info')
        assert t_info.icon_label.text() == 'ℹ'
        assert t_info.property('kind') == 'info'

    def test_toast_auto_dismiss_and_memory_cleanup(self, qapp, main_window):
        """In offscreen mode, toast dismisses via 0ms timer and cleans up without leaking."""
        win = main_window
        toast = ToastNotification(win, '自动关闭测试', kind='info', duration=100)
        assert toast in win.children()

        # In offscreen mode, get_animation_duration(100) == 0, so timer fires on event loop
        qapp.processEvents()
        qapp.sendPostedEvents()

        win.clean_exit()
        qapp.processEvents()

        leaked = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(leaked) == 0, f"Found {len(leaked)} leaked toasts in win.children()!"

    def test_clean_exit_purges_active_toasts(self, qapp, main_window):
        """clean_exit terminates and deletes all active toasts."""
        win = main_window
        ToastNotification(win, '通知 1', kind='info', duration=60000)
        ToastNotification(win, '通知 2', kind='success', duration=60000)
        assert len(win._active_toasts) >= 2

        win.clean_exit()
        qapp.processEvents()

        assert len(win._active_toasts) == 0
        leaked = [c for c in win.children() if isinstance(c, ToastNotification)]
        assert len(leaked) == 0


# ==============================================================================
# F12: Headless Animation Safety Controller
# ==============================================================================
class TestF12HeadlessAnimationSafetyController:
    """Verifies headless animation safety governor fast-forwards in offscreen mode."""

    def test_get_animation_duration_returns_zero_offscreen(self, monkeypatch):
        """Under QT_QPA_PLATFORM=offscreen, returns 0ms."""
        monkeypatch.setenv('QT_QPA_PLATFORM', 'offscreen')
        assert get_animation_duration(250) == 0
        assert get_animation_duration(3000) == 0

    def test_get_animation_duration_returns_default_when_desktop(self, monkeypatch):
        """In normal desktop runtime, returns requested default_ms."""
        monkeypatch.delenv('QT_QPA_PLATFORM', raising=False)
        monkeypatch.setattr(QGuiApplication, 'platformName', lambda: 'windows')
        assert get_animation_duration(250) == 250
        assert get_animation_duration(1000) == 1000

    def test_pulse_timer_governor_in_app(self, qapp, main_window):
        """app.py pulse animation timer respects get_animation_duration(250)."""
        win = main_window
        win._fill_table(
            win.room_table,
            rows=[{'校区': '花江', '组别': 'A组', '教室': '0101', '考勤状态': '待核对'}],
            columns=['校区', '组别', '教室', '考勤状态']
        )

        win._flash_room_row(0)
        assert getattr(win, '_pulse_timer', None) is not None
        # Under offscreen test runner, timer interval is 0ms, so events process quickly
        for _ in range(10):
            qapp.processEvents()
        assert win._pulse_timer is None
