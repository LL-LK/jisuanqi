# -*- coding: utf-8 -*-
"""系统托盘集成测试套件（Requirement R2）。

验证主窗口最小化至系统托盘、任务栏图标隐藏、双击与右键菜单恢复窗口、
彻底退出清理托盘图标（防 Ghost 图标残留）及状态保持等核心行为。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest
from PySide6.QtCore import QEvent, Qt
from PySide6.QtWidgets import QApplication, QSystemTrayIcon

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from app import APP_NAME, APP_VERSION, MainWindow


@pytest.fixture(scope='session')
def qapp():
    """PySide6 应用程序实例。"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv[:1])
    yield app


@pytest.fixture(autouse=True)
def mock_tray_available(monkeypatch):
    """确保在 CI/CD 或 pytest 全套件运行（设置了 QT_QPA_PLATFORM=offscreen）时，
    托盘检测判定为可用以完整覆盖托盘交互逻辑。
    """
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: True))


@pytest.fixture
def main_window(qapp):
    """每个测试独立创建并确保测试后彻底退出的 MainWindow 实例。"""
    win = MainWindow()
    yield win
    try:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
    except Exception:
        pass


def test_tray_initialization(qapp, main_window):
    """测试 1: 验证托盘图标、提示文本、右键菜单与多分辨率图标正常初始化。"""
    win = main_window

    # 1. 托盘图标实例与显示状态
    assert win.tray_icon is not None, "系统托盘图标对象应已初始化"
    assert win.tray_icon.isVisible() is True, "系统托盘图标应当处于可见状态"

    # 2. 图标非空且有效
    icon = win.tray_icon.icon()
    assert not icon.isNull(), "系统托盘图标对象不应为空"

    # 3. 提示信息格式
    tooltip = win.tray_icon.toolTip()
    assert APP_NAME in tooltip, f"Tooltip 中应包含应用名 {APP_NAME}"
    assert APP_VERSION in tooltip, f"Tooltip 中应包含版本号 {APP_VERSION}"

    # 4. 右键菜单结构与动作
    assert win.tray_menu is not None, "托盘右键菜单应当存在"
    actions = win.tray_menu.actions()
    action_texts = [a.text() for a in actions if not a.isSeparator()]
    assert '恢复显示' in action_texts, "托盘右键菜单应包含「恢复显示」动作"
    assert '彻底退出' in action_texts, "托盘右键菜单应包含「彻底退出」动作"

    # 5. 恢复显示应为加粗默认动作
    default_act = win.tray_menu.defaultAction()
    assert default_act is not None, "托盘菜单应设有默认动作"
    assert default_act.text() == '恢复显示', "默认动作应为「恢复显示」"
    assert default_act.font().bold() is True, "「恢复显示」动作字体应为粗体"


def test_native_minimize_keeps_taskbar_visibility(qapp, main_window):
    """测试 2a: 验证原生点击最小化时窗口保留在任务栏不强行隐藏（防止误入隐藏托盘模式）。"""
    win = main_window
    win.show()
    qapp.processEvents()
    assert win.isVisible() is True

    win.showMinimized()
    qapp.processEvents()
    qapp.processEvents()

    assert win.isMinimized() is True
    assert win.isVisible() is True, "原生最小化应保留在任务栏，不得擅自隐藏"


def test_minimize_hides_to_tray(qapp, main_window):
    """测试 2b: 验证主动点击「后台最小化运行」时主窗口隐藏至托盘，托盘图标保持可见。"""
    win = main_window
    win.show()
    qapp.processEvents()
    assert win.isVisible() is True, "窗口初始显示时应可见"

    # 触发显式隐藏至托盘
    win._minimize_to_tray()
    qapp.processEvents()
    qapp.processEvents()

    # 窗口应隐藏（从任务栏脱出），托盘依然可见
    assert win.isVisible() is False, "进入托盘模式后主窗口应从桌面与任务栏隐藏"
    assert win.tray_icon.isVisible() is True, "系统托盘图标应保持可见"


def test_tray_double_click_restores_window(qapp, main_window):
    """测试 3: 验证双击托盘图标能正确恢复主窗口至前台并取消最小化状态。"""
    win = main_window
    win.show()
    qapp.processEvents()

    # 先隐藏到托盘
    win._minimize_to_tray()
    qapp.processEvents()
    qapp.processEvents()
    assert win.isVisible() is False, "窗口前置条件必须为隐藏在托盘"

    # 模拟双击托盘图标
    win._on_tray_activated(QSystemTrayIcon.ActivationReason.DoubleClick)
    qapp.processEvents()

    assert win.isVisible() is True, "双击托盘图标后窗口应恢复可见"
    assert win.isMinimized() is False, "恢复后窗口不应继续处于最小化状态"


def test_tray_single_click_trigger_restores_window(qapp, main_window):
    """测试 3b: 验证单击（Trigger）托盘图标亦能唤起恢复主窗口。"""
    win = main_window
    win.show()
    qapp.processEvents()

    win._minimize_to_tray()
    qapp.processEvents()
    qapp.processEvents()
    assert win.isVisible() is False

    # 模拟单击托盘
    win._on_tray_activated(QSystemTrayIcon.ActivationReason.Trigger)
    qapp.processEvents()

    assert win.isVisible() is True
    assert win.isMinimized() is False


def test_tray_context_menu_restore(qapp, main_window):
    """测试 4: 验证右键菜单触发「恢复显示」能正确把窗口拉回前台。"""
    win = main_window
    win.show()
    qapp.processEvents()

    win._minimize_to_tray()
    qapp.processEvents()
    qapp.processEvents()
    assert win.isVisible() is False

    # 触发菜单中的「恢复显示」动作
    win.restore_action.trigger()
    qapp.processEvents()

    assert win.isVisible() is True
    assert win.isMinimized() is False


def test_clean_exit_prevents_ghost_icon(qapp):
    """测试 5: 验证「彻底退出」安全清理托盘图标，杜绝 Windows Ghost 托盘残留，且安全终止线程。"""
    win = MainWindow()
    win.show()
    qapp.processEvents()

    assert win.tray_icon is not None
    assert win.tray_icon.isVisible() is True

    # 触发彻底退出动作
    win.exit_action.trigger()
    qapp.processEvents()

    # 托盘图标应已被设为 None 或已 hide
    assert win.tray_icon is None or not win.tray_icon.isVisible(), "退出后托盘图标必须已被隐藏与释放"

    # 后台线程与动画计时器应已安全停止
    if getattr(win, 'thread', None):
        assert not win.thread.isRunning(), "计算工作线程必须已退出"
    if getattr(win, '_selfcheck_thread', None):
        assert not win._selfcheck_thread.isRunning(), "自测后台线程必须已退出"
    if getattr(win, '_pulse_timer', None):
        assert not win._pulse_timer.isActive(), "脉冲动画计时器必须已停止"

    win.deleteLater()
    qapp.processEvents()


def test_maximized_window_restores_to_maximized(qapp, main_window):
    """测试 6: 验证最大化窗口在最小化到托盘后，恢复显示时保持最大化状态而不是缩回默认尺寸。"""
    win = main_window
    win.showMaximized()
    qapp.processEvents()
    assert win.isMaximized() is True

    # 触发最小化到托盘
    win._minimize_to_tray()
    qapp.processEvents()
    qapp.processEvents()
    assert win.isVisible() is False

    # 恢复显示
    win.restore_window()
    qapp.processEvents()

    assert win.isVisible() is True
    assert win.isMinimized() is False
    assert win.isMaximized() is True, "原本最大化的窗口恢复后应当继续保持最大化"


def test_tray_unavailable_fallback(monkeypatch, qapp):
    """测试 7: 验证当系统无托盘服务可用时（如纯 Linux 容器无桌面），优雅降级不崩溃且不创建托盘图标。"""
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: False))
    win = MainWindow()
    assert win.tray_icon is None, "无托盘服务时应当降级为 None"
    win.show()
    qapp.processEvents()
    # 最小化时保持标准行为，不报错
    win.showMinimized()
    qapp.processEvents()
    win.clean_exit()
    win.deleteLater()
    qapp.processEvents()
