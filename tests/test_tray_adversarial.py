# -*- coding: utf-8 -*-
"""Milestone M1 系统托盘对抗性与压力测试套件（Adversarial Challenge Suite）。

覆盖高频最小化/恢复风暴、多状态几何记忆保真度、可见状态下托盘点击容错、
clean_exit 幂等性与重入、以及后台计算线程活跃时的强制退出安全性。
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import pytest
from PySide6.QtCore import QEvent, QObject, QThread, Qt, Signal
from PySide6.QtGui import QCloseEvent
from PySide6.QtWidgets import QApplication, QSystemTrayIcon

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from app import APP_NAME, APP_VERSION, MainWindow, Worker


@pytest.fixture(scope='session')
def qapp():
    """PySide6 应用程序实例。"""
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv[:1])
    yield app


@pytest.fixture(autouse=True)
def mock_tray_available(monkeypatch):
    """确保测试环境下托盘检测可用。"""
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: True))


@pytest.fixture
def window(qapp):
    """每个测试独立创建并确保测试后安全退出的 MainWindow 实例。"""
    win = MainWindow()
    yield win
    try:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
    except Exception:
        pass


# =====================================================================
# 1. 高频最小化与恢复压力测试 (Rapid Minimize & Restore Cycles)
# =====================================================================

def test_rapid_minimize_restore_cycles(qapp, window):
    """压力场景 1a: 50 次连续高速最小化与恢复循环，验证事件循环队列无堆积、窗口状态无偏离。"""
    win = window
    win.show()
    qapp.processEvents()
    assert win.isVisible() is True

    for i in range(50):
        # 触发最小化
        win.showMinimized()
        qapp.processEvents()
        assert win.isVisible() is False, f"迭代 {i}: 最小化后窗口必须隐藏"

        # 触发恢复
        win.restore_window()
        qapp.processEvents()
        assert win.isVisible() is True, f"迭代 {i}: 恢复后窗口必须可见"
        assert win.isMinimized() is False, f"迭代 {i}: 恢复后窗口不得处于最小化状态"


def test_sub_tick_minimize_restore_race_condition(qapp, window):
    """压力场景 1b: 探查子刻度（零等待）极速最小化紧接着恢复时，QTimer.singleShot 的行为表现。

    缺陷复现用例：若 changeEvent 中使用无条件的 singleShot(0, self.hide)，
    在恢复动作先于定时器触发执行后，定时器依然强行 hide 窗口，导致窗口处于 isMinimized=False 但 isVisible=False 的幽灵隐藏状态。
    """
    win = window
    win.show()
    qapp.processEvents()

    # 模拟在同一个事件周期内，最小化后立刻触发恢复
    win.showMinimized()
    win.restore_window()
    qapp.processEvents()

    # 窗口在恢复后理应保持可见
    assert win.isVisible() is True, (
        "竞争缺陷复现：最小化后紧接恢复动作，但在随后事件循环中被无条件的 singleShot(0, self.hide) 再次隐藏！"
    )
    assert win.isMinimized() is False, "恢复后窗口不得处于最小化状态"


# =====================================================================
# 2. 多窗口状态（Maximized / Normal / FullScreen）转换保真度测试
# =====================================================================

def test_window_state_transitions(qapp, window):
    """场景 2: 验证 Normal, Maximized, Fullscreen 及状态来回切换后的最小化与恢复保真度。"""
    win = window

    # 2a. Normal 状态最小化与恢复
    win.showNormal()
    qapp.processEvents()
    assert win.isMaximized() is False
    assert win.isFullScreen() is False

    win.showMinimized()
    qapp.processEvents()
    assert win.isVisible() is False
    win.restore_window()
    qapp.processEvents()
    assert win.isVisible() is True
    assert win.isMaximized() is False, "Normal 窗口恢复后不得错误变为最大化"
    assert win.isFullScreen() is False

    # 2b. Maximized 状态最小化与恢复
    win.showMaximized()
    qapp.processEvents()
    assert win.isMaximized() is True

    win.showMinimized()
    qapp.processEvents()
    assert win.isVisible() is False
    win.restore_window()
    qapp.processEvents()
    assert win.isVisible() is True
    assert win.isMaximized() is True, "原本最大化的窗口恢复后必须保持最大化"

    # 2c. Maximized -> 恢复 Normal -> 最小化 -> 恢复，验证 _was_maximized 正确重置
    win.showNormal()
    qapp.processEvents()
    assert win.isMaximized() is False

    win.showMinimized()
    qapp.processEvents()
    assert win.isVisible() is False
    win.restore_window()
    qapp.processEvents()
    assert win.isVisible() is True
    assert win.isMaximized() is False, "取消最大化后的窗口最小化恢复后不应重新变为最大化"

    # 2d. FullScreen 状态最小化与恢复
    win.showFullScreen()
    qapp.processEvents()
    assert win.isFullScreen() is True

    win.showMinimized()
    qapp.processEvents()
    assert win.isVisible() is False
    win.restore_window()
    qapp.processEvents()
    assert win.isVisible() is True
    assert win.isMinimized() is False


# =====================================================================
# 3. 窗口已可见时的托盘交互容错测试 (Visible vs Hidden Restorations)
# =====================================================================

def test_tray_activation_while_already_visible(qapp, window):
    """场景 3: 窗口在桌面已处于可见状态时，接收托盘双击、单击或菜单恢复，必须平滑保持可见无闪烁或状态颠倒。"""
    win = window
    win.show()
    qapp.processEvents()
    assert win.isVisible() is True

    # 3a. 在可见状态下模拟 10 次双击
    for _ in range(10):
        win._on_tray_activated(QSystemTrayIcon.ActivationReason.DoubleClick)
        qapp.processEvents()
        assert win.isVisible() is True
        assert win.isMinimized() is False

    # 3b. 在可见状态下模拟 10 次单击 (Trigger)
    for _ in range(10):
        win._on_tray_activated(QSystemTrayIcon.ActivationReason.Trigger)
        qapp.processEvents()
        assert win.isVisible() is True
        assert win.isMinimized() is False

    # 3c. 在可见状态下触发菜单项「恢复显示」
    win.restore_action.trigger()
    qapp.processEvents()
    assert win.isVisible() is True
    assert win.isMinimized() is False

    # 3d. 最大化状态下接收托盘点击，验证不坍缩
    win.showMaximized()
    qapp.processEvents()
    assert win.isMaximized() is True
    win._on_tray_activated(QSystemTrayIcon.ActivationReason.DoubleClick)
    qapp.processEvents()
    assert win.isVisible() is True
    assert win.isMaximized() is True, "可见的最大化窗口被托盘激活后应继续保持最大化"


# =====================================================================
# 4. clean_exit 幂等性与重入测试 (Idempotency & Re-entrancy)
# =====================================================================

def test_clean_exit_idempotency_multiple_calls(qapp, window):
    """场景 4: 连续调用 clean_exit 20 次，混合 closeEvent，验证无二次释放或 None 引用崩溃。"""
    win = window
    win.show()
    qapp.processEvents()

    # 第 1 次正常退出
    win.clean_exit()
    assert win.tray_icon is None
    assert getattr(win, '_is_cleaning_exit', False) is True

    # 随后连续调用 19 次
    for _ in range(19):
        win.clean_exit()

    # 模拟 QCloseEvent 再次进入
    ev = QCloseEvent()
    win.closeEvent(ev)
    assert ev.isAccepted() is True

    # 模拟菜单退出动作再次被触发
    if hasattr(win, 'exit_action') and win.exit_action is not None:
        try:
            win.exit_action.trigger()
        except Exception as e:
            pytest.fail(f"再次触发 exit_action 抛出异常: {e}")

    qapp.processEvents()


# =====================================================================
# 5. 后台计算工作线程活跃时的退出安全性 (Clean Exit with Active Workers)
# =====================================================================

def test_clean_exit_with_active_selfcheck_thread(qapp, window):
    """场景 5a: 自检后台线程（QThread）正在持续执行时调用 clean_exit，验证其被安全终止不挂起。"""
    win = window
    win.show()
    qapp.processEvents()

    class InfiniteSelfCheckThread(QThread):
        def run(self):
            while True:
                time.sleep(0.01)

    win._selfcheck_thread = InfiniteSelfCheckThread(win)
    win._selfcheck_thread.start()
    assert win._selfcheck_thread.isRunning() is True

    win.clean_exit()
    assert not win._selfcheck_thread.isRunning(), "自检线程必须已被 terminate 并退出"


def test_clean_exit_with_active_calculation_thread(monkeypatch, qapp, window):
    """场景 5b: 主计算后台线程（QThread + QObject Worker）正在持续运行时调用 clean_exit，

    验证 Worker 在 run() 中显式开启 QThread.setTerminationEnabled(True) 后，
    clean_exit 中的 self.thread.terminate() 能安全终止运行中的工作线程，
    避免工作线程持续存活、超时、僵尸进程与 Python 运行时退出异常 (Exit Code 1)。
    """
    win = window
    win.show()
    qapp.processEvents()

    monkeypatch.setattr('app.run', lambda cfg: time.sleep(10))

    win.thread = Worker(None, win)
    win.worker = win.thread
    win.thread.start()

    time.sleep(0.05)
    assert win.thread.isRunning() is True

    t0 = time.time()
    win.clean_exit()
    elapsed = time.time() - t0

    assert not win.thread.isRunning(), (
        f"计算线程未能退出（仍处于 isRunning=True）！耗时 {elapsed:.2f}s。"
        "根因：Worker.run 未开启 setTerminationEnabled(True)，导致 terminate() 被 Qt 静默忽略。"
    )


# =====================================================================
# 6. 无系统托盘支持环境下的全流程退化容错测试
# =====================================================================

def test_tray_unavailable_adversarial_suite(monkeypatch, qapp):
    """场景 6: 模拟系统完全缺乏通知区域支持（isSystemTrayAvailable 返回 False）下的高频状态交互与退出。"""
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: False))

    win = MainWindow()
    assert win.tray_icon is None

    win.show()
    qapp.processEvents()

    # 高频最小化与恢复
    for _ in range(5):
        win.showMinimized()
        qapp.processEvents()
        win.restore_window()
        qapp.processEvents()

    # 重复 clean_exit
    win.clean_exit()
    win.clean_exit()
    win.deleteLater()
    qapp.processEvents()
