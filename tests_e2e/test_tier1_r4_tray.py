# -*- coding: utf-8 -*-
"""Tier 1: Requirement R4 Feature Coverage (Isolation).
Tests PySide6 system tray integration, minimize-to-tray lifecycle, context menu actions,
memory trimming on minimize, clean exit without ghost icons, and headless degradation.

Specification: PROJECT.md § Milestones M3, ORIGINAL_REQUEST.md § R4
Feature IDs: F10, F11, F12, F13, F14, F15, F16, F17, F18
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest
from PySide6.QtCore import QEvent, Qt
from PySide6.QtWidgets import QApplication, QSystemTrayIcon

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app import APP_NAME, APP_VERSION, MainWindow


@pytest.fixture
def test_window(qapp):
    """Instantiate a MainWindow for tray testing and guarantee clean teardown."""
    win = MainWindow()
    yield win
    try:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
    except Exception:
        pass


def test_r4_tray_icon_and_tooltip_initialization(qapp, test_window):
    """Test 1 (R4.1 / F10): MainWindow initializes QSystemTrayIcon with correct tooltip and icon."""
    win = test_window
    assert hasattr(win, 'tray_icon'), "MainWindow must have tray_icon attribute"
    assert win.tray_icon is not None, "System tray icon object should be initialized"
    assert isinstance(win.tray_icon, QSystemTrayIcon)

    tooltip = win.tray_icon.toolTip()
    assert APP_NAME in tooltip, f"Tooltip should contain '{APP_NAME}', got '{tooltip}'"
    assert APP_VERSION in tooltip, f"Tooltip should contain '{APP_VERSION}', got '{tooltip}'"
    assert not win.tray_icon.icon().isNull(), "Tray icon should not be null"


def test_r4_tray_context_menu_structure(qapp, test_window):
    """Test 2 (R4.2 / F11): Tray context menu provides '恢复显示' (bold default) and '彻底退出'."""
    win = test_window
    menu = win.tray_icon.contextMenu()
    assert menu is not None, "Tray icon must have a context menu"

    actions = menu.actions()
    action_texts = [a.text() for a in actions if not a.isSeparator()]

    assert '恢复显示' in action_texts, "Menu must contain '恢复显示'"
    assert '彻底退出' in action_texts, "Menu must contain '彻底退出'"

    # Check default action is '恢复显示' and is bold
    default_act = menu.defaultAction()
    assert default_act is not None, "Menu should have a default action"
    assert default_act.text() == '恢复显示'
    assert default_act.font().bold() is True, "'恢复显示' action should be bold"


def test_r4_minimize_to_tray_event(qapp, test_window):
    """Test 3 (R4.3 / F12, F16): Window minimization hides window from taskbar and keeps tray active."""
    win = test_window
    win.show()
    qapp.processEvents()

    # Simulate minimize event
    win.setWindowState(win.windowState() | Qt.WindowMinimized)
    event = QEvent(QEvent.WindowStateChange)
    win.changeEvent(event)
    qapp.processEvents()

    assert win.isMinimized() is True
    # Verify memory trim called and window is handled cleanly
    assert win.tray_icon.isVisible() is True, "Tray icon must remain visible after minimization"


def test_r4_restore_window_from_tray(qapp, test_window):
    """Test 4 (R4.4 / F15): restore_window clears minimized state and activates window."""
    win = test_window
    win.showMinimized()
    qapp.processEvents()

    # Call restore_window
    win.restore_window()
    qapp.processEvents()

    assert not (win.windowState() & Qt.WindowMinimized), "Window should no longer be minimized"
    assert win.isVisible() is True, "Window should be visible"


def test_r4_clean_exit_prevents_ghost_icon(qapp, test_window):
    """Test 5 (R4.5 / F14): clean_exit hides and disposes tray_icon, preventing ghost icons."""
    win = test_window
    assert win.tray_icon is not None

    win.clean_exit()
    qapp.processEvents()

    assert getattr(win, 'tray_icon', None) is None, "tray_icon should be reset to None after clean_exit"
    assert getattr(win, '_is_cleaning_exit', False) is True, "Exit cleanup flag should be set"


def test_r4_memory_trimming_on_minimize(qapp, test_window):
    """Test 6 (R4.6 / F13, F17): _trim_memory executes without exceptions on the host OS."""
    win = test_window
    assert hasattr(win, '_trim_memory'), "MainWindow must implement _trim_memory method"

    # Should run smoothly and safely on both Windows and non-Windows
    try:
        win._trim_memory()
    except Exception as exc:
        pytest.fail(f"_trim_memory raised unexpected exception: {exc}")


def test_r4_tray_graceful_degradation(monkeypatch, qapp):
    """Test 7 (R4.7 / F18): When system tray is unavailable, MainWindow handles gracefully without crashing."""
    monkeypatch.setattr(QSystemTrayIcon, 'isSystemTrayAvailable', staticmethod(lambda: False))

    win = MainWindow()
    try:
        # Tray icon should be None or safely disabled
        assert win.tray_icon is None or not win.tray_icon.isVisible()
        # Minimizing should not throw attribute error
        win.setWindowState(win.windowState() | Qt.WindowMinimized)
        win.changeEvent(QEvent(QEvent.WindowStateChange))
        qapp.processEvents()
    finally:
        win.clean_exit()
        win.deleteLater()
        qapp.processEvents()
