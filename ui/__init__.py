# -*- coding: utf-8 -*-
"""UI 表现层模块包。

解耦包含：
- theme: 设计令牌、QSS 样式表与字体引擎
- icons: 矢量图标生成器、动态 ActionIcon 与小熊头像
- widgets: 交互卡槽、状态指示器、横向工作流步骤卡、Toast通知
- dialogs: 业务弹窗（历史设置、学院下钻、词典编辑）
"""

from .theme import TOKENS, COLORS, APP_QSS, build_app_font, build_mono_font, font_is_ok
from .icons import AttendanceIcon, ActionIcon, BearPainter
from .widgets import (
    PathPreview,
    ToastNotification,
    CuteBearAvatar,
    MetricCard,
    DropSlot,
    ResultSlot,
    EmptyStepCard,
    EmptyStateCard,
    _GenericTableModel,
    _format_file_size,
    get_animation_duration,
)
from .dialogs import (
    StatusDialog,
    StudentDrilldownDialog,
    DatabaseHistoryDialog,
    _DictModel,
    _DictTableWidget,
    _run_frozen_inprocess_selfcheck,
)

__all__ = [
    "TOKENS",
    "COLORS",
    "APP_QSS",
    "build_app_font",
    "build_mono_font",
    "font_is_ok",
    "AttendanceIcon",
    "ActionIcon",
    "BearPainter",
    "PathPreview",
    "ToastNotification",
    "CuteBearAvatar",
    "MetricCard",
    "DropSlot",
    "ResultSlot",
    "EmptyStepCard",
    "EmptyStateCard",
    "_GenericTableModel",
    "_format_file_size",
    "get_animation_duration",
    "StatusDialog",
    "StudentDrilldownDialog",
    "DatabaseHistoryDialog",
    "_DictModel",
    "_DictTableWidget",
    "_run_frozen_inprocess_selfcheck",
]
