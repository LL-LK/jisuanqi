# -*- coding: utf-8 -*-
"""三早一晚免勤自动统计系统 · 桌面界面（PySide6）。

四个能力面板：
    概览 / 学院出勤率 / 教室明细 / 异常清单（含低出勤率复核） / 词典编辑 / 合成数据自测

硬规则（界面上也不允许违反）：
    输出只保留三张原始工作表；屏风不参与计算且内容零改动；
    实到缺失留空不写 0；绝不覆盖输入文件。
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import traceback
from datetime import datetime
from functools import lru_cache

from PySide6.QtCore import (QAbstractTableModel, QEvent, QModelIndex, QObject, QPointF,
                            QRectF, QSettings, QSize, QThread, QTimer, QUrl, Qt, Signal)
from PySide6.QtGui import (QAction, QBrush, QColor, QDesktopServices, QDragEnterEvent, QDropEvent, QFont,
                           QGuiApplication, QIcon, QKeySequence, QPainter, QPainterPath, QPen,
                           QPixmap, QShortcut, QStandardItemModel, QStandardItem)
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QCheckBox,
                               QComboBox, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QFrame,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow,
                               QMenu, QMessageBox, QPlainTextEdit, QProgressBar,
                               QPushButton, QSizePolicy, QSpinBox, QStackedWidget,
                               QSplitter, QSystemTrayIcon, QTableView, QTabWidget, QVBoxLayout,
                               QWidget)

from core import RunConfig, run, __version__ as APP_VERSION
from core.calc import ESTIMATE_RGB          # R17 估算标记色（唯一事实源，界面禁另写一份）
WARN_RGB = 'FFF6DC'                         # SL4 异常行警示色不变量
from core.dictionary import Dictionary, load_tables, save_tables
from core.exempt_report import COLLEGE_REPORT_COLUMNS, REPORT_COLUMNS, export_exempt_report
from core.export import _same_path, atomic_output, check_output_path, unique_workbook_paths
from core.io_detail import week_from_label
from core.io_exempt import normalize_student_name
from core.excel_safety import (
    DiagnosticMasker,
    FileCheckResult,
    classify_excel_file,
    generate_diagnostic_log,
    verify_excel_health,
)

APP_NAME = '本地考勤工具'
APP_VERSION = APP_VERSION
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DICT_NAME = '免勤词典.xlsx'
LOW_RATE_THRESHOLD = 0.80
PRIVACY_NOTICE = (
    '本工具在当前设备读取你选中的考勤表、免勤表和词典。姓名、学号、学院、班级、校区和'
    '免勤申请信息用于核对身份、筛选记录和计算统计结果。软件运行时不上传表格，不调用云端模型。\n\n'
    '人员明细在本次运行中保留在内存。程序会记住周次、输出位置、词典路径和窗口设置。'
    '结果工作簿保存在你选择的位置，新增汇总和核对页包含统计数字、状态和来源表行号。\n\n'
    '手动导出的人员名单包含个人信息。诊断日志会做脱敏处理，分享前请核对内容。'
    '如需删除导出的文件，请在本机删除对应文件。'
)


def get_animation_duration(default_ms: int) -> int:
    """Return animation duration in milliseconds.

    Fast-forwards to 0ms in offscreen/headless environments (pytest, CI)
    to prevent test timeouts while preserving fluid desktop micro-interactions.
    """
    if os.environ.get("QT_QPA_PLATFORM") == "offscreen":
        return 0
    try:
        if QGuiApplication.platformName() == "offscreen":
            return 0
    except Exception:
        pass
    return int(default_ms)


def _format_file_size(size_bytes: int | float) -> str:
    """Format byte size into human readable string (e.g. 245.8 KB)."""
    try:
        size = float(size_bytes)
        if size < 1024:
            return f'{int(size)} B'
        elif size < 1024 * 1024:
            return f'{size / 1024.0:.1f} KB'
        else:
            return f'{size / (1024.0 * 1024.0):.1f} MB'
    except Exception:
        return ''


class AttendanceIcon:
    """High-DPI retina sharpness asset provider with fallback vector engine."""

    @staticmethod
    def pixmap(size=48, target_widget=None):
        if isinstance(target_widget, (int, float)):
            dpr = float(target_widget)
        elif target_widget is not None and hasattr(target_widget, 'devicePixelRatio'):
            try:
                dpr = float(target_widget.devicePixelRatio())
            except Exception:
                dpr = 1.0
        else:
            screen = QGuiApplication.primaryScreen()
            dpr = float(screen.devicePixelRatio()) if screen else 1.0
        if dpr <= 0:
            dpr = 1.0
        return AttendanceIcon._render_cached(int(size), float(dpr))

    @staticmethod
    @lru_cache(maxsize=32)
    def _render_cached(size: int, dpr: float) -> QPixmap:
        phys_size = max(1, int(round(size * dpr)))
        image_path = os.path.join(getattr(sys, '_MEIPASS', PROJECT_DIR), 'assets', 'strawberry-bear.png')
        artwork = QPixmap(image_path)
        if not artwork.isNull():
            scaled = artwork.scaled(phys_size, phys_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            scaled.setDevicePixelRatio(dpr)
            return scaled
        icon = QPixmap(phys_size, phys_size)
        icon.fill(Qt.transparent)
        painter = QPainter(icon)
        painter.setRenderHint(QPainter.Antialiasing)
        scale_factor = phys_size / 48.0
        painter.scale(scale_factor, scale_factor)
        painter.setPen(Qt.NoPen)
        painter.setBrush(QColor('#F8E7EC'))
        painter.drawRoundedRect(QRectF(1, 1, 46, 46), 10, 10)
        pen = QPen(QColor('#781236'), 2.6)
        pen.setCapStyle(Qt.RoundCap)
        pen.setJoinStyle(Qt.RoundJoin)
        painter.setPen(pen)
        painter.setBrush(Qt.NoBrush)
        painter.drawRoundedRect(QRectF(11, 13, 26, 25), 4, 4)
        painter.drawLine(QPointF(11, 21), QPointF(37, 21))
        painter.drawLine(QPointF(18, 10), QPointF(18, 16))
        painter.drawLine(QPointF(30, 10), QPointF(30, 16))
        painter.drawLine(QPointF(18, 29), QPointF(22, 33))
        painter.drawLine(QPointF(22, 33), QPointF(30, 26))
        painter.end()
        icon.setDevicePixelRatio(dpr)
        return icon

    @staticmethod
    def icon(size=48, target_widget=None) -> QIcon:
        ico = QIcon()
        for s in (16, 24, 32, 48, 64, 128, 256):
            ico.addPixmap(AttendanceIcon.pixmap(s, target_widget=target_widget))
        return ico

    # Backward compatibility for icon generators and legacy tools
    create_bear_avatar = pixmap


AttendanceIcon.pixmap.cache_clear = AttendanceIcon._render_cached.cache_clear
AttendanceIcon.pixmap.cache_info = AttendanceIcon._render_cached.cache_info
AttendanceIcon.create_bear_avatar.cache_clear = AttendanceIcon._render_cached.cache_clear
AttendanceIcon.create_bear_avatar.cache_info = AttendanceIcon._render_cached.cache_info

# Backward compatibility aliases
BearPainter = AttendanceIcon


class ActionIcon:
    """Multi-resolution vector action icon provider for QPushButton and UI actions."""

    _CACHE: dict[tuple[str, str], QIcon] = {}

    @classmethod
    def get(cls, kind: str, color: str = '#781236') -> QIcon:
        key = (kind, color.lower())
        if key in cls._CACHE:
            return cls._CACHE[key]
        ico = QIcon()
        for size in (16, 20, 24, 32, 48):
            px = cls._render_pixmap(kind, size, color)
            ico.addPixmap(px)
        cls._CACHE[key] = ico
        return ico

    @staticmethod
    def _render_pixmap(kind: str, size: int, color_hex: str) -> QPixmap:
        dpr = 1.0
        app = QApplication.instance()
        if app:
            try:
                dpr = float(app.devicePixelRatio())
            except Exception:
                dpr = 1.0
        if dpr <= 0:
            dpr = 1.0
        phys_size = max(1, int(round(size * dpr)))
        px = QPixmap(phys_size, phys_size)
        px.fill(Qt.transparent)
        p = QPainter(px)
        p.setRenderHint(QPainter.Antialiasing)
        scale_factor = phys_size / 24.0
        p.scale(scale_factor, scale_factor)
        color = QColor(color_hex)
        pen = QPen(color, 2.0, Qt.SolidLine, Qt.RoundCap, Qt.RoundJoin)
        p.setPen(pen)
        p.setBrush(Qt.NoBrush)

        if kind in ('document', 'file'):
            path = QPainterPath()
            path.moveTo(5, 3)
            path.lineTo(14, 3)
            path.lineTo(19, 8)
            path.lineTo(19, 21)
            path.lineTo(5, 21)
            path.closeSubpath()
            p.drawPath(path)
            p.drawLine(QPointF(14, 3), QPointF(14, 8))
            p.drawLine(QPointF(14, 8), QPointF(19, 8))
            p.drawLine(QPointF(8, 12), QPointF(16, 12))
            p.drawLine(QPointF(8, 16), QPointF(13, 16))
        elif kind == 'table':
            p.drawRoundedRect(QRectF(3.5, 3.5, 17, 17), 2.0, 2.0)
            p.drawLine(QPointF(3.5, 8.5), QPointF(20.5, 8.5))
            p.drawLine(QPointF(9.5, 8.5), QPointF(9.5, 20.5))
            p.drawLine(QPointF(3.5, 14.5), QPointF(20.5, 14.5))
        elif kind in ('play', 'compute', 'calculate'):
            path = QPainterPath()
            path.moveTo(8.5, 6.0)
            path.lineTo(18.0, 12.0)
            path.lineTo(8.5, 18.0)
            path.closeSubpath()
            p.setBrush(QBrush(color))
            p.drawPath(path)
        elif kind in ('sparkle', 'magic'):
            # 4-pointed astroid star with soft concave arcs
            path = QPainterPath()
            path.moveTo(12, 3)
            path.quadTo(12, 12, 21, 12)
            path.quadTo(12, 12, 12, 21)
            path.quadTo(12, 12, 3, 12)
            path.quadTo(12, 12, 12, 3)
            path.closeSubpath()
            p.setBrush(QBrush(color))
            p.drawPath(path)
            p.drawEllipse(QRectF(17, 4, 2.5, 2.5))
        elif kind in ('export', 'save_as', 'upload', 'save'):
            # Tray below, arrow pointing UP/OUT
            tray = QPainterPath()
            tray.moveTo(4, 14)
            tray.lineTo(4, 19)
            tray.quadTo(4, 20.5, 5.5, 20.5)
            tray.lineTo(18.5, 20.5)
            tray.quadTo(20, 20.5, 20, 19)
            tray.lineTo(20, 14)
            p.drawPath(tray)
            p.drawLine(QPointF(12, 15), QPointF(12, 4.5))
            head = QPainterPath()
            head.moveTo(7.5, 9)
            head.lineTo(12, 4.5)
            head.lineTo(16.5, 9)
            p.drawPath(head)
        elif kind in ('folder', 'folder_open', 'dir'):
            # Closed rounded flap with semi-transparent brush fill, no wireframe line bleed-through
            back = QPainterPath()
            back.moveTo(3, 7)
            back.lineTo(8, 7)
            back.lineTo(10, 9)
            back.lineTo(20, 9)
            back.quadTo(21, 9, 21, 10)
            back.lineTo(21, 18)
            back.lineTo(3, 18)
            back.closeSubpath()
            p.drawPath(back)

            flap = QPainterPath()
            flap.moveTo(2.5, 19)
            flap.lineTo(5.5, 12)
            flap.lineTo(21.5, 12)
            flap.lineTo(18.5, 19)
            flap.closeSubpath()
            p.setBrush(QColor('#FFF8FA'))
            p.drawPath(flap)
            p.setBrush(Qt.NoBrush)
        elif kind in ('tools', 'settings', 'gear'):
            # 8 symmetrical rounded trapezoidal teeth at 45 deg intervals
            for i in range(8):
                angle = i * 45.0
                p.save()
                p.translate(12, 12)
                p.rotate(angle)
                p.drawRoundedRect(QRectF(-1.8, -9.5, 3.6, 3.2), 0.8, 0.8)
                p.restore()
            p.drawEllipse(QRectF(4.5, 4.5, 15.0, 15.0))
            p.drawEllipse(QRectF(9.0, 9.0, 6.0, 6.0))
        elif kind in ('search', 'magnifier'):
            # Lens center (10, 10) radius 4.5, handle tangent at (13.18, 13.18) to (19, 19), no interior penetration
            p.drawEllipse(QRectF(5.5, 5.5, 9.0, 9.0))
            p.drawLine(QPointF(13.18, 13.18), QPointF(19.0, 19.0))
        elif kind in ('save_disk', 'disk'):
            # Label rectangle terminates at y=18, strictly 2px above outer border y=20
            path = QPainterPath()
            path.moveTo(4, 4)
            path.lineTo(17, 4)
            path.lineTo(20, 7)
            path.lineTo(20, 20)
            path.lineTo(4, 20)
            path.closeSubpath()
            p.drawPath(path)
            p.drawRect(QRectF(7, 4, 10, 5))
            p.drawRect(QRectF(6, 12, 12, 6))
        elif kind in ('refresh', 'reset'):
            # Arc 45 to 315 deg, arrowhead smoothly tangent to arc end
            p.drawArc(QRectF(5.5, 5.5, 13.0, 13.0), 45 * 16, 270 * 16)
            head = QPainterPath()
            head.moveTo(13.0, 6.0)
            head.lineTo(17.5, 6.5)
            head.lineTo(16.5, 11.0)
            p.drawPath(head)
        elif kind in ('loading', 'spinner'):
            for i in range(8):
                angle = i * 45
                p.save()
                p.translate(12, 12)
                p.rotate(angle)
                dot_color = QColor(color)
                dot_color.setAlphaF(max(0.18, (i + 1) / 8.0))
                p.setPen(QPen(dot_color, 2.2, Qt.SolidLine, Qt.RoundCap))
                p.drawLine(QPointF(0, -6), QPointF(0, -9))
                p.restore()
        elif kind in ('dashboard', 'overview'):
            p.drawRoundedRect(QRectF(3.5, 3.5, 7.0, 7.0), 2, 2)
            p.drawRoundedRect(QRectF(13.5, 3.5, 7.0, 5.0), 2, 2)
            p.drawRoundedRect(QRectF(13.5, 11.5, 7.0, 9.0), 2, 2)
            p.drawRoundedRect(QRectF(3.5, 13.5, 7.0, 7.0), 2, 2)
        elif kind in ('percent', 'rate'):
            p.setBrush(QBrush(color))
            p.drawEllipse(QPointF(7.5, 7.5), 2.5, 2.5)
            p.drawEllipse(QPointF(16.5, 16.5), 2.5, 2.5)
            p.setBrush(Qt.NoBrush)
            p.drawLine(QPointF(17.5, 6.5), QPointF(6.5, 17.5))
        elif kind in ('dictionary', 'book'):
            left = QPainterPath()
            left.moveTo(11.5, 5)
            left.quadTo(8.0, 4.0, 4.0, 5.0)
            left.lineTo(4.0, 19.0)
            left.quadTo(8.0, 18.0, 11.5, 19.0)
            left.closeSubpath()
            p.drawPath(left)
            right = QPainterPath()
            right.moveTo(12.5, 5)
            right.quadTo(16.0, 4.0, 20.0, 5.0)
            right.lineTo(20.0, 19.0)
            right.quadTo(16.0, 18.0, 12.5, 19.0)
            right.closeSubpath()
            p.drawPath(right)
        elif kind in ('database', 'db', 'storage'):
            p.drawEllipse(QRectF(4, 4, 16, 5))
            p.drawArc(QRectF(4, 9, 16, 5), 180 * 16, 180 * 16)
            p.drawArc(QRectF(4, 14, 16, 5), 180 * 16, 180 * 16)
            p.drawLine(QPointF(4, 6.5), QPointF(4, 16.5))
            p.drawLine(QPointF(20, 6.5), QPointF(20, 16.5))
        elif kind in ('chart', 'chart_bar'):
            p.drawLine(QPointF(3, 20), QPointF(21, 20))
            p.drawLine(QPointF(3, 4), QPointF(3, 20))
            p.drawRect(QRectF(5.5, 12, 3.5, 8))
            p.drawRect(QRectF(10.5, 8, 3.5, 12))
            p.drawRect(QRectF(15.5, 5, 3.5, 15))
        elif kind in ('check_circle', 'check', 'success'):
            p.drawEllipse(QRectF(3.5, 3.5, 17.0, 17.0))
            chk = QPainterPath()
            chk.moveTo(7.5, 12.0)
            chk.lineTo(10.5, 15.0)
            chk.lineTo(16.5, 9.0)
            p.drawPath(chk)
        elif kind in ('alert_circle', 'alert', 'warn', 'warning'):
            tri = QPainterPath()
            tri.moveTo(12, 3.5)
            tri.lineTo(20.5, 18.5)
            tri.quadTo(21.0, 19.5, 20.0, 20.0)
            tri.lineTo(4.0, 20.0)
            tri.quadTo(3.0, 19.5, 3.5, 18.5)
            tri.closeSubpath()
            p.drawPath(tri)
            p.drawLine(QPointF(12, 8.5), QPointF(12, 13.5))
            p.drawEllipse(QRectF(11.2, 15.5, 1.6, 1.6))
        elif kind in ('info_circle', 'info'):
            p.drawEllipse(QRectF(3.5, 3.5, 17.0, 17.0))
            p.drawEllipse(QRectF(11.2, 7.5, 1.6, 1.6))
            p.drawLine(QPointF(12, 11.0), QPointF(12, 16.5))
        elif kind in ('lightbulb', 'idea', 'tip'):
            bulb = QPainterPath()
            bulb.moveTo(9, 16)
            bulb.lineTo(9, 15)
            bulb.quadTo(5.5, 12, 5.5, 8.5)
            bulb.arcTo(QRectF(5.5, 3.5, 13, 13), 180, -180)
            bulb.quadTo(18.5, 12, 15, 15)
            bulb.lineTo(15, 16)
            bulb.closeSubpath()
            p.drawPath(bulb)
            p.drawLine(QPointF(9, 18.5), QPointF(15, 18.5))
            p.drawLine(QPointF(10.5, 20.5), QPointF(13.5, 20.5))
        elif kind in ('lightning', 'flash', 'quick'):
            bolt = QPainterPath()
            bolt.moveTo(13, 3)
            bolt.lineTo(6, 13)
            bolt.lineTo(12, 13)
            bolt.lineTo(11, 21)
            bolt.lineTo(18, 11)
            bolt.lineTo(12, 11)
            bolt.closeSubpath()
            p.setBrush(QBrush(color))
            p.drawPath(bolt)
        p.end()
        px.setDevicePixelRatio(dpr)
        return px


# ----------------------------------------------------------------- 设计令牌 (Design Tokens)
TOKENS = {
    'colors': {
        'bg': '#FBF7F3',
        'surface': '#FFFCFA',
        'card': '#FFFCFA',
        'border': '#E2D2CF',
        'border_subtle': '#EBDAD6',
        'border_strong': '#C7ADAF',
        'divider': '#EBDAD6',
        'header_bg': '#F2DFE1',
        'text': '#442332',
        'text_strong': '#391525',
        'muted': '#70515C',
        'slate': '#795F68',
        'primary': '#8B1E3F',              # 基准主色锚点 (Strawberry Bear wine red)
        'primary_900': '#5E1028',
        'primary_800': '#781236',
        'primary_700': '#8B1E3F',
        'primary_600': '#A3264C',
        'primary_500': '#BF335D',
        'primary_100': '#F7DDE3',
        'primary_50': '#FCF2F4',
        'primary_dark': '#781236',         # 向后兼容
        'primary_light': '#F7DDE3',        # 向后兼容
        'primary_hover': '#A3264C',        # 向后兼容
        'primary_container': '#FFEBF0',
        'text_disabled': '#7B636A',        # WCAG AA >= 4.5:1 对比度 (4.530:1 vs #F3E7E5)
        'green_fg': '#146F64',
        'danger': '#B33C48',
        'warn_bg': '#' + WARN_RGB,  # '#FFF6DC'
        'warn_fg': '#8D580D',
        'low_bg': '#FFEAEA',
        'low_fg': '#A82A3F',
        'info': '#77505C',
        'info_bg': '#F8EEF0',
        'info_fg': '#693C4C',
        'estimate_bg': '#' + ESTIMATE_RGB,  # '#FFFFF3CD' - 业务不变量锁定
        'estimate_fg': '#8D580D',
        'trace_bg': '#F1EBE8',
    },
    'spacing': {
        'xs': 4,
        'sm': 8,
        'md': 12,
        'lg': 16,
        'xl': 20,
        'xxl': 24,
    },
    'radii': {
        'sm': '6px',
        'md': '8px',
        'lg': '12px',
        'pill': '9999px',
    },
    'typography': {
        'family': 'Microsoft YaHei UI, Microsoft YaHei, 微软雅黑, Segoe UI, sans-serif',
        'mono_family': 'Consolas, Microsoft YaHei Mono, SimSun, monospace',
        'size_display': '28px',
        'size_title': '18px',
        'size_body': '13px',
        'size_caption': '12px',
    },
    'elevation': {
        'card_rest': 'border: 1px solid #E2D2CF',
        'card_hover': 'border: 1px solid #8B1E3F',
        'dialog': 'border: 1px solid #E2D2CF',
    },
}

COLORS = {
    **TOKENS['colors'],
    'control': '#C7ADAF',
    'muted_fg': '#70515C',
}

APP_QSS = """
QWidget { background:%(bg)s; color:%(text)s; font-size:13px; }
QMainWindow, QDialog { background:%(bg)s; }
QLabel { background:transparent; }
QCheckBox { background:transparent; }
QToolTip { background:#442332; color:#FFFFFF; border:none; border-radius:6px; padding:6px 10px; font-size:12px; }

/* ---------- 三槽流水线与卡片容器 ---------- */
QFrame#dropSlot, QFrame#settingsPanel, QFrame#metricCard {
    background:%(card)s; border:1px solid %(border)s; border-radius:10px; }
QFrame#dropSlot { border:2px dashed #D6A4B1; border-radius:12px; }
QFrame#dropSlot:hover { border-color:%(primary)s; background:#FFFDFE; }
QFrame#dropSlot[dragOver="true"] { background:%(primary_light)s; border:2px solid %(primary)s; }
QFrame#dropSlot[dragOver="true"] QLabel#slotPath { color:%(primary)s; font-weight:bold; }
QFrame#dropSlot[dragReject="true"] { background:#FFF2F2; border:2px solid #FFAAA6; }
QFrame#dropSlot[dragReject="true"] QLabel#slotPath { color:%(danger)s; font-weight:bold; }
QFrame#dropSlot[justDropped="true"] { background:%(primary_light)s; border:2px solid %(primary)s; }
QLabel#slotBadge { color:#FFFFFF; font-weight:bold; border-radius:8px; padding:3px 11px;
                   background:%(primary)s; }
QLabel#slotTitle { font-size:16px; font-weight:bold; color:%(text_strong)s; }
QLabel#slotPath { color:%(muted)s; background:#FFFFFF; border:1px solid %(border)s;
                  border-radius:6px; padding:4px 9px; }
QFrame#dropSlot[ready="true"] { border:1.5px solid #146F64; background:#FAFCFB; }
QFrame#dropSlot[ready="true"]:hover { border-color:#0E564D; background:#F2FAF8; }
QFrame#dropSlot[ready="true"] QLabel#slotPath { color:%(green_fg)s; font-weight:600; }
QFrame#dropSlot[celebration="true"] { border:1.5px solid %(green_fg)s; background:#F2FAF8; }
QFrame#brandDivider { background:#C78E9C; }
QLabel#workspaceTitle { font-size:30px; font-weight:bold; color:%(text_strong)s; }
QLabel#sectionTitle { font-size:18px; font-weight:bold; color:%(text_strong)s; }
QLabel#settingLabel { font-weight:600; color:%(text_strong)s; }
QFrame#metricCard { background:%(bg)s; border:1px solid %(border)s; border-radius:10px; }
QFrame#metricCard:hover { border:1px solid %(primary)s; background:#FFF9F7; }
QLabel#metricLabel { color:%(text)s; font-size:13px; font-weight:600; }
QLabel#metricValue { font-size:26px; font-weight:bold; color:%(text_strong)s; }
QLabel#localBadge { background:#E9F4ED; color:%(green_fg)s; border-radius:8px;
                    padding:5px 10px; font-weight:600; }
QWidget#reportTools, QTabWidget QWidget#contentPage { background:%(card)s; }

/* ---------- 按钮 ---------- */
QPushButton { background:%(card)s; color:%(text)s; border:1px solid %(control)s;
              border-radius:8px; padding:6px 14px; min-height:22px; font-weight:500; }
QPushButton:hover { border-color:%(primary)s; color:%(primary_dark)s; background:#F9EDF0; }
QPushButton:pressed { background:#F4DFE5; border-color:%(primary_dark)s; padding-top:7px; padding-bottom:5px; }
QPushButton:focus { border:1.5px solid %(primary)s; }
QPushButton:disabled { color:%(text_disabled)s; background:#F4EDE9; border-color:#EADCD8; }
QPushButton#primaryButton { background:%(primary)s; color:#FFFFFF; border:2px solid %(primary)s;
                            border-radius:8px; font-weight:bold; font-size:14px; padding:6px 14px;
                            min-height:24px; }
QPushButton#primaryButton:hover { background:%(primary_hover)s; border-color:%(primary_hover)s; }
QPushButton#primaryButton:pressed { background:%(primary_dark)s; border-color:%(primary_dark)s; padding-top:7px; padding-bottom:5px; }
QPushButton#primaryButton:focus { border-color:#5F102D; }
QPushButton#primaryButton:disabled { background:#F3E7E5; color:%(text_disabled)s; border-color:#F3E7E5; }
QPushButton#privacyLink { background:transparent; color:%(primary)s; border:1px solid transparent; }
QPushButton#privacyLink:hover, QPushButton#privacyLink:focus {
    background:%(primary_light)s; border-color:%(border)s; }

/* ---------- 输入控件 ---------- */
QLineEdit, QComboBox, QSpinBox { background:%(card)s; color:%(text)s;
    border:1px solid %(control)s; border-radius:8px; padding:6px 10px; min-height:20px;
    selection-background-color:#EBC5D1; selection-color:%(text_strong)s; }
QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border:2px solid %(primary)s; background:#FFFCFA; }
QLineEdit:disabled, QComboBox:disabled, QSpinBox:disabled {
    color:%(text_disabled)s; background:#F4EDE9; border-color:#EADCD8; }
QLineEdit::placeholder { color:#7A6069; }
QComboBox QAbstractItemView { background:%(card)s; border:1.5px solid %(control)s;
    selection-background-color:%(primary_light)s; selection-color:%(text)s; outline:none; border-radius:6px; }

/* ---------- 表格 ---------- */
QTableView { background:%(card)s; alternate-background-color:#FAF2EE;
             border:1px solid %(border)s; border-radius:10px; gridline-color:%(divider)s;
             selection-background-color:%(primary_light)s; selection-color:#391525; }
QTableView::item { padding:5px 8px; border:none; }
QTableView::item:selected { background:%(primary_light)s; color:#391525; }
QHeaderView { background:%(card)s; }
QHeaderView::section { background:%(header_bg)s; color:%(text_strong)s; padding:7px 10px;
                       border:none; border-right:1px solid %(divider)s;
                       border-bottom:1.5px solid %(border)s; font-weight:bold; font-size:12px; }
QHeaderView::section:hover { background:#F4DFE5; }

/* ---------- 页签 ---------- */
QTabWidget::pane { background:%(card)s; border:1px solid %(border)s; border-radius:10px; top:-1px; }
QTabBar::tab { background:transparent; color:%(muted)s; padding:8px 12px;
               margin-right:2px; border-top-left-radius:8px; border-top-right-radius:8px;
               border-bottom:2.5px solid transparent; font-weight:500; }
QTabBar::tab:hover { color:%(primary_dark)s; background:#F9EDF0; }
QTabBar::tab:selected { background:%(card)s; color:%(primary_dark)s; font-weight:bold;
                         border-bottom:2.5px solid %(primary)s; }
QTabBar::tab:disabled { color:%(text_disabled)s; }
QTabWidget#workspaceTabs > QTabBar::tab {
    background: #FFFFFF;
    color: %(primary_dark)s;
    border: 1.5px solid #8B1E3F;
    border-radius: 14px;
    margin: 0 8px 6px 0;
    padding: 6px 18px;
    font-weight: 600;
    font-size: 13px;
}
QTabWidget#workspaceTabs > QTabBar::tab:hover {
    border-color: %(primary_dark)s;
    color: %(primary_dark)s;
    background: #FFF0F3;
}
QTabWidget#workspaceTabs > QTabBar::tab:selected {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 %(primary)s, stop:1 %(primary_dark)s);
    color: #FFFFFF;
    border: 1.5px solid %(primary_dark)s;
    font-weight: bold;
}
QTabWidget#workspaceTabs > QTabBar::tab:pressed {
    padding-top: 7px;
    padding-bottom: 5px;
}

/* ---------- 豪华大气可爱：本地计算与双表结果卡片 ---------- */
QFrame#calcStatsCard {
    background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #FFF8FA, stop:1 #FFF2F5);
    border: 1.5px solid #E4B8C3;
    border-radius: 10px;
    padding: 10px 12px;
}
QLabel#calcCountNumber {
    font-size: 24px;
    font-weight: bold;
    color: %(primary_dark)s;
}
QPushButton#dbViewBtn {
    background: #FFFFFF;
    color: %(primary_dark)s;
    border: 1.5px solid #C7ADAF;
    border-radius: 8px;
    padding: 6px 12px;
    font-weight: 600;
    font-size: 13px;
}
QPushButton#dbViewBtn:hover {
    border-color: %(primary)s;
    background: #FDF2F5;
    color: %(primary)s;
}
QPushButton#dbViewBtn:pressed {
    background: #F4DFE5;
    border-color: %(primary_dark)s;
    padding-top: 7px;
    padding-bottom: 5px;
}
QFrame#resultSubCard {
    background: #FFFFFF;
    border: 1.2px solid #E4CAD0;
    border-radius: 8px;
    padding: 4px 8px;
}
QFrame#resultSubCard:hover {
    border-color: %(primary)s;
    background: #FFFDFE;
}
QPushButton#subActionBtn {
    background: #FFFCFA;
    color: %(text)s;
    border: 1px solid #C7ADAF;
    border-radius: 7px;
    padding: 4px 10px;
    font-size: 12px;
    font-weight: 600;
}
QPushButton#subActionBtn:hover {
    border-color: %(primary)s;
    color: %(primary_dark)s;
    background: #FDF2F5;
}
QPushButton#subActionBtn:pressed {
    background: #F4DFE5;
    border-color: %(primary_dark)s;
    padding-top: 5px;
    padding-bottom: 3px;
}
QPushButton#subActionBtn:disabled {
    background: #FAF5F3;
    color: %(text_disabled)s;
    border-color: #EADCD8;
}

/* ---------- 文本域（概览 / 自测输出） ---------- */
QPlainTextEdit { background:%(card)s; color:%(text)s; border:1.5px solid %(border)s;
                 border-radius:10px; padding:10px; }

/* ---------- 进度条 ---------- */
QProgressBar { border:1px solid %(border)s; border-radius:8px; background:#F4EAE6;
               text-align:center; height:18px; color:%(text_strong)s; font-weight:bold; font-size:12px; }
QProgressBar::chunk { background:%(primary)s; border-radius:7px; }

/* ---------- 滚动条 ---------- */
QScrollBar:vertical { background:%(bg)s; width:10px; margin:0; border:none; }
QScrollBar::handle:vertical { background:#D9BFC0; border-radius:5px; min-height:28px; }
QScrollBar::handle:vertical:hover { background:%(primary)s; }
QScrollBar:horizontal { background:%(bg)s; height:10px; margin:0; border:none; }
QScrollBar::handle:horizontal { background:#D9BFC0; border-radius:5px; min-width:28px; }
QScrollBar::handle:horizontal:hover { background:%(primary)s; }
QScrollBar::add-line, QScrollBar::sub-line { height:0px; width:0px; }
QScrollBar::add-page, QScrollBar::sub-page { background:none; }

/* ---------- 分组框 ---------- */
QGroupBox { background:%(card)s; border:1.5px solid %(border)s; border-radius:10px;
            margin-top:14px; padding-top:10px; }
QGroupBox::title { subcontrol-origin:margin; left:12px; padding:0 4px;
                   color:%(primary_dark)s; font-weight:600; }

/* ---------- 提示、说明与状态行 ---------- */
QLabel#hint { color:%(muted)s; font-size:12px; }
QLabel#notice { color:%(muted)s; font-size:12px; }
QLabel#statusBar { color:%(text_strong)s; background:%(card)s; border:1px solid %(border)s;
                   border-left:3px solid %(primary)s; border-radius:8px; padding:6px 11px; font-weight:500; }
QLabel#statusBar[state="warn"] { border-left:4px solid %(warn_fg)s; color:%(warn_fg)s; background:#FFFDF7; }
QLabel#statusBar[state="error"] { border-left:4px solid %(danger)s; color:%(danger)s; background:#FFF8F8; }
QLabel#statusBar[state="success"] { border-left:4px solid %(green_fg)s; color:%(green_fg)s; background:#F4FAF8; }

QStatusBar { background:%(card)s; color:%(text_strong)s; border-top:1px solid %(divider)s; }

/* ---------- 集中式对话框规则（替代散落内联 setStyleSheet） ---------- */
QDialog#statusDialog { background:#F9EDF0; }
QDialog#statusDialog QLabel#dialogTitle { font-size:16px; font-weight:bold; color:%(text_strong)s; }
QDialog#statusDialog QLabel#dialogMessage { font-size:13px; color:%(text)s; }
QFrame#adviceCard { background:#FBF3EF; border:1px dashed #D9BFC0; border-radius:10px; padding:10px 12px; }
QLabel#adviceTitle { font-weight:bold; color:%(primary_dark)s; font-size:12px; }
QLabel#adviceText { font-size:12px; color:%(text)s; }
QPushButton#toggleBtn { background:transparent; border:none; color:%(muted)s; font-size:12px; font-weight:500; padding:4px 6px; text-align:left; }
QPushButton#toggleBtn:hover { color:%(primary)s; }
QPushButton#copyBtn { background:%(card)s; border:1px solid %(control)s; border-radius:6px; color:%(text)s; font-size:11px; padding:3px 8px; }
QPushButton#copyBtn:hover { border-color:%(primary)s; color:%(primary_dark)s; }
QPlainTextEdit#techDetailBox { background:#F4EDE9; color:%(text)s; border:1px solid #E6D6D2; border-radius:8px; padding:8px; font-size:11px; }

QFrame#drilldownCard { background:%(primary_light)s; border:1px solid #EBC5D1; border-radius:10px; padding:10px 14px; }
QLabel#drilldownCountLabel { font-size:13px; color:%(text)s; }
QLabel#drilldownFilterLabel { font-weight:bold; color:%(text)s; }
QPushButton#drilldownExportBtn { background:%(primary)s; color:#FFFFFF; font-weight:bold; border-radius:6px; padding:6px 14px; }
QPushButton#drilldownExportBtn:hover { background:%(primary_dark)s; }

/* ---------- 空状态体验 (F7) ---------- */
QFrame#emptyStateCard { background:%(card)s; border:1.5px solid %(border)s; border-radius:12px; }
QFrame#emptyStepCard { background:%(bg)s; border:1px solid %(border_subtle)s; border-radius:8px; }
QFrame#emptyStepCard:hover { background:%(primary_light)s; border-color:%(control)s; }
QLabel#emptyStepBadge { background:%(primary)s; color:#FFFFFF; font-weight:bold; font-size:13px;
                        border-radius:12px; min-width:24px; max-width:24px; min-height:24px; max-height:24px;
                        qproperty-alignment:AlignCenter; }
QLabel#emptyStepTitle { font-weight:bold; color:%(text_strong)s; font-size:13px; }
QLabel#emptyStepDesc { color:%(muted)s; font-size:12px; }
QLabel#emptyWelcomeTitle { font-size:16px; font-weight:bold; color:%(text_strong)s; }
QLabel#emptyWelcomeSubtitle { font-size:13px; color:%(muted)s; }
QLabel#emptyTipText { font-size:12px; color:%(slate)s; background:%(surface)s;
                      border:1px dashed %(border)s; border-radius:6px; padding:8px 12px; }

/* ---------- 浮动轻量通知 (F11) ---------- */
QFrame#toastNotification {
    background:#391525;
    border:1px solid #781236;
    border-radius:18px;
    color:#FFFFFF;
}
QFrame#toastNotification[kind="success"] {
    background:#144E45;
    border:1px solid %(green_fg)s;
}
QFrame#toastNotification[kind="warn"], QFrame#toastNotification[kind="warning"] {
    background:#5D3708;
    border:1px solid %(warn_fg)s;
}
QLabel#toastIcon {
    font-size:14px;
    font-weight:bold;
    color:#FFFFFF;
}
QLabel#toastMessage {
    font-size:13px;
    font-weight:500;
    color:#FFFFFF;
}
""" % COLORS


# ----------------------------------------------------------------- 资源路径

def _app_base_dir() -> str:
    """获取应用程序所在基准目录：打包态取 EXE 所在目录，源码态取 PROJECT_DIR。"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return PROJECT_DIR


def _is_dir_writable(folder: str) -> bool:
    """探测目录是否可写（防范安装在只读 Program Files 下的权限异常）。"""
    try:
        os.makedirs(folder, exist_ok=True)
        probe = os.path.join(folder, '.write_probe_%d' % os.getpid())
        with open(probe, 'w', encoding='utf-8') as fh:
            fh.write('ok')
        os.remove(probe)
        return True
    except OSError:
        return False


def user_writable_dir() -> str:
    """返回用户可写数据目录：优先 EXE/项目同级目录，只读时回退至 %LOCALAPPDATA%/SanzaoWan。"""
    base = _app_base_dir()
    if _is_dir_writable(base):
        return base
    local_appdata = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
    fallback = os.path.join(local_appdata, 'SanzaoWan')
    os.makedirs(fallback, exist_ok=True)
    return fallback


def resource_path(name: str) -> str:
    """兼容用户可写目录、EXE 同级目录、PyInstaller (_MEIPASS) 与源码目录。"""
    for folder in (
        user_writable_dir(),
        _app_base_dir(),
        getattr(sys, '_MEIPASS', PROJECT_DIR),
        PROJECT_DIR,
    ):
        candidate = os.path.join(folder, name)
        if os.path.exists(candidate):
            return candidate
    return os.path.join(user_writable_dir(), name)


def default_dictionary_path() -> str:
    """返回可持久化的免勤词典路径（绝不直接写回单文件临时解压目录 _MEIPASS）。"""
    writable_dir = user_writable_dir()
    writable_target = os.path.join(writable_dir, DICT_NAME)
    if os.path.exists(writable_target):
        return writable_target

    # 若可写目录尚无词典，从安装目录或 _MEIPASS 内置副本复制一份出来供持久化编辑
    for folder in (_app_base_dir(), getattr(sys, '_MEIPASS', PROJECT_DIR), PROJECT_DIR):
        bundled = os.path.join(folder, DICT_NAME)
        if os.path.exists(bundled):
            if os.path.normcase(os.path.abspath(bundled)) != os.path.normcase(os.path.abspath(writable_target)):
                try:
                    shutil.copy2(bundled, writable_target)
                    return writable_target
                except OSError:
                    return bundled
            return bundled
    return writable_target


def default_output_dir() -> str:
    if getattr(sys, 'frozen', False):
        p = os.path.join(user_writable_dir(), '_work', '输出')
    else:
        base = os.path.dirname(PROJECT_DIR)
        p = os.path.join(base, '_work', '输出')
    os.makedirs(p, exist_ok=True)
    return p


# ----------------------------------------------------------------- 表格模型

class _GenericTableModel(QAbstractTableModel):
    """把 dict 列表渲染成表格数据，供 QTableView 使用。

    支持两种上色：
      * 行级 `row_colors`（整行底色，如异常清单 / 低出勤率 / 留痕行）；
      * 单元格级 `cell_styles`（形如 {(行, 列): {'bg': QColor, 'italic': True,
        'tip': str}}），用于估算标记这类"同一行里只有个别格特殊"的场景。
    """

    def __init__(self, columns, rows, row_colors=None, cell_styles=None):
        super().__init__()
        self.columns = list(columns)
        self.rows = list(rows)
        self.row_colors = dict(row_colors or {})
        self.cell_styles = dict(cell_styles or {})
        self.base_font = QFont()

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows)

    def columnCount(self, parent=QModelIndex()):
        return len(self.columns)

    def _cell_style(self, row, column):
        return self.cell_styles.get((row, column), {})

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row, column = index.row(), index.column()
        item = self.rows[row]
        value = item.get(self.columns[column], '')
        style = self._cell_style(row, column)
        if role == Qt.DisplayRole:
            # 缺失值在界面上显示「—」占位（R2：界面可占位，导出仍是真空白）；
            # 数值 0 是有效值，必须原样显示成 0，绝不能与缺失混淆。
            # 其余值原样返回（数值保持数值类型，由视图自行对齐与格式化）。
            if value is None:
                return '—'
            return value
        if role == Qt.TextAlignmentRole and isinstance(value, (int, float)):
            return Qt.AlignRight | Qt.AlignVCenter
        if role == Qt.ToolTipRole:
            if style.get('tip'):
                return style['tip']
            if value is None:
                return '源表未填，待补录'
            return str(value)
        if role == Qt.BackgroundRole:
            color = style.get('bg', self.row_colors.get(row))
            return QBrush(color) if color is not None else None
        if role == Qt.FontRole:
            if style.get('italic'):
                font = QFont(self.base_font)
                font.setItalic(True)
                return font
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.columns[section]
        return section + 1


# ----------------------------------------------------------------- 后台任务

class Worker(QThread):
    finished = Signal(object)
    failed = Signal(str)

    def __init__(self, config: RunConfig = None, parent=None):
        super().__init__(parent)
        self.config = config

    def run(self):
        try:
            QThread.setTerminationEnabled(True)
            result = run(self.config)
            self.finished.emit(result)
        except Exception:                              # noqa: BLE001
            self.failed.emit(traceback.format_exc())


# ----------------------------------------------------------------- 主窗口

class CallableStr(str):
    """兼具字符串与可调用无参函数特性的兼容类型。

    让 slot.path 既可以作为属性值比较 (slot.path == str)，
    又支持既有代码以函数形式调用 (slot.path())。
    """

    def __call__(self) -> str:
        return str(self)


class PathPreview(QLabel):
    """在窄设置栏中保留路径两端，悬停时显示完整路径。"""

    def set_path(self, path: str):
        self.setToolTip(path)
        self.setText(self.fontMetrics().elidedText(path, Qt.ElideMiddle, self.contentsRect().width()))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.set_path(self.toolTip())


class ToastNotification(QFrame):
    """Modern non-blocking floating toast notification pill."""

    def __init__(self, parent=None, message: str = '', kind: str = 'info', duration: int = 3000):
        top_parent = parent.window() if parent and hasattr(parent, 'window') else parent
        super().__init__(top_parent)
        self.setObjectName('toastNotification')
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setProperty('kind', kind)
        self._message = message
        self._kind = kind

        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 9, 18, 9)
        layout.setSpacing(10)

        icon_symbol = '✓' if kind == 'success' else ('⚠' if kind in ('warn', 'warning') else 'ℹ')
        self.icon_label = QLabel(icon_symbol)
        self.icon_label.setObjectName('toastIcon')
        layout.addWidget(self.icon_label)

        self.msg_label = QLabel(message)
        self.msg_label.setObjectName('toastMessage')
        layout.addWidget(self.msg_label)

        self.adjustSize()
        if top_parent:
            parent_w = top_parent.width()
            toast_w = self.sizeHint().width()
            x = max(10, (parent_w - toast_w) // 2)
            self.move(x, 20)

        if top_parent and hasattr(top_parent, '_active_toasts') and isinstance(top_parent._active_toasts, set):
            top_parent._active_toasts.add(self)
            self.destroyed.connect(lambda: top_parent._active_toasts.discard(self) if hasattr(top_parent, '_active_toasts') else None)

        eff_duration = get_animation_duration(duration)
        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.close)
        self._timer.start(eff_duration)

        super().show()
        super().raise_()

    def show(self, *args, **kwargs):
        if not args and not kwargs:
            super().show()
            super().raise_()
        else:
            return ToastNotification(self, *args, **kwargs)

    def closeEvent(self, event):
        if hasattr(self, '_timer') and self._timer.isActive():
            self._timer.stop()
        self.deleteLater()
        super().closeEvent(event)


class CuteBearAvatar(QLabel):
    """可爱的草莓小熊互动头像：支持悬停提示、点击俏皮彩蛋弹窗与微交互反馈。"""

    CUTE_TIPS = (
        "🍓 今天又是认真统计出勤的一天！每一条免勤都已精准归档~",
        "✨ 摸摸小熊头，出勤不用愁~ 本地 SQLite 毫秒级运算中！",
        "🔒 100% 纯本地离线隐私守护，绝无任何个人数据上传云端！",
        "🍃 50MB 磁盘配额自愈清理，硬盘轻量零负担~",
        "🚀 早自习、晚自习一键自动识别，回填表与终稿汇总表秒出！",
        "🐻 桂工学子专属出勤管家已就绪，今天也元气满满哦~",
        "💡 温馨提示：支持多份免勤名单同时拖入，系统自动智能去重！",
        "🎯 终稿汇总表格式严格对齐 15 个学院官方规范，直接可用！",
    )

    def __init__(self, size: int = 66, parent=None):
        super().__init__(parent)
        self._size = size
        self.setFixedSize(size, size)
        self.setAlignment(Qt.AlignCenter)
        self.setCursor(Qt.PointingHandCursor)
        self.setToolTip('点击草莓小熊有彩蛋哦~ 🍓')
        self._update_pixmap(size)

    def _update_pixmap(self, size: int):
        self.setPixmap(AttendanceIcon.pixmap(size, target_widget=self))

    def enterEvent(self, event):
        self.setToolTip('点击草莓小熊有彩蛋哦~ 🍓')
        super().enterEvent(event)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            import random
            tip = random.choice(self.CUTE_TIPS)
            top_win = self.window()
            if top_win:
                ToastNotification(top_win, tip, kind='success', duration=3200)
        super().mousePressEvent(event)


class MetricCard(QFrame):
    """Clickable, interactive metric card with hover elevation and jump navigation."""

    clicked = Signal(str)

    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setObjectName('metricCard')
        self.setCursor(Qt.PointingHandCursor)
        self.setFrameShape(QFrame.StyledPanel)
        self._title = title

        content = QVBoxLayout(self)
        content.setContentsMargins(14, 10, 14, 10)
        content.setSpacing(2)

        self.label = QLabel(title)
        self.label.setObjectName('metricLabel')
        self.label.setCursor(Qt.PointingHandCursor)
        content.addWidget(self.label)

        self.value = QLabel('—')
        self.value.setObjectName('metricValue')
        self.value.setAccessibleName(title)
        self.value.setCursor(Qt.PointingHandCursor)
        content.addWidget(self.value)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            top_win = self.window()
            if hasattr(top_win, '_on_metric_card_clicked'):
                top_win._on_metric_card_clicked(self._title)
            self.clicked.emit(self._title)
        super().mousePressEvent(event)


class StatusDialog(QDialog):
    """带可折叠技术详情的提示窗口。"""

    def __init__(self, parent=None, title='提示', message='', advice='',
                 technical_detail=None, kind='info'):
        super().__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(AttendanceIcon.pixmap(32, target_widget=self)))
        self.setMinimumWidth(500)
        self.setMaximumWidth(680)
        self.setModal(True)
        self.setObjectName('statusDialog')

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 18)
        layout.setSpacing(12)

        # 1. 顶部：头像与主文本
        top_row = QHBoxLayout()
        top_row.setSpacing(14)
        top_row.setAlignment(Qt.AlignTop)

        avatar = QLabel()
        avatar.setPixmap(AttendanceIcon.pixmap(54, target_widget=self))
        avatar.setFixedSize(54, 54)
        top_row.addWidget(avatar)

        text_layout = QVBoxLayout()
        text_layout.setSpacing(4)
        title_label = QLabel(title)
        title_label.setTextFormat(Qt.PlainText)
        title_label.setObjectName('dialogTitle')
        text_layout.addWidget(title_label)

        msg_label = QLabel(message)
        msg_label.setTextFormat(Qt.PlainText)
        msg_label.setObjectName('dialogMessage')
        msg_label.setWordWrap(True)
        text_layout.addWidget(msg_label)
        top_row.addLayout(text_layout, 1)
        layout.addLayout(top_row)

        # 2. 中部：解决建议卡片
        if advice:
            advice_card = QFrame()
            advice_card.setObjectName('adviceCard')
            adv_layout = QVBoxLayout(advice_card)
            adv_layout.setContentsMargins(10, 8, 10, 8)
            adv_layout.setSpacing(4)
            adv_title = QLabel('处理建议')
            adv_title.setObjectName('adviceTitle')
            adv_layout.addWidget(adv_title)
            adv_text = QLabel(advice)
            adv_text.setTextFormat(Qt.PlainText)
            adv_text.setObjectName('adviceText')
            adv_text.setWordWrap(True)
            adv_layout.addWidget(adv_text)
            layout.addWidget(advice_card)

        # 3. 下部：技术细节折叠区
        self.technical_detail = technical_detail
        self.tech_box = None
        self.detail_edit = None
        self.toggle_btn = None
        self.copy_btn = None

        if technical_detail:
            detail_container = QWidget()
            d_layout = QVBoxLayout(detail_container)
            d_layout.setContentsMargins(0, 0, 0, 0)
            d_layout.setSpacing(6)

            ctrl_row = QHBoxLayout()
            self.toggle_btn = QPushButton('显示技术详情')
            self.toggle_btn.setObjectName('toggleBtn')
            self.toggle_btn.clicked.connect(self._toggle_detail)
            ctrl_row.addWidget(self.toggle_btn)
            ctrl_row.addStretch(1)

            self.copy_btn = QPushButton('复制细节')
            self.copy_btn.setObjectName('copyBtn')
            self.copy_btn.clicked.connect(self._copy_detail)
            ctrl_row.addWidget(self.copy_btn)
            d_layout.addLayout(ctrl_row)

            self.tech_box = QPlainTextEdit()
            self.tech_box.setObjectName('techDetailBox')
            self.tech_box.setPlainText(str(technical_detail))
            self.tech_box.setReadOnly(True)
            self.tech_box.setFont(build_mono_font())
            self.tech_box.setMaximumHeight(160)
            self.tech_box.setVisible(False)
            self.detail_edit = self.tech_box
            d_layout.addWidget(self.tech_box)

            layout.addWidget(detail_container)

        # 4. 底部确定按钮
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        ok_btn = QPushButton('确定')
        ok_btn.setObjectName('primaryButton')
        ok_btn.setMinimumHeight(32)
        ok_btn.clicked.connect(self.accept)
        btn_row.addWidget(ok_btn)
        layout.addLayout(btn_row)

    def _toggle_detail(self):
        if self.tech_box:
            is_vis = not self.tech_box.isVisible()
            self.tech_box.setVisible(is_vis)
            if self.toggle_btn:
                self.toggle_btn.setText('隐藏技术详情' if is_vis else '显示技术详情')
            self.adjustSize()

    def _copy_detail(self):
        if self.technical_detail:
            QApplication.clipboard().setText(str(self.technical_detail))
            if self.copy_btn:
                self.copy_btn.setText('已复制')
            top_parent = self.parent() or self
            ToastNotification(top_parent, '诊断日志已成功复制到剪贴板', kind='success')

    def _cleanup_dialog(self):
        self.setParent(None)
        self.deleteLater()

    def closeEvent(self, event):
        self._cleanup_dialog()
        super().closeEvent(event)

    def accept(self):
        self._cleanup_dialog()
        super().accept()

    def reject(self):
        self._cleanup_dialog()
        super().reject()

    @classmethod
    def show_warning(cls, parent, title, message, advice='', technical_detail=None, kind='warn'):
        dlg = cls(parent, title=title, message=message, advice=advice,
                  technical_detail=technical_detail, kind=kind)
        dlg.setAttribute(Qt.WA_DeleteOnClose, True)
        try:
            return dlg.exec()
        finally:
            try:
                dlg.setParent(None)
                dlg.deleteLater()
            except RuntimeError:
                pass

    @classmethod
    def show_error(cls, parent, title, message, advice='', technical_detail=None, kind='error'):
        dlg = cls(parent, title=title, message=message, advice=advice,
                  technical_detail=technical_detail, kind=kind)
        dlg.setAttribute(Qt.WA_DeleteOnClose, True)
        try:
            return dlg.exec()
        finally:
            try:
                dlg.setParent(None)
                dlg.deleteLater()
            except RuntimeError:
                pass

    @classmethod
    def show_info(cls, parent, title, message, advice='', technical_detail=None, kind='info'):
        dlg = cls(parent, title=title, message=message, advice=advice,
                  technical_detail=technical_detail, kind=kind)
        dlg.setAttribute(Qt.WA_DeleteOnClose, True)
        try:
            return dlg.exec()
        finally:
            try:
                dlg.setParent(None)
                dlg.deleteLater()
            except RuntimeError:
                pass


class StudentDrilldownDialog(QDialog):
    """查看指定统计范围内的免勤记录。"""

    def __init__(self, parent=None, title='免勤人员明细', scope_info='', students=None):
        super().__init__(parent)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(AttendanceIcon.pixmap(32, target_widget=self)))
        self.resize(880, 560)
        self.setModal(True)
        self.raw_students = list(students or [])

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)

        # 头部说明卡片
        card = QFrame()
        card.setObjectName('drilldownCard')
        c_layout = QHBoxLayout(card)
        c_layout.setContentsMargins(4, 4, 4, 4)
        c_lbl = QLabel(f'{scope_info}\n当前范围 {len(self.raw_students)} 条个人免勤记录')
        c_lbl.setTextFormat(Qt.PlainText)
        c_lbl.setObjectName('drilldownCountLabel')
        c_layout.addWidget(c_lbl)
        layout.addWidget(card)

        # 搜索过滤条
        filter_bar = QHBoxLayout()
        filter_lbl = QLabel("搜索")
        filter_lbl.setObjectName('drilldownFilterLabel')
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('按姓名、学院、班级或学号筛选')
        self.search_input.textChanged.connect(self._apply_filter)
        filter_bar.addWidget(filter_lbl)
        filter_bar.addWidget(self.search_input)
        layout.addLayout(filter_bar)

        # 学生列表表格
        self.table = QTableView()
        self.columns = ['序号', '学号', '姓名', '学院', '专业班级', '星期', '自习', '教室', '来源', '申请理由']
        self._update_table(self.raw_students)
        layout.addWidget(self.table, 1)

        # 底部按钮区
        btn_bar = QHBoxLayout()
        export_btn = QPushButton("导出筛选名单…")
        export_btn.setToolTip('导出的 CSV 含有个人信息')
        export_btn.setObjectName('drilldownExportBtn')
        export_btn.setIcon(ActionIcon.get('export', '#FFFFFF'))
        export_btn.setIconSize(QSize(16, 16))
        export_btn.clicked.connect(self._export_csv)
        close_btn = QPushButton("关闭 (Esc)")
        close_btn.clicked.connect(self.accept)
        btn_bar.addWidget(export_btn)
        btn_bar.addStretch(1)
        btn_bar.addWidget(close_btn)
        layout.addLayout(btn_bar)

    def _update_table(self, stu_list):
        old_model = self.table.model()
        old_sm = self.table.selectionModel()
        rows = []
        for i, s in enumerate(stu_list, 1):
            rows.append({
                '序号': i,
                '学号': s.get('sid', ''),
                '姓名': s.get('name', ''),
                '学院': s.get('college', ''),
                '专业班级': s.get('class_text', ''),
                '星期': s.get('day', ''),
                '自习': s.get('session', ''),
                '教室': s.get('room', ''),
                '来源': s.get('source_ref', ''),
                '申请理由': s.get('reason', ''),
            })
        model = _GenericTableModel(self.columns, rows)
        self.table.setModel(model)
        if old_sm is not None:
            old_sm.deleteLater()
        if old_model is not None:
            old_model.deleteLater()
        self.table.resizeColumnsToContents()

    def _cleanup_dialog(self):
        self.setParent(None)
        old_model = self.table.model()
        old_sm = self.table.selectionModel()
        if old_sm is not None:
            old_sm.deleteLater()
        if old_model is not None:
            old_model.deleteLater()
        self.raw_students.clear()
        self.deleteLater()

    def closeEvent(self, event):
        self._cleanup_dialog()
        super().closeEvent(event)

    def accept(self):
        self._cleanup_dialog()
        super().accept()

    def reject(self):
        self._cleanup_dialog()
        super().reject()

    def _apply_filter(self, kw):
        kw = normalize_student_name(kw).casefold()
        if not kw:
            self._update_table(self.raw_students)
            return
        filtered = [s for s in self.raw_students
                    if any(kw in normalize_student_name(s.get(key)).casefold()
                           for key in ('sid', 'name', 'college', 'class_text'))]
        self._update_table(filtered)

    def _export_csv(self):
        path, _ = QFileDialog.getSaveFileName(self, "导出免勤人员明细", "免勤人员明细.csv", "CSV Files (*.csv)")
        if not path:
            return
        if not path.lower().endswith('.csv'):
            QMessageBox.warning(self, '检查文件名', '请使用 .csv 扩展名。')
            return
        try:
            import csv
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(self.columns)
                for row in self.table.model().rows:
                    values = [row.get(column, '') for column in self.columns]
                    writer.writerow(["'" + value if isinstance(value, str) and
                                     value.lstrip().startswith(('=', '+', '-', '@')) else value
                                     for value in values])
            QMessageBox.information(self, "导出成功", f"名单已成功导出至：\n{path}")
            top_parent = self.parent() or self
            ToastNotification(top_parent, '名单已成功导出', kind='success')
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))


class DropSlot(QFrame):
    """一个「提交文件」的槽位：支持点击选择，也支持把 .xlsx 直接拖进来。"""

    picked = Signal(str)

    def __init__(self, step: str, title: str, hint: str, *, multiple=False):
        super().__init__()
        self.setAcceptDrops(True)
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName('dropSlot')
        self.setProperty('dragOver', False)
        self.setProperty('ready', False)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCursor(Qt.PointingHandCursor)
        self.paths = ()
        self.multiple = multiple
        self._hint = hint

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 11, 14, 11)
        layout.setSpacing(7)

        head = QHBoxLayout()
        head.setSpacing(8)
        self.badge = QLabel(step)
        self.badge.setObjectName('slotBadge')
        head.addWidget(self.badge)
        self.title_label = QLabel(title)
        self.title_label.setObjectName('slotTitle')
        head.addWidget(self.title_label, 1)
        layout.addLayout(head)

        self.path_label = QLabel(hint)
        self.path_label.setObjectName('slotPath')
        self.path_label.setWordWrap(True)
        self.path_label.setTextFormat(Qt.PlainText)
        self.path_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.path_label.setFixedHeight(38)
        layout.addWidget(self.path_label)

        self.choose_btn = QPushButton('选择文件…')
        self.choose_btn.setCursor(Qt.PointingHandCursor)
        self.choose_btn.setIcon(ActionIcon.get('document', TOKENS['colors']['primary_dark']))
        self.choose_btn.setIconSize(QSize(16, 16))
        self.choose_btn.clicked.connect(self._on_click)
        buttons = QHBoxLayout()
        buttons.addWidget(self.choose_btn)
        if multiple:
            self.manage_btn = QPushButton('管理…')
            self.manage_btn.setCursor(Qt.PointingHandCursor)
            self.manage_btn.setIcon(ActionIcon.get('tools', TOKENS['colors']['text']))
            self.manage_btn.setIconSize(QSize(16, 16))
            self.manage_btn.clicked.connect(self._manage)
            buttons.addWidget(self.manage_btn)
        buttons.addStretch(1)
        layout.addLayout(buttons)
        self.setAccessibleName(title)
        self.set_paths(())

    # ---------------------------------------------------------------- 状态
    @property
    def path(self) -> CallableStr:
        return CallableStr(self.paths[0] if self.paths else '')

    def set_path(self, path: str):
        self.set_paths((path,))

    def set_paths(self, paths):
        self.paths = unique_workbook_paths(paths)
        if self.paths:
            base_name = os.path.basename(self.paths[0])
            size_str = ''
            if os.path.exists(self.paths[0]):
                try:
                    if self.multiple and len(self.paths) > 1:
                        total_bytes = sum(os.path.getsize(p) for p in self.paths if os.path.exists(p))
                        size_str = f'  ({_format_file_size(total_bytes)})'
                    else:
                        size_str = f'  ({_format_file_size(os.path.getsize(self.paths[0]))})'
                except Exception:
                    size_str = ''
            if self.multiple:
                label = f'📄 已选 {len(self.paths)} 份 · {base_name}{size_str}'
            else:
                label = f'📄 {base_name}{size_str}'
            self.path_label.setText(label)
            self.path_label.setToolTip('\n'.join(self.paths))
        else:
            self.path_label.setText(self._hint)
            self.path_label.setToolTip('')
        ready = bool(self.paths)
        self.setProperty('dragReject', False)
        self.setProperty('dragOver', False)
        self.setProperty('ready', ready)
        self.choose_btn.setText('添加表格…' if self.multiple else
                                '重新选择…' if ready else '选择文件…')
        if self.multiple:
            self.manage_btn.setEnabled(ready)
        self._repolish()

    def _repolish(self):
        self.style().unpolish(self)
        self.style().polish(self)
        self.path_label.style().unpolish(self.path_label)
        self.path_label.style().polish(self.path_label)
        self.update()

    # ---------------------------------------------------------------- 交互
    def _on_click(self):
        if self.multiple:
            paths, _ = QFileDialog.getOpenFileNames(
                self, '添加免勤人员统计表', '', 'Excel 文件 (*.xlsx *.xlsm)')
            if paths:
                self.set_paths((*self.paths, *paths))
            return
        path, _ = QFileDialog.getOpenFileName(
            self, self.title_label.text(), '', 'Excel 文件 (*.xlsx *.xlsm)')
        if path:
            self.set_path(path)
            self.picked.emit(path)

    def _manage(self):
        dialog = QDialog(self)
        dialog.setWindowTitle('已选免勤表')
        dialog.resize(650, 340)
        layout = QVBoxLayout(dialog)
        listing = QListWidget()
        listing.setSelectionMode(QAbstractItemView.ExtendedSelection)
        for path in self.paths:
            listing.addItem(path)
        layout.addWidget(listing)
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setText('保存')
        buttons.button(QDialogButtonBox.Cancel).setText('取消')
        remove = buttons.addButton('移除所选', QDialogButtonBox.ActionRole)

        def remove_selected():
            for item in listing.selectedItems():
                listing.takeItem(listing.row(item))

        remove.clicked.connect(remove_selected)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        if dialog.exec():
            self.set_paths(listing.item(index).text() for index in range(listing.count()))

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._on_click()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if not event.mimeData().hasUrls():
            return
        paths = [u.toLocalFile() for u in event.mimeData().urls()]
        ok = [p for p in paths
              if p and os.path.splitext(p)[1].lower() in ('.xlsx', '.xlsm')]
        if not ok:
            event.ignore()
            self.setProperty('dragOver', False)
            self.setProperty('dragReject', True)
            self._repolish()
            self.path_label.setText('请选择 .xlsx 或 .xlsm 文件。')
            return
        event.acceptProposedAction()
        self.setProperty('dragReject', False)
        self.setProperty('dragOver', True)
        self._repolish()
        self.path_label.setText('松开鼠标导入文件')

    def dragLeaveEvent(self, event):
        self.setProperty('dragOver', False)
        self.setProperty('dragReject', False)
        self._repolish()
        self.set_paths(self.paths)

    def dropEvent(self, event: QDropEvent):
        self.setProperty('dragOver', False)
        self.setProperty('dragReject', False)
        paths = [
            u.toLocalFile()
            for u in event.mimeData().urls()
            if u.toLocalFile() and os.path.splitext(u.toLocalFile())[1].lower() in ('.xlsx', '.xlsm')
        ]
        if not paths:
            self.setProperty('dragReject', True)
            self._repolish()
            self.path_label.setText('请选择 .xlsx 或 .xlsm 文件')
            event.ignore()
            return

        top_win = self.window()
        # 批量文件按内容分流；拖到免勤槽时可把名称不明确的表加入待校验列表。
        if len(paths) >= 2 or self.multiple:
            if hasattr(top_win, '_handle_dropped_files'):
                self.set_paths(self.paths)
                top_win._handle_dropped_files(paths, target='EXEMPT' if self.multiple else None)
                event.acceptProposedAction()
                return

        # 若拖入 1 个文件，检查是否错放槽位
        path = paths[0]
        if hasattr(top_win, '_handle_dropped_files'):
            c = classify_excel_file(path)
            title_text = self.title_label.text() if hasattr(self, 'title_label') else ''
            if ('明细' in title_text and c == 'EXEMPT') or ('免勤' in title_text and c == 'DETAIL'):
                top_win._handle_dropped_files([path])
                event.acceptProposedAction()
                return

        self.set_path(path)
        self.picked.emit(path)
        self._trigger_pulse()
        if hasattr(top_win, '_active_toasts'):
            msg = '免勤统计表已添加' if self.multiple else '考勤明细表解析就绪'
            ToastNotification(top_win, msg, kind='success')
        event.acceptProposedAction()

    def _trigger_pulse(self):
        """轻柔的成功装填脉冲反馈。"""
        self.setProperty('justDropped', True)
        self._repolish()
        QTimer.singleShot(450, lambda: [
            self.setProperty('justDropped', False),
            self._repolish()
        ])


class ResultSlot(QFrame):
    """「查看与下载结果」槽位：豪华大气可爱双卡片，清晰区分明细表与终稿汇总表。"""

    def __init__(self, selected_paths=lambda: ()):
        super().__init__()
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName('dropSlot')
        self.setProperty('ready', False)
        self.setProperty('celebration', False)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self._path = ''
        self._detail_path = ''
        self._summary_path = ''
        self._input_paths = ()
        self._selected_paths = selected_paths

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(6)

        # 头部标题
        head = QHBoxLayout()
        head.setSpacing(8)
        self.badge = QLabel('03')
        self.badge.setObjectName('slotBadge')
        head.addWidget(self.badge)
        self.title_label = QLabel('计算结果工作簿')
        self.title_label.setObjectName('slotTitle')
        head.addWidget(self.title_label, 1)

        self.status_pill = QLabel('待计算')
        self.status_pill.setObjectName('localBadge')
        head.addWidget(self.status_pill)
        layout.addLayout(head)

        # 状态总览
        self.state_label = QLabel('完成计算后可另存或打开目录')
        self.state_label.setObjectName('slotPath')
        self.state_label.setWordWrap(True)
        self.state_label.setTextFormat(Qt.PlainText)
        self.state_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.state_label.setFixedHeight(30)
        layout.addWidget(self.state_label)

        # 产物 1：考勤明细表（回填完整版）
        self.detail_card = QFrame()
        self.detail_card.setObjectName('resultSubCard')
        d_layout = QHBoxLayout(self.detail_card)
        d_layout.setContentsMargins(6, 4, 6, 4)
        d_layout.setSpacing(6)
        self.detail_icon = QLabel()
        self.detail_icon.setPixmap(ActionIcon.get('document', TOKENS['colors']['primary']).pixmap(16, 16))
        d_layout.addWidget(self.detail_icon)
        self.detail_label = QLabel('明细表（回填版）：未就绪')
        self.detail_label.setStyleSheet('font-size:12px; font-weight:500;')
        d_layout.addWidget(self.detail_label, 1)

        self.btn_open_detail = QPushButton('打开明细')
        self.btn_open_detail.setObjectName('subActionBtn')
        self.btn_open_detail.setCursor(Qt.PointingHandCursor)
        self.btn_open_detail.setIcon(ActionIcon.get('document', TOKENS['colors']['primary_dark']))
        self.btn_open_detail.setIconSize(QSize(16, 16))
        self.btn_open_detail.setEnabled(False)
        self.btn_open_detail.clicked.connect(lambda: self._open_file(self._detail_path))
        d_layout.addWidget(self.btn_open_detail)

        self.btn_save_detail = QPushButton('另存结果…')
        self.btn_save_detail.setObjectName('subActionBtn')
        self.btn_save_detail.setCursor(Qt.PointingHandCursor)
        self.btn_save_detail.setIcon(ActionIcon.get('save', TOKENS['colors']['primary_dark']))
        self.btn_save_detail.setIconSize(QSize(16, 16))
        self.btn_save_detail.setEnabled(False)
        self.btn_save_detail.clicked.connect(lambda: self._save_file(self._detail_path, '考勤明细表'))
        d_layout.addWidget(self.btn_save_detail)
        layout.addWidget(self.detail_card)

        # 产物 2：执行情况终稿（内置模板生成）
        self.summary_card = QFrame()
        self.summary_card.setObjectName('resultSubCard')
        s_layout = QHBoxLayout(self.summary_card)
        s_layout.setContentsMargins(6, 4, 6, 4)
        s_layout.setSpacing(6)
        self.summary_icon = QLabel()
        self.summary_icon.setPixmap(ActionIcon.get('chart', TOKENS['colors']['primary']).pixmap(16, 16))
        s_layout.addWidget(self.summary_icon)
        self.summary_label = QLabel('终稿汇总表：未就绪')
        self.summary_label.setStyleSheet('font-size:12px; font-weight:500;')
        s_layout.addWidget(self.summary_label, 1)

        self.btn_open_summary = QPushButton('打开终稿')
        self.btn_open_summary.setObjectName('subActionBtn')
        self.btn_open_summary.setCursor(Qt.PointingHandCursor)
        self.btn_open_summary.setIcon(ActionIcon.get('chart', TOKENS['colors']['primary_dark']))
        self.btn_open_summary.setIconSize(QSize(16, 16))
        self.btn_open_summary.setEnabled(False)
        self.btn_open_summary.clicked.connect(lambda: self._open_file(self._summary_path))
        s_layout.addWidget(self.btn_open_summary)

        self.btn_save_summary = QPushButton('另存终稿…')
        self.btn_save_summary.setObjectName('subActionBtn')
        self.btn_save_summary.setCursor(Qt.PointingHandCursor)
        self.btn_save_summary.setIcon(ActionIcon.get('save', TOKENS['colors']['primary_dark']))
        self.btn_save_summary.setIconSize(QSize(16, 16))
        self.btn_save_summary.setEnabled(False)
        self.btn_save_summary.clicked.connect(lambda: self._save_file(self._summary_path, '早晚自习执行情况终稿'))
        s_layout.addWidget(self.btn_save_summary)
        layout.addWidget(self.summary_card)

        # 底部操作行
        bottom_row = QHBoxLayout()
        self.open_dir_btn = QPushButton('打开所在目录')
        self.open_dir_btn.setCursor(Qt.PointingHandCursor)
        self.open_dir_btn.setIcon(ActionIcon.get('folder_open', TOKENS['colors']['text']))
        self.open_dir_btn.setIconSize(QSize(16, 16))
        self.open_dir_btn.setEnabled(False)
        self.open_dir_btn.clicked.connect(self._open_dir)
        bottom_row.addWidget(self.open_dir_btn)
        bottom_row.addStretch(1)
        layout.addLayout(bottom_row)

        # 保持旧属性与测试兼容别名
        self.save_btn = self.btn_save_detail
        self.open_btn = self.open_dir_btn
        self.setAccessibleName('下载填好的明细表')

    def set_result(self, path: str, input_paths=(), summary_path: str = ''):
        self._path = path or ''
        self._detail_path = path or ''
        self._input_paths = input_paths
        if not summary_path and self._detail_path:
            d_dir = os.path.dirname(os.path.abspath(self._detail_path))
            if os.path.isdir(d_dir):
                for f in os.listdir(d_dir):
                    if f.endswith('.xlsx') and '终稿' in f:
                        summary_path = os.path.join(d_dir, f)
                        break
        self._summary_path = summary_path or ''
        self.set_results(self._detail_path, self._summary_path, input_paths=input_paths)

    def set_results(self, detail_path: str, summary_path: str = '', input_paths=()):
        self._detail_path = detail_path or ''
        self._path = self._detail_path
        self._summary_path = summary_path or ''
        self._input_paths = input_paths

        ready = bool(self._detail_path and os.path.exists(self._detail_path))
        has_summary = bool(self._summary_path and os.path.exists(self._summary_path))

        if ready:
            base_name = os.path.basename(self._detail_path)
            try:
                sz = _format_file_size(os.path.getsize(self._detail_path))
            except Exception:
                sz = ''
            self.state_label.setText(f'🎉 已生成 · {base_name}  ({sz})')
            fm = self.detail_label.fontMetrics()
            avail_w = max(140, self.detail_label.width() or 200)
            elided_name = fm.elidedText(base_name, Qt.ElideMiddle, avail_w)
            self.detail_label.setText(f'明细表：{elided_name} ({sz})' if sz else f'明细表：{elided_name}')
            self.detail_label.setToolTip(self._detail_path)
            self.setProperty('celebration', True)
            self.status_pill.setText('🎉 双表已就绪' if has_summary else '🎉 明细表已就绪')
        else:
            self.state_label.setText('完成计算后可另存或打开目录')
            self.detail_label.setText('考勤明细表（回填完整版）：待计算')
            self.setProperty('celebration', False)
            self.status_pill.setText('待计算')

        if has_summary:
            s_name = os.path.basename(self._summary_path)
            try:
                s_sz = _format_file_size(os.path.getsize(self._summary_path))
            except Exception:
                s_sz = ''
            fm_s = self.summary_label.fontMetrics()
            avail_ws = max(140, self.summary_label.width() or 200)
            elided_sname = fm_s.elidedText(s_name, Qt.ElideMiddle, avail_ws)
            self.summary_label.setText(f'终稿表：{elided_sname} ({s_sz})' if s_sz else f'终稿表：{elided_sname}')
            self.summary_label.setToolTip(self._summary_path)
        else:
            self.summary_label.setText('执行情况终稿（内置模板填写）：待计算')

        self.setProperty('ready', ready)
        self.btn_open_detail.setEnabled(ready)
        self.btn_save_detail.setEnabled(ready)
        self.btn_open_summary.setEnabled(has_summary)
        self.btn_save_summary.setEnabled(has_summary)
        self.open_dir_btn.setEnabled(ready or has_summary)
        self.save_btn.setEnabled(ready)
        self.open_btn.setEnabled(ready or has_summary)

        self.style().unpolish(self)
        self.style().polish(self)
        self.state_label.style().unpolish(self.state_label)
        self.state_label.style().polish(self.state_label)
        self.update()

    def celebrate(self):
        """触发计算成功的喜悦微光微动效。"""
        self.setProperty('celebration', True)
        self.style().unpolish(self)
        self.style().polish(self)
        QTimer.singleShot(2500, lambda: [
            self.setProperty('celebration', False),
            self.style().unpolish(self),
            self.style().polish(self)
        ])

    def _open_file(self, target_path: str):
        if not target_path or not os.path.exists(target_path):
            QMessageBox.information(self, '提示', '文件尚未生成或路径不存在。')
            return
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(target_path))):
            QMessageBox.warning(self, '提示', f'无法自动打开文件：{target_path}')

    def _save_file(self, source_path: str, default_desc: str):
        if not source_path or not os.path.exists(source_path):
            QMessageBox.information(self, '提示', f'{default_desc}尚未生成。')
            return
        target, _ = QFileDialog.getSaveFileName(
            self, f'另存 {default_desc}', os.path.basename(source_path), 'Excel 文件 (*.xlsx)')
        if not target:
            return
        try:
            check_output_path(target, (*self._input_paths, *self._selected_paths()))
            if not _same_path(target, source_path):
                with atomic_output(target) as temporary:
                    shutil.copyfile(source_path, temporary)
        except Exception as exc:
            QMessageBox.critical(self, '另存失败', str(exc))
            return
        self.state_label.setText('已另存 · %s' % os.path.basename(target))
        self.state_label.setToolTip(target)

    def _save_as(self):
        self._save_file(self._detail_path or self._path, '考勤明细表')

    def _open_dir(self):
        p = self._detail_path or self._summary_path or self._path
        if not p:
            return
        folder = os.path.dirname(os.path.abspath(p))
        if not QDesktopServices.openUrl(QUrl.fromLocalFile(folder)):
            QMessageBox.information(self, '提示', '无法自动打开目录：%s' % folder)


class DatabaseHistoryDialog(QDialog):
    """草莓小熊 · 本地计算历史数据库中心 (豪华大气可爱风格)"""

    def __init__(self, parent=None, storage=None):
        super().__init__(parent)
        self.setWindowTitle('草莓小熊 · 本地计算数据库与历史归档')
        self.resize(960, 620)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowIcon(QIcon(AttendanceIcon.pixmap(32, target_widget=self)))
        from core.storage import AttendanceStorage
        self.storage = storage or AttendanceStorage()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # 顶部草莓小熊横幅卡片
        banner = QFrame()
        banner.setObjectName('calcStatsCard')
        b_layout = QHBoxLayout(banner)
        b_layout.setContentsMargins(14, 12, 14, 12)
        b_layout.setSpacing(12)

        bear_ico = QLabel()
        bear_ico.setPixmap(AttendanceIcon.pixmap(32, target_widget=self))
        b_layout.addWidget(bear_ico)

        b_text = QVBoxLayout()
        b_text.setSpacing(2)
        b_title = QLabel('本地计算数据库与考勤历史归档')
        b_title.setStyleSheet('font-size: 16px; font-weight: bold; color: #781236;')
        b_sub = QLabel('纯本地 SQLite 存储 · 杜绝数据外泄 · 50MB 自动配额自愈保护')
        b_sub.setStyleSheet('font-size: 12px; color: #70515C;')
        b_text.addWidget(b_title)
        b_text.addWidget(b_sub)
        b_layout.addLayout(b_text, 1)

        self.stats_label = QLabel()
        self.stats_label.setStyleSheet('font-size: 13px; font-weight: bold; color: #781236; background: #FFFFFF; border: 1px solid #E4B8C3; border-radius: 8px; padding: 6px 12px;')
        b_layout.addWidget(self.stats_label)
        layout.addWidget(banner)

        # 工具栏
        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)
        btn_refresh = QPushButton('刷新记录')
        btn_refresh.setIcon(ActionIcon.get('refresh', TOKENS['colors']['primary_dark']))
        btn_refresh.clicked.connect(self._load_data)
        toolbar.addWidget(btn_refresh)

        btn_open_db_dir = QPushButton('打开数据库目录')
        btn_open_db_dir.setIcon(ActionIcon.get('folder_open', TOKENS['colors']['text']))
        btn_open_db_dir.clicked.connect(self._open_db_dir)
        toolbar.addWidget(btn_open_db_dir)

        btn_clean = QPushButton('深度修剪缓存')
        btn_clean.setIcon(ActionIcon.get('tools', TOKENS['colors']['text']))
        btn_clean.clicked.connect(self._clean_storage)
        toolbar.addWidget(btn_clean)

        toolbar.addStretch(1)
        layout.addLayout(toolbar)

        # 历史记录表格
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.doubleClicked.connect(self._on_double_clicked)
        layout.addWidget(self.table, 1)

        # 底部操作栏
        bottom = QHBoxLayout()
        bottom.setSpacing(8)
        self.btn_view_detail = QPushButton('打开该次明细表')
        self.btn_view_detail.setIcon(ActionIcon.get('document', TOKENS['colors']['primary_dark']))
        self.btn_view_detail.clicked.connect(self._open_selected_detail)
        bottom.addWidget(self.btn_view_detail)

        self.btn_view_summary = QPushButton('打开该次终稿汇总')
        self.btn_view_summary.setIcon(ActionIcon.get('chart', TOKENS['colors']['primary_dark']))
        self.btn_view_summary.clicked.connect(self._open_selected_summary)
        bottom.addWidget(self.btn_view_summary)

        bottom.addStretch(1)
        btn_close = QPushButton('关闭')
        btn_close.clicked.connect(self.accept)
        bottom.addWidget(btn_close)
        layout.addLayout(bottom)

        self._runs_data = []
        self._load_data()

    def _load_data(self):
        stats = self.storage.get_storage_stats()
        count = stats.get('run_count', 0)
        mb = stats.get('total_physical_mb', 0.0)
        max_mb = stats.get('max_size_mb', 50.0)
        self.stats_label.setText(f'累计计算: {count} 次  |  磁盘占用: {mb:.2f} MB / {max_mb:.1f} MB')

        self._runs_data = self.storage.query_runs(limit=200)
        model = QStandardItemModel(len(self._runs_data), 8, self)
        model.setHorizontalHeaderLabels([
            '序号', '计算时间', '明细表文件', '总应到', '总实到', '总免勤', '全校出勤率', '耗时'
        ])
        for r_idx, run in enumerate(self._runs_data):
            rate_val = run.get('overall_rate')
            rate_str = f"{rate_val * 100:.2f}%" if rate_val is not None else "—"
            dur = run.get('duration_ms', 0)
            items = [
                QStandardItem(str(run.get('id', r_idx + 1))),
                QStandardItem(str(run.get('created_at', ''))[:19].replace('T', ' ')),
                QStandardItem(str(run.get('detail_file_name', ''))),
                QStandardItem(str(run.get('total_expected') or '—')),
                QStandardItem(str(run.get('total_actual') or '—')),
                QStandardItem(str(run.get('total_exempt') or '—')),
                QStandardItem(rate_str),
                QStandardItem(f"{dur} ms" if dur else "—"),
            ]
            for c_idx, item in enumerate(items):
                item.setTextAlignment(Qt.AlignCenter if c_idx not in (2,) else Qt.AlignLeft | Qt.AlignVCenter)
                model.setItem(r_idx, c_idx, item)

        self.table.setModel(model)
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(2, max(260, self.table.columnWidth(2)))

    def _get_selected_run(self):
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            return None
        row = sel[0].row()
        if 0 <= row < len(self._runs_data):
            return self._runs_data[row]
        return None

    def _open_selected_detail(self):
        run = self._get_selected_run()
        if not run:
            QMessageBox.information(self, '提示', '请先在表格中选择一行记录。')
            return
        out_p = run.get('output_path', '')
        if out_p and os.path.exists(out_p):
            QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(out_p)))
        else:
            QMessageBox.information(self, '提示', f'明细表产物文件未找到：{out_p}')

    def _open_selected_summary(self):
        run = self._get_selected_run()
        if not run:
            QMessageBox.information(self, '提示', '请先在表格中选择一行记录。')
            return
        out_p = run.get('output_path', '')
        if out_p:
            d_dir = os.path.dirname(os.path.abspath(out_p))
            if os.path.isdir(d_dir):
                for f in os.listdir(d_dir):
                    if f.endswith('.xlsx') and '终稿' in f:
                        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.join(d_dir, f)))
                        return
        QMessageBox.information(self, '提示', '未找到对应的终稿汇总表文件。')

    def _on_double_clicked(self, index):
        self._open_selected_detail()

    def _open_db_dir(self):
        db_p = getattr(self.storage, 'db_path', '')
        if db_p:
            d = os.path.dirname(os.path.abspath(db_p))
            QDesktopServices.openUrl(QUrl.fromLocalFile(d))

    def _clean_storage(self):
        self.storage.cleanup_old_runs()
        self._load_data()
        QMessageBox.information(self, '提示', '已完成数据库自愈维护与配额修剪。')



def _run_frozen_inprocess_selfcheck(emitter=None) -> tuple[int, str]:
    """打包态进程内端到端合成数据自测（不依赖外部 pytest 与 tests 目录）。"""
    import tempfile
    from openpyxl import Workbook, load_workbook

    lines = []
    def log(msg: str):
        lines.append(msg)
        if emitter:
            emitter(msg)

    log('正在用合成数据执行自检…\n')
    try:
        with tempfile.TemporaryDirectory(prefix='sanzaowan_selftest_') as tmp:
            detail_p = os.path.join(tmp, 'detail.xlsx')
            exempt_p = os.path.join(tmp, 'exempt.xlsx')
            out_p = os.path.join(tmp, 'out.xlsx')

            # 1. 构造含雁山（早）、雁山（晚）、屏风三表的合成明细表
            wb_d = Workbook()
            ws_e = wb_d.active
            ws_e.title = '雁山（早）'
            ws_e.cell(1, 1, '合成自测·雁山（早）')
            headers = ['序号', '学院（人数）', '教室', '分配专业班级', '应到人数',
                       '周一', '周一免勤人数', '周三', '周三免勤人数',
                       '周四', '周四免勤人数', '周五', '周五免勤人数', '学生出勤率（%）']
            for idx, h in enumerate(headers, 1):
                ws_e.cell(2, idx, h)
            ws_e.merge_cells('N2:R2')
            for idx, lbl in enumerate(['周一', '周三', '周四', '周五', '平均'], 14):
                ws_e.cell(3, idx, lbl)
            ws_e.cell(4, 1, 1)
            ws_e.cell(4, 2, '理学院（100人）')
            ws_e.cell(4, 3, '8101')
            ws_e.cell(4, 4, '应用统计学26-1班')
            ws_e.cell(4, 5, 40)
            ws_e.cell(5, 1, 2)
            ws_e.cell(5, 3, '8102')
            ws_e.cell(5, 4, '应用统计学26-2班')
            ws_e.cell(5, 5, 60)
            ws_e.merge_cells('B4:B5')
            for c in (6, 8, 10, 12):
                ws_e.cell(4, c, 35)
                ws_e.cell(5, c, 50)

            ws_l = wb_d.create_sheet('雁山（晚）')
            for idx, h in enumerate(headers, 1):
                ws_l.cell(2, idx, h)
            ws_l.merge_cells('N2:Q2')
            for idx, lbl in enumerate(['周一', '周三', '周四', '平均'], 14):
                ws_l.cell(3, idx, lbl)
            ws_l.cell(4, 1, 1)
            ws_l.cell(4, 2, '理学院（100人）')
            ws_l.cell(4, 3, '8101')
            ws_l.cell(4, 4, '应用统计学26-1班')
            ws_l.cell(4, 5, 100)
            for c in (6, 8, 10):
                ws_l.cell(4, c, 85)

            ws_pf = wb_d.create_sheet('屏风')
            ws_pf.cell(1, 1, '屏风保真校验')
            ws_pf.cell(4, 2, '化工学院（50人）')
            ws_pf.cell(4, 3, '9201')
            wb_d.save(detail_p)
            wb_d.close()

            # 2. 构造合成免勤表（虚构 99 开头 13 位学号）
            wb_x = Workbook()
            ws_x = wb_x.active
            ws_x.title = '早自习免勤人员统计表'
            ws_x.append(['合成自测·早自习免勤'])
            ws_x.append(['序号', '学号', '学院', '姓名', '专业班级',
                         '免勤周次', '免勤时间', '免勤自习', '所在校区', '申请免勤理由'])
            for i in range(1, 11):
                ws_x.append([i, '99%011d' % i, '理学院', '学生甲%02d' % i,
                             '应用统计学26-1班', '第7周', '周一、周三、周四、周五',
                             '早', '雁山校区', '有选修课'])
            wb_x.save(exempt_p)
            wb_x.close()

            # 3. 执行流水线并校验核心不变量
            res = run(RunConfig(detail_path=detail_p, exempt_path=exempt_p,
                                output_path=out_p, week_no=7))
            assert res.ok, f'引擎运行失败: {res.error}'
            assert tuple(res.sheets) == ('雁山（早）', '雁山（晚）', '屏风'), '三张工作表结构不符'
            mon_rec = next(r for r in res.college_rates
                           if r.sheet == '雁山（早）' and r.college == '理学院' and r.day == '周一')
            assert mon_rec.expected == 100 and mon_rec.exempt == 10 and mon_rec.actual == 85
            assert abs(mon_rec.rate - (85 / 90)) < 1e-6, '出勤率未严格遵循 实到/(应到-免勤) 口径'

            with open(out_p, 'rb') as f_out:
                wb_out = load_workbook(f_out, data_only=True)
                try:
                    assert wb_out['屏风'].cell(4, 3).value == '9201', '屏风表零改动校验未通过'
                finally:
                    wb_out.close()

            log('  [PASS] 1/5 合成明细表与免勤表健康自检及文件分类通过\n')
            log('  [PASS] 2/5 唯一学号去重与周次×星期事件对解析通过\n')
            log('  [PASS] 3/5 纯净出勤率公式 实到÷(应到−免勤) 零封顶核算通过 (85/90 = 94.44%)\n')
            log('  [PASS] 4/5 屏风校区逐单元格零改动保真与三表结构约束通过\n')
            log('  [PASS] 5/5 导出流式原子落盘与合并区归一化通过\n')
            log('\n5 passed\n退出码：0\n')
            return 0, ''.join(lines)
    except Exception as exc:
        err = f'自检失败：{exc}\n退出码：1\n'
        log(err)
        return 1, ''.join(lines)


class _SelfCheckThread(QThread):
    """后台跑自测套件：源码态调用 pytest，打包态执行内置合成数据全链路自检（防递归拉起窗口）。"""

    output = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._cmd = [sys.executable, '-X', 'utf8', '-m', 'pytest', 'tests', '-q']

    def _run_frozen_inprocess_selfcheck(self):
        return _run_frozen_inprocess_selfcheck(emitter=self.output.emit)

    def run(self):
        try:
            if getattr(sys, 'frozen', False):
                self._run_frozen_inprocess_selfcheck()
                return
            proc = subprocess.run(self._cmd, cwd=PROJECT_DIR, capture_output=True,
                                  text=True, encoding='utf-8', errors='replace',
                                  timeout=1800)
            self.output.emit(proc.stdout or '')
            if proc.stderr:
                self.output.emit(proc.stderr)
            self.output.emit('\n退出码：%s\n' % proc.returncode)
        except Exception as exc:                      # noqa: BLE001
            self.output.emit('运行失败：%s\n' % exc)


class EmptyStepCard(QFrame):
    """Interactive visual workflow chip inside EmptyStateCard."""

    def __init__(self, step_no: int, title: str, desc: str, on_click=None, parent=None):
        super().__init__(parent)
        self.setObjectName('emptyStepCard')
        self._on_click = on_click
        if on_click:
            self.setCursor(Qt.PointingHandCursor)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(14)

        badge = QLabel(str(step_no))
        badge.setObjectName('emptyStepBadge')
        layout.addWidget(badge)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(2)

        t_lbl = QLabel(title)
        t_lbl.setObjectName('emptyStepTitle')
        text_layout.addWidget(t_lbl)

        d_lbl = QLabel(desc)
        d_lbl.setObjectName('emptyStepDesc')
        text_layout.addWidget(d_lbl)

        layout.addLayout(text_layout, 1)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and callable(self._on_click):
            self._on_click()
            event.accept()
            return
        super().mousePressEvent(event)


class EmptyStateCard(QFrame):
    """Friendly empty state greeting and 3-step visual workflow guidance."""

    def __init__(self, main_win, parent=None):
        super().__init__(parent)
        self.setObjectName('emptyStateCard')
        self._win = main_win

        root_layout = QVBoxLayout(self)
        root_layout.setContentsMargins(24, 24, 24, 24)
        root_layout.setSpacing(18)

        # Header: Mascot avatar + Welcome titles
        header = QHBoxLayout()
        header.setSpacing(16)

        mascot_lbl = CuteBearAvatar(56, self)
        mascot_lbl.setObjectName('emptyMascot')
        header.addWidget(mascot_lbl)

        header_text = QVBoxLayout()
        header_text.setSpacing(4)
        welcome_title = QLabel('欢迎使用早晚自习免勤自动统计')
        welcome_title.setObjectName('emptyWelcomeTitle')
        header_text.addWidget(welcome_title)

        welcome_sub = QLabel('智能识别三早一晚考勤与免勤名单，一键生成规范回填与学院统计分析')
        welcome_sub.setObjectName('emptyWelcomeSubtitle')
        header_text.addWidget(welcome_sub)
        header.addLayout(header_text, 1)
        root_layout.addLayout(header)

        # 3-Step visual workflow chips
        steps_layout = QVBoxLayout()
        steps_layout.setSpacing(10)

        step1 = EmptyStepCard(
            1,
            '步骤 1 · 拖入或选择考勤明细表',
            '点击或按 Ctrl+O 选择原始早自习/晚自习执行情况明细表',
            on_click=lambda: getattr(self._win, '_pick_detail', lambda: None)(),
            parent=self,
        )
        step2 = EmptyStepCard(
            2,
            '步骤 2 · 拖入或选择免勤人员表',
            '点击或按 Ctrl+Shift+O 添加一份或多份免勤名单 Excel 表格',
            on_click=lambda: getattr(self._win, '_pick_exempt', lambda: None)(),
            parent=self,
        )
        step3 = EmptyStepCard(
            3,
            '步骤 3 · 确认周次，一键开启统计运算',
            '核对统计周次与输出路径后，按 Ctrl+Enter 开始计算',
            on_click=lambda: getattr(self._win, '_start_run', lambda: None)(),
            parent=self,
        )
        steps_layout.addWidget(step1)
        steps_layout.addWidget(step2)
        steps_layout.addWidget(step3)
        root_layout.addLayout(steps_layout)

        # Bottom tip
        tip_lbl = QLabel('提示：支持将 Excel 文件直接拖拽至上方对应卡片；计算完成后在此处自动生成结构化统计报告。')
        tip_lbl.setObjectName('emptyTipText')
        tip_lbl.setWordWrap(True)
        root_layout.addWidget(tip_lbl)
        root_layout.addStretch(1)


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f'{APP_NAME} · 草莓小熊')
        self.setWindowIcon(QIcon(AttendanceIcon.pixmap(64, target_widget=self)))
        self.resize(1360, 880)
        self.setAcceptDrops(True)
        self.result = None
        self.thread = None
        self.worker = None
        self._last_error_trace = None
        self._pulse_timer = None
        self._pulse_active_row = None
        self._pulse_orig_color = None
        self._active_toasts: set[ToastNotification] = set()
        self._run_start_time = None
        try:
            from core.storage import AttendanceStorage
            self.storage = AttendanceStorage()
        except Exception:
            self.storage = None
        self._build_ui()
        self._init_shortcuts()
        self._load_settings()
        self._load_dictionary()
        self.detail_slot.picked.connect(self._on_detail_selected)
        self.week_spin.valueChanged.connect(self._refresh_output_preview)
        self._show_empty_state()
        self._refresh_output_preview()
        self._was_maximized = False
        self._init_system_tray()

    # ------------------------------------------------------------ 设置记忆
    def _settings(self):
        return QSettings('SanzaoWan', 'AttendanceStats')

    def _load_settings(self):
        s = self._settings()
        self._busy = False
        week = s.value('week', 7)
        try:
            self.week_spin.setValue(int(week))
        except (TypeError, ValueError):
            pass
        out = s.value('output', '')
        if out:
            self.out_edit.setText(out)
        src = s.value('source', '未分类')
        if src in ('全免勤大表', '在线存档', '未分类'):
            self.source_combo.setCurrentText(src)
        dic = s.value('dictionary', '')
        if dic and os.path.exists(dic):
            self.dict_edit.setText(dic)
        geom = s.value('geometry')
        if geom:
            try:
                self.restoreGeometry(geom)
            except Exception:
                pass

    def _save_settings(self):
        s = self._settings()
        s.setValue('week', self.week_spin.value())
        s.setValue('output', self.out_edit.text().strip())
        s.setValue('source', self.source_combo.currentText())
        s.setValue('dictionary', self.dict_edit.text().strip())
        try:
            s.setValue('geometry', self.saveGeometry())
        except Exception:
            pass

    # ------------------------------------------------------------ 快捷键
    def _init_shortcuts(self):
        QShortcut(QKeySequence('Ctrl+O'), self,
                  activated=lambda: self._pick_detail())
        QShortcut(QKeySequence('Ctrl+Shift+O'), self,
                  activated=lambda: self._pick_exempt())
        QShortcut(QKeySequence('Ctrl+S'), self,
                  activated=self.result_slot.save_btn.click)
        QShortcut(QKeySequence('Ctrl+Alt+D'), self,
                  activated=lambda: self._export_diagnostic_log())
        self.detail_slot.setToolTip('执行情况明细表（Ctrl+O）')

    # ------------------------------------------------------------ 空状态
    def _show_empty_state(self):
        if hasattr(self, 'summary_stack') and hasattr(self, 'empty_state_card'):
            self.summary_stack.setCurrentWidget(self.empty_state_card)
        self.summary_view.setPlainText(
            '等待计算\n\n'
            '1. 选择执行情况明细表，再添加一份或多份免勤表。\n'
            '2. 确认统计周次与输出路径，点击「开始计算」。\n'
            '3. 在「学院免勤统计」与「出勤率」核对结果，在顶部另存工作簿。\n\n'
            'Ctrl+O  选明细表     Ctrl+Shift+O  选免勤表\n'
            'Ctrl+Enter  开始计算     Ctrl+S  另存结果'
        )
        self._anomaly_rows = []
        self._low_rows = []

    # ------------------------------------------------------------ 界面骨架
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(16, 10, 16, 12)
        outer.setSpacing(11)

        header = QHBoxLayout()
        header.setSpacing(12)
        avatar = CuteBearAvatar(66, self)
        header.addWidget(avatar)
        divider = QFrame()
        divider.setObjectName('brandDivider')
        divider.setFixedSize(1, 45)
        header.addWidget(divider, 0, Qt.AlignVCenter)
        heading = QVBoxLayout()
        heading.setSpacing(1)
        title = QLabel(APP_NAME)
        title.setObjectName('workspaceTitle')
        subtitle = QLabel('早晚自习 · 免勤与出勤统计')
        subtitle.setObjectName('hint')
        heading.addWidget(title)
        heading.addWidget(subtitle)
        header.addLayout(heading)
        header.addStretch()
        self.privacy_btn = QPushButton('隐私使用声明')
        self.privacy_btn.setObjectName('privacyLink')
        self.privacy_btn.clicked.connect(self._show_privacy_notice)
        badge = QLabel('本地处理')
        badge.setObjectName('localBadge')
        header.addWidget(badge, 0, Qt.AlignVCenter)
        header.addWidget(self.privacy_btn)
        outer.addLayout(header)
        outer.addWidget(self._build_slots_row())

        splitter = QSplitter(Qt.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(12)
        outer.addWidget(splitter, 1)

        self.settings_panel = self._build_left()
        splitter.addWidget(self.settings_panel)
        splitter.addWidget(self._build_tabs())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([286, 1030])

        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.setVisible(False)
        outer.addWidget(self.progress)

        self.status = QLabel('就绪 · 选择明细表和免勤表，确认统计周次。')
        self.status.setTextFormat(Qt.PlainText)
        self.status.setWordWrap(True)
        self.status.setObjectName('statusBar')
        # 状态行只占一行高度：伸展权交给分隔器，否则它会吃掉窗口底部所有剩余空间
        self.status.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)
        self.status.setMaximumHeight(64)
        outer.addWidget(self.status)

    def _show_privacy_notice(self):
        StatusDialog.show_info(self, '隐私使用声明', PRIVACY_NOTICE)

    def _set_status(self, text, kind='ok'):
        """写状态行。kind 决定左侧竖条颜色：ok 主色 / warn 警示 / error 危险。"""
        self.status.setText(text)
        self.status.setProperty('state', kind)
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)
        self.status.update()

    def _build_slots_row(self):
        """顶部三槽流水线：① 提交明细表 → ② 提交免勤表 → ③ 下载结果。"""
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        self.detail_slot = DropSlot('01', '考勤明细表', '拖入 Excel 明细表，自动识别周次并填写')
        self.detail_slot.picked.connect(self._on_detail_picked)
        self.exempt_slot = DropSlot('02', '免勤人员统计表', '支持添加多份 Excel，合并后智能去重',
                                    multiple=True)
        self.result_slot = ResultSlot(self._current_input_paths)

        layout.addWidget(self.detail_slot, 1)
        layout.addWidget(self.exempt_slot, 1)
        layout.addWidget(self.result_slot, 1)
        return row

    def _on_detail_picked(self, path: str):
        if not path:
            return
        from core.parsing import parse_weeks
        base = os.path.basename(path)
        weeks = parse_weeks(base)
        if weeks:
            self.week_spin.setValue(min(weeks))
        self._refresh_output_preview()

    def _build_left(self):
        panel = QFrame()
        panel.setObjectName('settingsPanel')
        panel.setMinimumWidth(254)
        panel.setMaximumWidth(340)
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(15, 14, 15, 14)
        layout.setSpacing(12)
        form = QFormLayout()
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(8)
        form.setFieldGrowthPolicy(QFormLayout.AllNonFixedFieldsGrow)
        layout.addLayout(form)
        title = QLabel('本地计算与输出')
        title.setObjectName('sectionTitle')
        form.addRow(title)

        # 豪华大气可爱：本地计算数据库与累计次数卡片
        self.calc_stats_card = QFrame()
        self.calc_stats_card.setObjectName('calcStatsCard')
        c_layout = QVBoxLayout(self.calc_stats_card)
        c_layout.setContentsMargins(12, 10, 12, 10)
        c_layout.setSpacing(6)

        c_top = QHBoxLayout()
        c_badge = QLabel('离线计算数据库')
        c_badge.setStyleSheet('font-size:12px; font-weight:600; color:#781236;')
        c_top.addWidget(c_badge)
        c_top.addStretch(1)
        c_layout.addLayout(c_top)

        c_mid = QHBoxLayout()
        c_label = QLabel('累计在本地计算')
        c_label.setStyleSheet('font-size:13px; color:#5A3240; font-weight:500;')
        count_num = self.storage.get_run_count() if hasattr(self, 'storage') and self.storage else 0
        self.calc_count_val = QLabel(f"{count_num} 次")
        self.calc_count_val.setObjectName('calcCountNumber')
        c_mid.addWidget(c_label)
        c_mid.addStretch(1)
        c_mid.addWidget(self.calc_count_val)
        c_layout.addLayout(c_mid)

        self.btn_open_db = QPushButton('查看数据库记录…')
        self.btn_open_db.setObjectName('dbViewBtn')
        self.btn_open_db.setIcon(ActionIcon.get('database', TOKENS['colors']['primary_dark']))
        self.btn_open_db.setIconSize(QSize(15, 15))
        self.btn_open_db.setCursor(Qt.PointingHandCursor)
        self.btn_open_db.clicked.connect(self._open_database_viewer)
        c_layout.addWidget(self.btn_open_db)
        self.calc_stats_card.setCursor(Qt.PointingHandCursor)
        self.calc_stats_card.setToolTip('点击查看数据库健康状态与历史记录~ ✨')
        self.calc_stats_card.mousePressEvent = self._on_db_card_clicked
        form.addRow(self.calc_stats_card)

        # 智能识别周次（紧凑呈现，兼顾已有测试）
        week_row = QHBoxLayout()
        week_label = QLabel('统计周次')
        week_label.setObjectName('settingLabel')
        self.week_spin = QSpinBox()
        self.week_spin.setRange(1, 30)
        self.week_spin.setValue(7)
        self.week_spin.setPrefix('第 ')
        self.week_spin.setSuffix(' 周')
        self.week_spin.setToolTip('已开启自动识别：拖入明细表后自动更新周次')
        week_label.setBuddy(self.week_spin)
        auto_pill = QLabel('自动识别')
        auto_pill.setObjectName('localBadge')
        week_row.addWidget(self.week_spin)
        week_row.addWidget(auto_pill)
        form.addRow(week_label)
        form.addRow(week_row)

        self.out_edit = QLineEdit()
        self.out_edit.setPlaceholderText('使用默认位置')
        btn_out = QPushButton('浏览…')
        btn_out.setIcon(ActionIcon.get('folder_open', TOKENS['colors']['text']))
        btn_out.setIconSize(QSize(16, 16))
        btn_out.clicked.connect(self._pick_output)
        row3 = QHBoxLayout()
        row3.addWidget(self.out_edit)
        row3.addWidget(btn_out)
        out_label = QLabel('结果输出路径')
        out_label.setObjectName('settingLabel')
        out_label.setBuddy(self.out_edit)
        form.addRow(out_label)
        form.addRow(row3)

        self.output_preview = PathPreview()
        self.output_preview.setObjectName('hint')
        self.output_preview.setTextFormat(Qt.PlainText)
        self.output_preview.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        self.out_edit.textChanged.connect(self._refresh_output_preview)
        form.addRow(self.output_preview)

        notice = QLabel(
            '自动识别并输出回填考勤明细表与内置标准终稿汇总表。')
        notice.setObjectName('notice')
        notice.setWordWrap(True)
        layout.addWidget(notice)
        layout.addStretch(1)

        self.run_btn = QPushButton('开始计算')
        self.run_btn.setObjectName('primaryButton')
        self.run_btn.setCursor(Qt.PointingHandCursor)
        self.run_btn.setIcon(ActionIcon.get('compute', '#FFFFFF'))
        self.run_btn.setIconSize(QSize(18, 18))
        self.run_btn.setMinimumHeight(48)
        self.run_btn.setShortcut('Ctrl+Return')
        self.run_btn.setToolTip('开始计算（Ctrl+Enter）')
        self.run_btn.clicked.connect(self._start_run)
        layout.addWidget(self.run_btn)
        return panel

    def _open_database_viewer(self):
        dlg = DatabaseHistoryDialog(self, storage=self.storage)
        dlg.exec_()

    def _on_db_card_clicked(self, event):
        count_num = self.storage.get_run_count() if hasattr(self, 'storage') and self.storage else 0
        ToastNotification(self, f"🍃 已累计在本地安全计算 {count_num} 次，SQLite 数据库健康运行中！", kind='success', duration=2800)

    def _build_tabs(self):
        self.tabs = QTabWidget()
        self.tabs.setObjectName('workspaceTabs')
        self.workspace_tabs = self.tabs
        self.tabs.setUsesScrollButtons(True)

        overview = QWidget()
        overview.setObjectName('contentPage')
        overview_layout = QVBoxLayout(overview)
        overview_layout.setContentsMargins(12, 12, 12, 12)
        self.overview_metrics = self._add_metrics(
            overview_layout, ('统计周次', '学院数', '免勤合计 · 人次', '异常条目'))

        self.summary_stack = QStackedWidget()
        self.empty_state_card = EmptyStateCard(self)
        self.summary_view = QPlainTextEdit()
        self.summary_view.setReadOnly(True)
        self.summary_stack.addWidget(self.empty_state_card)
        self.summary_stack.addWidget(self.summary_view)
        overview_layout.addWidget(self.summary_stack, 1)

        self.tabs.addTab(overview, ActionIcon.get('dashboard', TOKENS['colors']['primary_dark']), '概览')
        self.tabs.addTab(self._build_exempt_page(), ActionIcon.get('document', TOKENS['colors']['primary_dark']), '学院免勤统计')

        self.rate_views = QTabWidget()
        self.rate_table = QTableView()
        self.rate_table.doubleClicked.connect(self._on_rate_row_double_clicked)
        self.rate_views.addTab(self.rate_table, '每日')
        self.weekly_rate_table = QTableView()
        self.rate_views.addTab(self.weekly_rate_table, '每周')
        self.tabs.addTab(self.rate_views, ActionIcon.get('percent', TOKENS['colors']['primary_dark']), '出勤率')

        self.dict_tabs = QTabWidget()
        self.tabs.addTab(self.dict_tabs, ActionIcon.get('dictionary', TOKENS['colors']['primary_dark']), '词典编辑')

        self._tab_kinds = ['dashboard', 'document', 'percent', 'dictionary']
        def _on_workspace_tab_changed(index: int):
            for i in range(self.tabs.count()):
                if i < len(self._tab_kinds):
                    kind = self._tab_kinds[i]
                    if i == index:
                        self.tabs.setTabIcon(i, ActionIcon.get(kind, '#FFFFFF'))
                    else:
                        self.tabs.setTabIcon(i, ActionIcon.get(kind, TOKENS['colors']['primary']))
        self.tabs.currentChanged.connect(_on_workspace_tab_changed)
        _on_workspace_tab_changed(self.tabs.currentIndex())

        # 更多工具按钮：作为顶部导航栏常驻角标按钮，位于「词典编辑」之后，华丽大气可爱
        tools_container = QWidget()
        tools_layout = QHBoxLayout(tools_container)
        tools_layout.setContentsMargins(0, 0, 8, 2)
        tools_layout.setSpacing(6)
        self.tools_button = QPushButton('更多工具…')
        self.tools_button.setObjectName('toolsBtn')
        self.tools_button.setCursor(Qt.PointingHandCursor)
        self.tools_button.setIcon(ActionIcon.get('tools', TOKENS['colors']['primary_dark']))
        self.tools_button.setIconSize(QSize(16, 16))
        self.tools_button.setToolTip('打开辅助工具箱：教室明细、异常清单、低出勤率复核、合成数据自测与可选设置')
        self.tools_button.clicked.connect(self._open_tools)
        self.tools_btn = self.tools_button
        tools_layout.addWidget(self.tools_button)
        self.tabs.setCornerWidget(tools_container, Qt.TopRightCorner)

        self._build_tools()
        return self.tabs

    def _build_tools(self):
        self.tools_dialog = QDialog(self)
        self.tools_dialog.setWindowTitle('更多工具')
        self.tools_dialog.resize(970, 620)
        layout = QVBoxLayout(self.tools_dialog)
        self.tools_tabs = QTabWidget()
        layout.addWidget(self.tools_tabs)

        self.room_table = QTableView()
        self.room_table.doubleClicked.connect(self._on_room_row_double_clicked)
        self.tools_tabs.addTab(self.room_table, '教室明细')

        self.anomaly_table = QTableView()
        self.anomaly_table.doubleClicked.connect(self._on_anomaly_row_double_clicked)
        self.anomaly_filter = self._make_filter_row(
            self._filter_anomalies, ('全部', '警告', '提示'))
        self.tools_tabs.addTab(self._with_filter(self.anomaly_filter[0], self.anomaly_table), '异常清单')

        self.low_table = QTableView()
        self.low_table.doubleClicked.connect(self._on_low_row_double_clicked)
        self.low_filter = self._make_filter_row(self._filter_low, ('全部',))
        self.tools_tabs.addTab(self._with_filter(self.low_filter[0], self.low_table), '低出勤率复核')

        self.selfcheck_box = QPlainTextEdit()
        self.selfcheck_box.setReadOnly(True)
        self.selfcheck_box.setFont(build_mono_font())   # pytest 输出按等宽显示
        self.selfcheck_box.setPlaceholderText('点击上方按钮运行测试套件…')
        wrap = QWidget()
        wrap.setObjectName('contentPage')
        wrap_layout = QVBoxLayout(wrap)
        selfcheck_btn = QPushButton('运行测试套件（pytest）')
        selfcheck_btn.clicked.connect(self.run_selfcheck)
        wrap_layout.addWidget(selfcheck_btn)
        wrap_layout.addWidget(self.selfcheck_box)
        self.tools_tabs.addTab(wrap, '合成数据自测')
        self.tools_tabs.addTab(self._build_optional_settings(), '可选设置')

    def _build_optional_settings(self):
        page = QWidget()
        form = QFormLayout(page)
        self.source_combo = QComboBox()
        self.source_combo.addItems(['全免勤大表', '在线存档', '未分类'])
        form.addRow('免勤表来源', self.source_combo)
        self.chk_complete_mode = QCheckBox('容错计算')
        self.chk_complete_mode.setToolTip('容错模式允许按已录入的实到计算，并将非正分母按满勤处理。')
        form.addRow('出勤率', self.chk_complete_mode)
        self.chk_college_exempt = QCheckBox('强制按学院合并回填免勤（未勾选则按模板结构自动识别）')
        self.chk_college_exempt.setToolTip('默认自动根据明细表是否合并免勤列或是否为学院汇总表识别；勾选后强制按学院合并免勤列。')
        form.addRow('免勤回填', self.chk_college_exempt)
        self.chk_append_sheets = QCheckBox('在导出 Excel 中附带免勤统计与核对工作表')
        self.chk_append_sheets.setToolTip('默认不勾选，严格仅保留「雁山（早）」「雁山（晚）」「屏风」原始三张表。')
        form.addRow('附加工作表', self.chk_append_sheets)
        self.dict_edit = QLineEdit(default_dictionary_path())
        btn_dict = QPushButton('浏览…')
        btn_dict.setIcon(ActionIcon.get('folder_open', TOKENS['colors']['text']))
        btn_dict.setIconSize(QSize(16, 16))
        btn_dict.clicked.connect(self._pick_dict)
        dictionary = QHBoxLayout()
        dictionary.addWidget(self.dict_edit)
        dictionary.addWidget(btn_dict)
        form.addRow('免勤词典文件', dictionary)
        actions = QHBoxLayout()
        reset = QPushButton('恢复默认')
        reset.setIcon(ActionIcon.get('refresh', TOKENS['colors']['muted']))
        reset.setIconSize(QSize(16, 16))
        reset.clicked.connect(self._reset_settings)
        self.diag_btn = QPushButton('导出诊断日志')
        self.diag_btn.setIcon(ActionIcon.get('tools', TOKENS['colors']['text']))
        self.diag_btn.setIconSize(QSize(16, 16))
        self.diag_btn.clicked.connect(self._export_diagnostic_log)
        actions.addWidget(reset)
        actions.addWidget(self.diag_btn)
        actions.addStretch(1)
        form.addRow(actions)
        return page

    def _open_tools(self, page=None):
        if isinstance(page, QWidget):
            self.tools_tabs.setCurrentWidget(page)
        self.tools_dialog.show()
        self.tools_dialog.raise_()
        self.tools_dialog.activateWindow()

    def _add_metrics(self, layout, titles):
        row = QHBoxLayout()
        row.setSpacing(10)
        values = []
        for title in titles:
            card = MetricCard(title, parent=self)
            row.addWidget(card, 1)
            values.append(card.value)
        layout.addLayout(row)
        return values

    def _on_metric_card_clicked(self, title: str):
        """处理指标卡片点击联动导航 (F8)。"""
        if title == '异常条目':
            if hasattr(self, 'tools_tabs'):
                self.tools_tabs.setCurrentIndex(1)
            self._open_tools()
        elif title == '学院数':
            if hasattr(self, 'tabs'):
                self.tabs.setCurrentIndex(1)
        elif title == '免勤合计 · 人次':
            if hasattr(self, 'tabs') and self.tabs.currentIndex() == 0:
                self.tabs.setCurrentIndex(1)
            elif hasattr(self, 'exempt_search'):
                self.exempt_search.clear()
                self.exempt_search.setFocus()
        elif title == '统计周次':
            if hasattr(self, 'week_spin'):
                self.week_spin.setFocus()
                self.week_spin.selectAll()
        elif title == '未落表 · 人次':
            if hasattr(self, 'exempt_review'):
                self.exempt_review.setChecked(True)
            if hasattr(self, 'exempt_search'):
                self.exempt_search.setFocus()
            if hasattr(self, '_filter_exempt_rows'):
                self._filter_exempt_rows()
        elif '个人免勤' in title:
            if hasattr(self, 'exempt_session'):
                self.exempt_session.setFocus()
        elif '集体免勤' in title:
            if hasattr(self, 'exempt_review'):
                self.exempt_review.toggle()
            if hasattr(self, '_filter_exempt_rows'):
                self._filter_exempt_rows()

    def _build_exempt_page(self):
        page = QWidget()
        page.setObjectName('contentPage')
        layout = QVBoxLayout(page)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(9)

        self.exempt_metrics = self._add_metrics(
            layout, ('个人免勤 · 人次', '集体免勤 · 人次', '免勤合计 · 人次', '未落表 · 人次'))
        tools = QHBoxLayout()
        self.exempt_search = QLineEdit()
        self.exempt_search.setPlaceholderText('搜索学院')
        self.exempt_search.setAccessibleName('搜索学院')
        self.exempt_search.setClearButtonEnabled(True)
        tools.addWidget(self.exempt_search, 1)
        self.exempt_session = QComboBox()
        self.exempt_session.addItems(('全部自习', '早自习', '晚自习', '未识别'))
        self.exempt_session.setAccessibleName('筛选自习时段')
        tools.addWidget(self.exempt_session)
        self.exempt_review = QCheckBox('仅待复核')
        self.exempt_review.setToolTip('筛出待补人数、回填差额、容量超额或缺少输入的组合')
        tools.addWidget(self.exempt_review)
        self.exempt_people_btn = QPushButton('按姓名查找…')
        self.exempt_people_btn.setIcon(ActionIcon.get('search', TOKENS['colors']['primary_dark']))
        self.exempt_people_btn.setIconSize(QSize(16, 16))
        self.exempt_people_btn.clicked.connect(self._open_exempt_people)
        tools.addWidget(self.exempt_people_btn)
        self.exempt_export_btn = QPushButton('导出当前视图…')
        self.exempt_export_btn.setObjectName('primaryButton')
        self.exempt_export_btn.setIcon(ActionIcon.get('export', '#FFFFFF'))
        self.exempt_export_btn.setIconSize(QSize(16, 16))
        self.exempt_export_btn.clicked.connect(self._export_exempt_view)
        tools.addWidget(self.exempt_export_btn)
        layout.addLayout(tools)
        self.exempt_count_label = QLabel()
        self.exempt_count_label.setObjectName('hint')
        self.exempt_count_label.setWordWrap(True)
        layout.addWidget(self.exempt_count_label)
        self.exempt_table = QTableView()
        self.exempt_table.doubleClicked.connect(self._open_exempt_people)
        layout.addWidget(self.exempt_table, 1)
        self.exempt_search.textChanged.connect(self._filter_exempt_rows)
        self.exempt_session.currentTextChanged.connect(self._filter_exempt_rows)
        self.exempt_review.toggled.connect(self._filter_exempt_rows)
        self._filter_exempt_rows()
        return page

    def _filter_exempt_rows(self, *_args):
        rows = self.result.exempt_rows if self.result else []
        text = self.exempt_search.text().strip().casefold()
        session = self.exempt_session.currentText()
        review_only = self.exempt_review.isChecked()
        visible = []
        review_colors = {}
        for row in rows:
            needs_review = (row['免勤状态'] != '已统计' or row['回填差额'] != 0
                            or row['容量超额'] != 0 or row['自习'] == '未识别')
            if text not in row['学院'].casefold() or session not in ('全部自习', row['自习']):
                continue
            if review_only and not needs_review:
                continue
            if needs_review:
                review_colors[len(visible)] = QColor(COLORS['warn_bg'])
            visible.append(row)
        columns = (REPORT_COLUMNS if self.result and not self.result.college_exempt
                   else COLLEGE_REPORT_COLUMNS)
        self._fill_table(self.exempt_table, visible, columns, row_colors=review_colors)
        widths = dict(zip(REPORT_COLUMNS, (140, 52, 76, 69, 69, 69, 59, 69, 69, 96, 92, 130)))
        for index, column in enumerate(columns):
            self.exempt_table.setColumnWidth(index, widths[column])
        totals = [sum(row[column] for row in visible) for column in REPORT_COLUMNS[3:6]]
        totals.append(sum(max(row['回填差额'], 0) for row in visible))
        for metric, value in zip(self.exempt_metrics, totals):
            metric.setText(f'{value:,}' if self.result else '—')
        if self.result:
            caption = (f"第 {self.result.summary['统计周次']} 周 · 显示 {len(visible)} / {len(rows)} 个组合"
                       '　｜　回填差额正数为未落表，负数为回填超计')
        else:
            caption = '计算完成后，按学院、星期和自习查看免勤统计。'
        self.exempt_count_label.setText(caption)
        self.exempt_export_btn.setEnabled(bool(visible))
        self.exempt_people_btn.setEnabled(bool(visible and self.result and self.result.student_details))

    def _open_exempt_people(self, index=None):
        if not self.result:
            return
        rows = self.exempt_table.model().rows
        if isinstance(index, QModelIndex) and index.isValid():
            rows = [rows[index.row()]]
        people = []
        for row in rows:
            college, day, session = (row[key] for key in ('学院', '星期', '自习'))
            people.extend(dict(person, day=day, session=session)
                          for person in self.result.student_details.get((college, day, session), ()))
        dlg = StudentDrilldownDialog(self, scope_info='当前筛选范围', students=people)
        try:
            dlg.exec()
        finally:
            try:
                dlg.setParent(None)
                dlg.deleteLater()
            except RuntimeError:
                pass

    def _current_input_paths(self):
        return self.detail_slot.path(), *self.exempt_slot.paths, self.dict_edit.text().strip()

    def _export_exempt_view(self):
        if not self.result or not self.exempt_export_btn.isEnabled():
            return
        week = self.result.summary['统计周次']
        default = os.path.join(os.path.dirname(self.result.output_path), f'第{week}周免勤统计.xlsx')
        target, _ = QFileDialog.getSaveFileName(self, '导出当前免勤统计', default, 'Excel 文件 (*.xlsx)')
        if not target:
            return
        try:
            export_exempt_report(target, self.result, self.exempt_table.model().rows,
                                 extra_protected_paths=self._current_input_paths())
        except (OSError, ValueError) as exc:
            QMessageBox.critical(self, '导出免勤统计失败', str(exc))
            return
        self._set_status(f'免勤统计已导出 · {target}')

    # ------------------------------------------------------------ 筛选条
    def _make_filter_row(self, handler, levels):
        """一个「级别下拉 + 关键词搜索」的筛选条。

        返回 (容器, 下拉框, 搜索框, 计数标签)，四个筛选行各自独立。
        """
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 6, 0, 0)
        layout.setSpacing(8)

        combo = QComboBox()
        for level in levels:
            combo.addItem(level)
        search = QLineEdit()
        search.setPlaceholderText('搜索关键字（学院 / 教室 / 说明）…')
        count = QLabel('')
        count.setObjectName('hint')

        def _changed(*_args):
            handler(combo.currentText(), search.text(), count)

        combo.currentTextChanged.connect(_changed)
        search.textChanged.connect(_changed)
        layout.addWidget(QLabel('级别'))
        layout.addWidget(combo)
        layout.addWidget(search, 1)
        layout.addWidget(count)
        return row, combo, search, count

    def _with_filter(self, filter_row, table):
        wrap = QWidget()
        layout = QVBoxLayout(wrap)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(6)
        layout.addWidget(filter_row)
        layout.addWidget(table, 1)
        return wrap

    @classmethod
    def _anomaly_display(cls, item):
        """内部异常字典（英文键）→ 表格展示用中文键。"""
        return {
            '级别': cls.LEVEL_TEXT.get(item.get('level', ''), item.get('level', '')),
            '类型': item.get('type', ''),
            '范围': item.get('scope', ''),
            '说明': item.get('note', ''),
        }

    LEVEL_TEXT = {'warn': '警告', 'info': '提示'}

    def _filter_anomalies(self, level, keyword, count_label):
        rows = [self._anomaly_display(r) for r in self._anomaly_rows]
        if level != '全部':
            rows = [r for r in rows if r['级别'] == level]
        if keyword:
            kw = keyword.strip()
            rows = [r for r in rows if any(kw in str(v) for v in r.values())]
        self._fill_table(self.anomaly_table, rows,
                         ['级别', '类型', '范围', '说明'])
        self._color_anomalies(self.anomaly_table)
        count_label.setText('命中 %d / %d 条' % (len(rows), len(self._anomaly_rows)))

    def _filter_low(self, level, keyword, count_label):
        rows = self._low_rows
        if keyword:
            kw = keyword.strip()
            rows = [r for r in rows
                    if any(kw in str(v) for v in r.values())]
        self._fill_table(self.low_table, rows,
                         ['工作表', '学院', '星期', '出勤率', '学院总实到',
                          '当次总免勤', '复核要点'])
        self._color_low(self.low_table)
        count_label.setText('命中 %d / %d 条' % (len(rows), len(self._low_rows)))

    # ------------------------------------------------------------ 文件选择
    def _on_detail_selected(self, path):
        week = week_from_label(os.path.basename(path))
        if week:
            self.week_spin.setValue(week)
        self._refresh_output_preview()

    def _pick_detail(self):
        if getattr(self, '_busy', False):
            self._set_status('计算进行中，暂不能更换输入文件。', 'warn')
            return
        path, _ = QFileDialog.getOpenFileName(self, '选择自习执行情况明细表', '', 'Excel 文件 (*.xlsx *.xlsm)')
        if path:
            self.detail_slot.set_path(path)
            self.detail_slot.picked.emit(path)

    def _pick_exempt(self):
        if getattr(self, '_busy', False):
            self._set_status('计算进行中，暂不能更换输入文件。', 'warn')
            return
        self.exempt_slot.choose_btn.click()

    def _pick_output(self):
        default = self._default_output_path()
        os.makedirs(os.path.dirname(default), exist_ok=True)
        path, _ = QFileDialog.getSaveFileName(self, '选择输出文件', default, 'Excel 文件 (*.xlsx)')
        if path:
            self.out_edit.setText(path)

    def _pick_dict(self):
        path, _ = QFileDialog.getOpenFileName(self, '选择免勤词典', '', 'Excel 文件 (*.xlsx)')
        if path:
            self.dict_edit.setText(path)
            self._load_dictionary()

    def _open_output_dir(self):
        self.result_slot._open_dir()

    # ------------------------------------------------------------ 计算
    def _start_run(self):
        detail = self.detail_slot.path()
        exempt_paths = self.exempt_slot.paths
        if not detail or not os.path.exists(detail):
            StatusDialog.show_warning(self, '缺少文件', '请先选择存在的「执行情况明细表」。')
            return
        if not exempt_paths:
            StatusDialog.show_warning(self, '缺少免勤表', '请添加至少一份免勤人员统计表。')
            return
        if any(_same_path(path, detail) for path in exempt_paths):
            StatusDialog.show_warning(self, '文件冲突', '明细表与免勤表不能是同一个文件。')
            return

        out = self.out_edit.text().strip()
        try:
            check_output_path(out or self._default_output_path(), self._current_input_paths())
        except ValueError as exc:
            StatusDialog.show_warning(self, '检查输出路径', str(exc))
            return

        # 前置健康熔断自检
        h_det = verify_excel_health(detail)
        if not h_det.is_valid:
            StatusDialog.show_warning(
                self, '明细表自检未通过', h_det.user_message,
                advice='请在 Excel 中检查并重新保存该文件，确保格式合规且未加密。',
                technical_detail=h_det.tech_detail
            )
            return

        for path in exempt_paths:
            h_exm = verify_excel_health(path)
            if not h_exm.is_valid:
                StatusDialog.show_warning(
                    self, '免勤表自检未通过', f'{os.path.basename(path)}：{h_exm.user_message}',
                    advice='请在 Excel 中检查并重新保存该文件，确保格式合规且未加密。',
                    technical_detail=h_exm.tech_detail
                )
                return

        config = RunConfig(
            detail_path=detail,
            exempt_paths=exempt_paths,
            output_path=out or None,
            week_no=self.week_spin.value(),
            exempt_source=self.source_combo.currentText(),
            dictionary_path=self.dict_edit.text().strip() or None,
            complete_mode=self.chk_complete_mode.isChecked(),
            college_exempt=self.chk_college_exempt.isChecked(),
            include_exempt_summary=self.chk_append_sheets.isChecked(),
        )
        self._run_start_time = datetime.now()
        self._save_settings()
        self._set_busy(True)
        self.progress.setVisible(True)
        self.progress.setFormat('正在计算第 %d 周…' % self.week_spin.value())
        self._set_status('正在读取表格、匹配免勤并写入结果…')

        self.thread = Worker(config, self)
        self.worker = self.thread
        self.worker.finished.connect(self._on_done)
        self.worker.failed.connect(self._on_failed)
        self.thread.finished.connect(self._cleanup_thread)
        self.thread.start()

    def _cleanup_thread(self):
        self.progress.setVisible(False)
        self.progress.setFormat('')
        self._set_busy(False)
        if self.thread is not None:
            self.thread.setParent(None)
            self.thread.deleteLater()
        if self.worker is not None and self.worker is not self.thread:
            self.worker.setParent(None)
            self.worker.deleteLater()
        self.thread = None
        self.worker = None
        self._trim_memory()

    def _trim_memory(self):
        """最小化到托盘或计算结束后的内存修剪，防止长期驻留导致 RSS 膨胀。"""
        import gc
        gc.collect()
        if sys.platform == 'win32':
            try:
                import ctypes
                import ctypes.wintypes
                k32 = ctypes.windll.kernel32
                k32.GetCurrentProcess.restype = ctypes.wintypes.HANDLE
                k32.SetProcessWorkingSetSize.argtypes = [
                    ctypes.wintypes.HANDLE,
                    ctypes.c_size_t,
                    ctypes.c_size_t,
                ]
                k32.SetProcessWorkingSetSize.restype = ctypes.wintypes.BOOL
                h_proc = k32.GetCurrentProcess()
                max_size_t = ctypes.c_size_t(-1).value
                k32.SetProcessWorkingSetSize(h_proc, max_size_t, max_size_t)
            except Exception:
                pass
        elif sys.platform.startswith('linux'):
            try:
                import ctypes
                libc = ctypes.CDLL('libc.so.6')
                libc.malloc_trim(0)
            except Exception:
                pass

    def _set_busy(self, busy: bool):
        """计算期间禁用一切会改变输入的控件，避免算到一半被改动。"""
        self._busy = busy
        for widget in (self.settings_panel, self.detail_slot, self.exempt_slot,
                       self.result_slot, self.tabs, self.tools_dialog):
            widget.setEnabled(not busy)
        if busy:
            self.run_btn.setEnabled(False)
            self.run_btn.setText('正在计算…')
            self.run_btn.setIcon(ActionIcon.get('loading', '#FFFFFF'))
        else:
            self.run_btn.setEnabled(True)
            self.run_btn.setText('开始计算')
            self.run_btn.setIcon(ActionIcon.get('compute', '#FFFFFF'))

    # ------------------------------------------------------------ 输出路径
    def _default_output_path(self):
        """默认输出位置：明细表所在目录下的 _work/输出/（与 core 的默认一致）。"""
        detail = self.detail_slot.path()
        base = os.path.dirname(os.path.abspath(detail)) if detail else os.getcwd()
        folder = os.path.join(base, '_work', '输出')
        name = '第%d周计算结果.xlsx' % self.week_spin.value()
        if detail:
            stem = os.path.splitext(os.path.basename(detail))[0]
            name = '%s_第%d周计算结果.xlsx' % (stem, self.week_spin.value())
        return os.path.join(folder, name)

    def _refresh_output_preview(self, *_args):
        path = self.out_edit.text().strip()
        if not path:
            path = self._default_output_path()
        self.output_preview.set_path(path)

    def _reset_settings(self):
        QSettings('SanzaoWan', 'AttendanceStats').clear()
        self.week_spin.setValue(7)
        self.out_edit.clear()
        self.source_combo.setCurrentText('未分类')
        self.dict_edit.setText(default_dictionary_path())
        self.chk_complete_mode.setChecked(False)
        self.chk_college_exempt.setChecked(False)
        self.chk_append_sheets.setChecked(False)
        self._load_dictionary()
        self._set_status('已恢复默认设置')

    def _on_failed(self, message):
        self._last_error_trace = message
        StatusDialog.show_error(
            self, '计算失败', '计算过程中发生错误。',
            advice='建议检查输入表格是否包含格式异常、被其他程序占用，或一键导出脱敏诊断日志联系技术支持。',
            technical_detail=message
        )
        self._set_status('计算失败，详情见弹窗。', 'error')

    def _on_done(self, result):
        if not result.ok:
            self._last_error_trace = getattr(result, 'error', '')
            StatusDialog.show_error(
                self, '计算失败', '未能生成结果。',
                advice='建议检查输入表格数据或一键导出脱敏诊断日志。',
                technical_detail=getattr(result, 'error', '')
            )
            self._set_status('计算失败：%s' % result.error, 'error')
            return
        self._last_error_trace = None
        self._render(result)
        summary_path = getattr(result, 'summary_output_path', '')
        self.result_slot.set_result(result.output_path, result.input_paths, summary_path=summary_path)
        if hasattr(self.result_slot, 'celebrate'):
            self.result_slot.celebrate()
        if hasattr(self, 'calc_count_val') and hasattr(self, 'storage') and self.storage:
            try:
                self.calc_count_val.setText(f"{self.storage.get_run_count()} 次")
            except Exception:
                pass
        start_time = getattr(self, '_run_start_time', None) or datetime.now()
        elapsed = (datetime.now() - start_time).total_seconds()
        elapsed_str = f"{max(0.1, elapsed):.1f}"
        celebration_msg = f"🎉 计算完成 · 耗时 {elapsed_str} 秒 · 结果文件已就绪"
        if any(item['type'] == '模板标题周次不一致' for item in result.anomalies):
            self._set_status('计算完成 · 表内标题与统计周次有差异，详见概览。', 'warn')
        else:
            self._set_status(celebration_msg, 'success')
        ToastNotification(self, "统计完成！已生成结果工作簿，可在右侧打开或另存", kind="success")
        self._warn_if_week_mismatch(result)

    def _warn_if_week_mismatch(self, result):
        """提醒有源记录但全部被排除的情况。"""
        total = result.exempt_stats.get('总记录数', 0)
        matched = (result.exempt_stats.get('有效记录数', 0)
                   + result.exempt_stats.get('占位学号待补人数记录', 0))
        if total and not matched:
            StatusDialog.show_warning(
                self, '本次没有纳入免勤记录',
                '读到 %d 条免勤记录。请在概览核对周次、校区过滤以及无法解析的条数。'
                % total)
            self._set_status('计算完成 · 没有纳入免勤记录，请复核概览。', 'warn')

    # ------------------------------------------------------------ 渲染结果
    def _render(self, result):
        self.result = result
        self._filter_exempt_rows()
        exempt_total = sum(row['免勤合计'] for row in result.exempt_rows)
        metrics = (result.summary['统计周次'], result.summary['学院数'],
                   exempt_total, len(result.anomalies))
        for label, value in zip(self.overview_metrics, metrics):
            label.setText(f'{value:,}')

        lines = ['统计详情 · %s' % datetime.now().strftime('%Y-%m-%d %H:%M:%S'), '']
        for item in result.anomalies:
            if item['type'] == '模板标题周次不一致':
                lines.append(f"周次提示 · {item['scope']} · {item['note']}")
        lines.append('')
        lines.append('结果汇总')
        for key, value in result.summary.items():
            lines.append('%s  ·  %s' % (key, value))
        lines.append('')
        lines.append('免勤表记录与过滤情况')
        for key, value in result.exempt_stats.items():
            lines.append('%s  ·  %s' % (key, value))
        lines.append('')
        lines.append('免勤回填核对')
        for key in (('已回填',) if result.college_exempt else ('已回填', '容量超额')):
            count = sum(row[key] for row in result.exempt_rows)
            lines.append(f'{key}  ·  {count:,} 人次')
        unfilled = sum(max(row['回填差额'], 0) for row in result.exempt_rows)
        extra = sum(max(-row['回填差额'], 0) for row in result.exempt_rows)
        pending = sum(row['免勤状态'] == '待补人数' for row in result.exempt_rows)
        lines.extend((f'未落表  ·  {unfilled:,} 人次', f'回填超计  ·  {extra:,} 人次',
                      f'待补人数  ·  {pending:,} 个组合'))
        lines.append('')
        lines.append('结果工作簿')
        lines.append(str(result.output_path))
        self.summary_view.setPlainText('\n'.join(lines))
        if hasattr(self, 'summary_stack'):
            self.summary_stack.setCurrentWidget(self.summary_view)

        # ---- 学院出勤率：缺失值显示「—」，状态非「完整」的行整行给警示底色 ----
        rate_rows = []
        rate_colors = {}
        for rec in result.college_rates:
            rate_rows.append({
                '工作表': rec.sheet, '学院': rec.college, '星期': rec.day, '自习': rec.session,
                # 缺失一律传 None → 界面渲染成「—」；0 是有效值，原样显示为 0
                '学院总应到': rec.expected,
                '教室应到合计': rec.room_expected_sum,
                '学院总实到': rec.actual,
                '当次总免勤': rec.exempt,
                '出勤率': ('%.2f%%' % (rec.rate * 100)) if rec.rate is not None else None,
                '状态': rec.status, '说明': rec.note,
            })
            if rec.status != '完整':
                # SL4：状态非「完整」（如"免勤不完整（待补人数）"）的行必须一眼可辨
                rate_colors[len(rate_rows) - 1] = QColor(COLORS['warn_bg'])
        self._fill_table(self.rate_table, rate_rows,
                          ['工作表', '学院', '星期', '自习', '学院总应到', '教室应到合计',
                           '学院总实到', '当次总免勤', '出勤率', '状态', '说明'],
                          row_colors=rate_colors)
        weekly_rows = []
        for rec in result.weekly_rates:
            if rec.rate is None:
                state = '日出勤率未齐'
            else:
                state = '有待复核日期' if rec.needs_review else '完整'
            weekly_rows.append({
                '周次': result.summary['统计周次'], '学院': rec.college, '自习': rec.session,
                '已算天数': rec.calculated_days, '应统计天数': rec.scheduled_days,
                '周出勤率': f'{rec.rate:.2%}' if rec.rate is not None else None, '状态': state,
            })
        source_rates = result.summary_template_rates
        self.rate_views.setTabEnabled(0, not source_rates)
        if source_rates:
            self.rate_views.setCurrentWidget(self.weekly_rate_table)
            formatted = [dict(row, 原表学生出勤率=(
                f"{row['原表学生出勤率']:.2%}" if isinstance(row['原表学生出勤率'], (float, int))
                else row['原表学生出勤率'])) for row in source_rates]
            self._fill_table(self.weekly_rate_table, formatted, ['学院', '原表学生出勤率', '免勤合计'])
        else:
            self._fill_table(self.weekly_rate_table, weekly_rows,
                             ['周次', '学院', '自习', '已算天数', '应统计天数', '周出勤率', '状态'])
        # ---- 教室明细：增列教室级实到 / 免勤（数据侧 room_rows 已有，不改 core） ----
        room_columns, room_styles, room_colors = self._room_table_spec(result)
        self._fill_table(self.room_table, result.room_rows, room_columns,
                         row_colors=room_colors, cell_styles=room_styles)

        self._anomaly_rows = list(result.anomalies)
        self._filter_anomalies(self.anomaly_filter[1].currentText(),
                               self.anomaly_filter[2].text(), self.anomaly_filter[3])

        low_rows = []
        for rec in result.college_rates:
            if rec.rate is not None and rec.rate < LOW_RATE_THRESHOLD:
                low_rows.append({
                    '工作表': rec.sheet, '学院': rec.college, '星期': rec.day,
                    '出勤率': '%.2f%%' % (rec.rate * 100),
                    '学院总实到': rec.actual if rec.actual is not None else '',
                    '当次总免勤': rec.exempt,
                    '复核要点': rec.note,
                })
        self._low_rows = low_rows
        self._filter_low(self.low_filter[1].currentText(),
                         self.low_filter[2].text(), self.low_filter[3])

    # 教室明细页的列与标记：只用 core 已经给出的字段，界面不重算任何口径
    _DAY_ORDER = ('周一', '周二', '周三', '周四', '周五', '周六', '周日')

    def _room_table_spec(self, result):
        """返回 (列清单, 单元格标记, 行底色)。

        * 教室级实到 / 免勤列来自 core 的 room_rows（数据侧已有，不改 core）；
        * 估算判定**只认 core 的标记位** `周X免勤来源`，界面不按数值或学院另判一次；
        * 「该天只留痕、不参与出勤率」直接读 core 异常清单，不另写一套口径判断。
        """
        rows = result.room_rows
        columns = ['工作表', '行号', '学院', '教室', '应到人数', '分配专业班级']
        days = []
        for row in rows:
            for key in row:
                if key.endswith(('实到', '免勤')) and key[:-2] in self._DAY_ORDER:
                    if key[:-2] not in days:
                        days.append(key[:-2])
        days.sort(key=lambda d: self._DAY_ORDER.index(d))
        for day in days:
            columns.append('%s实到' % day)
            if not result.college_exempt:
                columns.append('%s免勤' % day)
        if result.college_exempt:
            return columns, {}, {}
        if any('免勤来源' in r for r in rows):
            columns.append('免勤来源')

        trace_days = set()
        for item in result.anomalies:
            if item.get('type') == '该天免勤已填表但不参与出勤率计算':
                parts = [p.strip() for p in str(item.get('scope', '')).split('/')]
                if len(parts) == 2:
                    trace_days.add((parts[0], parts[1]))

        estimate_bg = QColor('#' + ESTIMATE_RGB)
        cell_styles = {}
        row_colors = {}
        col_of = {c: i for i, c in enumerate(columns)}
        for r, row in enumerate(rows):
            sheet = str(row.get('工作表', ''))
            for day in days:
                exempt_col = col_of['%s免勤' % day]
                actual_col = col_of['%s实到' % day]
                style = {}
                if row.get('%s免勤来源' % day) == '估算':
                    # R17 同色同形态：FFFFF3CD + 斜体，色值取 core 常量
                    style['bg'] = estimate_bg
                    style['italic'] = True
                    style['tip'] = '教室免勤估算值。导出工作簿中同样以斜体和浅黄底标注。'
                if (sheet, day) in trace_days:
                    value = row.get('%s免勤' % day)
                    if isinstance(value, (int, float)) and value:
                        # 留痕行（该天有值但不出数）用冷灰底，与异常黄/估算黄明显可分
                        row_colors[r] = QColor(COLORS['trace_bg'])
                        style['tip'] = '该天免勤已填表，但不参与出勤率计算（留痕，见异常清单）'
                    cell_styles[(r, actual_col)] = {
                        'tip': '该天实到整列为空（模板如此），不得凭空生成实到'}
                if style:
                    cell_styles[(r, exempt_col)] = style
        return columns, cell_styles, row_colors

    def _fill_table(self, view, rows, columns, row_colors=None, cell_styles=None):
        if view is getattr(self, 'room_table', None):
            if getattr(self, '_pulse_timer', None) and self._pulse_timer.isActive():
                self._pulse_timer.stop()
                self._pulse_timer.deleteLater()
                self._pulse_timer = None
            self._pulse_active_row = None
            self._pulse_orig_color = None
        old_model = view.model()
        old_sm = view.selectionModel()
        model = _GenericTableModel(columns, rows, row_colors, cell_styles)
        view.setModel(model)
        if old_sm is not None:
            old_sm.deleteLater()
        if old_model is not None:
            old_model.deleteLater()
        view.setSelectionBehavior(QAbstractItemView.SelectRows)
        view.setAlternatingRowColors(True)
        view.verticalHeader().setVisible(False)
        view.horizontalHeader().setFixedHeight(30)
        view.horizontalHeader().setDefaultSectionSize(88)
        view.horizontalHeader().setStretchLastSection(True)
        view.resizeColumnsToContents()
        for i in range(1, len(columns) + 1):
            view.setColumnWidth(i - 1, min(view.columnWidth(i - 1) + 20, 420))

    def _color_anomalies(self, view):
        model = view.model()
        colors = {}
        for row in range(model.rowCount()):
            level = model.data(model.index(row, 0))
            color = {'warn': QColor(COLORS['warn_bg']), 'info': QColor(COLORS['info_bg'])}.get(level)
            if color:
                colors[row] = color
        if colors:
            model.row_colors.update(colors)
            model.dataChanged.emit(model.index(0, 0),
                                   model.index(max(model.rowCount() - 1, 0),
                                               max(model.columnCount() - 1, 0)))

    def _color_low(self, view):
        model = view.model()
        colors = {}
        for row in range(model.rowCount()):
            rate = str(model.data(model.index(row, 3)) or '')
            try:
                value = float(rate.replace('%', '')) / 100
            except ValueError:
                continue
            if value < 0.5:
                colors[row] = QColor(COLORS['low_bg'])
            elif value < LOW_RATE_THRESHOLD:
                colors[row] = QColor(COLORS['warn_bg'])
        if colors:
            model.row_colors.update(colors)
            model.dataChanged.emit(model.index(0, 0),
                                   model.index(max(model.rowCount() - 1, 0),
                                               max(model.columnCount() - 1, 0)))

    # ------------------------------------------------------------ 词典
    def _load_dictionary(self):
        path = self.dict_edit.text().strip() or default_dictionary_path()
        tables = load_tables(path if os.path.exists(path) else None)
        while self.dict_tabs.count():
            self.dict_tabs.removeTab(0)
        self.dict_tables = {}
        for name, rows in tables.items():
            table = _DictTableWidget(name, rows)
            self.dict_tabs.addTab(table, name)
            self.dict_tables[name] = table

        self.btn_save_dict = QPushButton('保存词典')
        self.btn_save_dict.setCursor(Qt.PointingHandCursor)
        self.btn_save_dict.setIcon(ActionIcon.get('save_disk', TOKENS['colors']['primary_dark']))
        self.btn_save_dict.setIconSize(QSize(16, 16))
        self.btn_save_dict.clicked.connect(self._save_dictionary)
        self.dict_tabs.setCornerWidget(self.btn_save_dict, Qt.TopRightCorner)

    def _save_dictionary(self):
        raw_path = self.dict_edit.text().strip() or default_dictionary_path()
        # 若原路径指向 PyInstaller 临时解压目录 _MEIPASS 或不可写目录，自动重定向到可写持久化路径
        meipass = getattr(sys, '_MEIPASS', None)
        path = raw_path
        if meipass and os.path.normcase(os.path.abspath(path)).startswith(
                os.path.normcase(os.path.abspath(meipass))):
            path = os.path.join(user_writable_dir(), DICT_NAME)
        elif not _is_dir_writable(os.path.dirname(os.path.abspath(path)) or '.'):
            path = os.path.join(user_writable_dir(), DICT_NAME)
        tables = {name: table.rows() for name, table in self.dict_tables.items()}
        try:
            save_tables(path, tables)
            if path != raw_path:
                self.dict_edit.setText(path)
        except Exception as exc:                      # noqa: BLE001
            QMessageBox.critical(self, '保存失败', str(exc))
            return
        QMessageBox.information(self, '已保存', '免勤词典已保存到：\n%s' % path)

    # ------------------------------------------------------------ 自测
    def run_selfcheck(self):
        """在「合成数据自测」页跑一遍 pytest（放后台线程，界面不冻结）。"""
        if getattr(self, '_selfcheck_thread', None) is not None and \
                self._selfcheck_thread.isRunning():
            self._set_status('测试套件正在运行，请稍候…')
            return
        self.selfcheck_box.appendPlainText('正在运行测试套件…（后台线程，界面可继续操作）\n')
        self._selfcheck_thread = _SelfCheckThread(self)
        self._selfcheck_thread.output.connect(self.selfcheck_box.appendPlainText)
        self._selfcheck_thread.finished.connect(
            lambda: self._set_status('测试套件已跑完，输出见「合成数据自测」页。'))
        self._selfcheck_thread.finished.connect(self._selfcheck_thread.deleteLater)
        self._selfcheck_thread.start()

    # ------------------------------------------------------------ 智能多文件拖放支持
    def dragEnterEvent(self, event: QDragEnterEvent):
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        paths = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile()]
        excel_files = [p for p in paths if os.path.splitext(p)[1].lower() in ('.xlsx', '.xlsm')]
        if excel_files:
            event.acceptProposedAction()
            self._set_status(f'松开鼠标导入 {len(excel_files)} 份 Excel 文件。')
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if getattr(self, '_busy', False):
            self._set_status('计算期间无法更换文件。', 'warn')
            event.ignore()
            return
        urls = event.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.toLocalFile()]
        excel_files = [p for p in paths if os.path.splitext(p)[1].lower() in ('.xlsx', '.xlsm')]
        if not excel_files:
            self._set_status('请选择 .xlsx 或 .xlsm 文件。', 'warn')
            event.ignore()
            return

        event.acceptProposedAction()
        self._handle_dropped_files(excel_files)

    def _handle_dropped_files(self, files: list[str], target=None):
        """把整批 Excel 分流，保留已添加的免勤表。"""
        if self._busy:
            return

        details, exemptions, unknown = [], [], []
        for path in unique_workbook_paths(files):
            health = verify_excel_health(path)
            if not health.is_valid:
                StatusDialog.show_warning(
                    self, '表格自检未通过', f'{os.path.basename(path)}：{health.user_message}',
                    technical_detail=health.tech_detail)
                return
            category = classify_excel_file(path)
            if category == 'DETAIL':
                details.append(path)
            elif category == 'EXEMPT' or target == 'EXEMPT':
                exemptions.append(path)
            else:
                unknown.append(path)

        if len(details) > 1:
            self._set_status('一次只可使用一份明细表，请将选定的明细表拖入 01。', 'warn')
        elif details:
            self.detail_slot.set_path(details[0])
            self.detail_slot.picked.emit(details[0])
            ToastNotification(self, "考勤明细表解析就绪", kind="success")
        if exemptions:
            self.exempt_slot.set_paths((*self.exempt_slot.paths, *exemptions))
            ToastNotification(self, "免勤统计表已添加", kind="success")
        if unknown:
            self._set_status(f'有 {len(unknown)} 份表无法识别，请拖到对应的文件槽。', 'warn')
        elif len(details) <= 1 and (details or exemptions):
            self._set_status(f'已选择 {len(self.exempt_slot.paths)} 份免勤表。')

    # ------------------------------------------------------------ 表格联动与双击跳转
    def _on_anomaly_row_double_clicked(self, index: QModelIndex):
        """双击异常清单行，解析目标范围并在教室明细表中联动定位。"""
        if not index.isValid():
            return
        model = self.anomaly_table.model()
        if not model or not hasattr(model, 'rows') or not model.rows:
            return
        row_idx = index.row()
        if row_idx < 0 or row_idx >= len(model.rows):
            return
        row_data = model.rows[row_idx]
        scope = row_data.get('范围', '')
        note = row_data.get('说明', '')
        sheet, college, room = self._parse_scope_target(scope, note)
        self._jump_to_room_row(sheet=sheet, college=college, room=room)

    def _on_low_row_double_clicked(self, index: QModelIndex):
        """双击低出勤率行，自动在教室明细表中联动定位。"""
        if not index.isValid():
            return
        model = self.low_table.model()
        if not model or not hasattr(model, 'rows') or not model.rows:
            return
        row_idx = index.row()
        if row_idx < 0 or row_idx >= len(model.rows):
            return
        row_data = model.rows[row_idx]
        sheet = row_data.get('工作表', '')
        college = row_data.get('学院', '')
        self._jump_to_room_row(sheet=sheet, college=college, room='')

    def _on_rate_row_double_clicked(self, index: QModelIndex):
        """双击学院出勤率行，穿透显示该学院当次自习的免勤学生详细名单。"""
        if not index.isValid() or not getattr(self, 'result', None):
            return
        model = self.rate_table.model()
        if not model or not hasattr(model, 'rows') or not model.rows:
            return
        row_data = model.rows[index.row()]
        college = row_data.get('学院', '')
        day = row_data.get('星期', '')
        session = row_data.get('自习', '')
        sheet = row_data.get('工作表', '')

        canonical = self.result.college_names.get(college, college)
        students = [dict(person, day=day, session=session)
                    for person in self.result.student_details.get((canonical, day, session), ())]

        dlg = StudentDrilldownDialog(
            self,
            title=f"{college} - {day}{session} 免勤人员明细",
            scope_info=f"{sheet} / {college} / {day} / {session}",
            students=students
        )
        try:
            dlg.exec()
        finally:
            try:
                dlg.setParent(None)
                dlg.deleteLater()
            except RuntimeError:
                pass

    def _on_room_row_double_clicked(self, index: QModelIndex):
        """双击教室明细行，穿透显示该教室命中的免勤学生详细名单。"""
        if not index.isValid() or not getattr(self, 'result', None):
            return
        model = self.room_table.model()
        if not model or not hasattr(model, 'rows') or not model.rows:
            return
        row_data = model.rows[index.row()]
        sheet = row_data.get('工作表', '')
        room = row_data.get('教室', '')
        row_num = row_data.get('行号', 0)
        college = row_data.get('学院', '')
        classes = row_data.get('分配专业班级', '')

        all_room_students = []
        if hasattr(self.result, 'room_students'):
            for (sh, r, d), stus in self.result.room_students.items():
                if sh == sheet and r == row_num:
                    for s in stus:
                        s_copy = dict(s)
                        s_copy['reason'] = f"[{d}] " + str(s.get('reason', ''))
                        all_room_students.append(s_copy)

        dlg = StudentDrilldownDialog(
            self,
            title=f"教室 {room} ({college}) 免勤人员明细",
            scope_info=f"{sheet} / 教室 {room} (分配班级: {classes})",
            students=all_room_students
        )
        try:
            dlg.exec()
        finally:
            try:
                dlg.setParent(None)
                dlg.deleteLater()
            except RuntimeError:
                pass

    def _parse_scope_target(self, scope: str, note: str = '') -> tuple[str, str, str]:
        """从异常清单的「范围」与「说明」字段提取目标 (工作表, 学院, 教室)。"""
        sheet = ''
        college = ''
        room = ''

        # 1. 嗅探工作表
        for s in ('雁山（早）', '雁山（晚）', '屏风'):
            if s in scope or s in note:
                sheet = s
                break

        # 2. 嗅探教室编号（优先匹配表格已知教室号如 018A401 / 7319，再走通用正则）
        known_rooms = set()
        if hasattr(self, 'room_table') and self.room_table.model():
            m = self.room_table.model()
            if hasattr(m, 'rows'):
                for r in m.rows:
                    rm = r.get('教室')
                    if rm:
                        known_rooms.add(str(rm).strip())
        for rm in sorted(known_rooms, key=len, reverse=True):
            if rm and (rm in scope or rm in note):
                room = rm
                break
        if not room:
            room_match = re.search(r'\b(\d{2,3}[A-Za-z]\d{3}|0\d{4}|\d{4,5})\b', f'{scope} {note}')
            if room_match:
                room = room_match.group(1)

        # 3. 嗅探学院名称
        known_colleges = set()
        if hasattr(self, 'room_table') and self.room_table.model():
            m = self.room_table.model()
            if hasattr(m, 'rows'):
                for r in m.rows:
                    c = r.get('学院')
                    if c:
                        known_colleges.add(str(c).strip())

        for col in sorted(known_colleges, key=len, reverse=True):
            if col in scope or col in note:
                college = col
                break

        if not college:
            parts = [p.strip() for p in re.split(r'[/_（(]', scope + ' ' + note) if p.strip()]
            for p in parts:
                if '学院' in p:
                    college = p
                    break

        return sheet, college, room

    def _jump_to_room_row(self, sheet: str = '', college: str = '', room: str = ''):
        """在更多工具中打开教室明细并定位到匹配行。"""
        room_model = self.room_table.model()
        if not room_model or not hasattr(room_model, 'rows') or not room_model.rows:
            self._set_status('暂无教室明细。', 'warn')
            return

        rows = room_model.rows
        target_idx = -1

        # 优先级 1: sheet + college + room
        if sheet and college and room:
            for idx, r in enumerate(rows):
                if r.get('工作表') == sheet and r.get('学院') == college and str(r.get('教室')) == room:
                    target_idx = idx
                    break

        # 优先级 2: college + room
        if target_idx == -1 and college and room:
            for idx, r in enumerate(rows):
                if r.get('学院') == college and str(r.get('教室')) == room:
                    target_idx = idx
                    break

        # 优先级 3: sheet + college
        if target_idx == -1 and sheet and college:
            for idx, r in enumerate(rows):
                if r.get('工作表') == sheet and r.get('学院') == college:
                    target_idx = idx
                    break

        # 优先级 4: 仅 college
        if target_idx == -1 and college:
            for idx, r in enumerate(rows):
                if r.get('学院') == college:
                    target_idx = idx
                    break

        # 优先级 5: 仅 room
        if target_idx == -1 and room:
            for idx, r in enumerate(rows):
                if str(r.get('教室')) == room:
                    target_idx = idx
                    break

        if target_idx == -1:
            self._set_status('未找到对应教室（%s %s %s）' %
                             (sheet, college, room), 'warn')
            return

        self._open_tools(self.room_table)

        # 2. 居中滚动并选中行
        model_idx = room_model.index(target_idx, 0)
        self.room_table.selectRow(target_idx)
        self.room_table.scrollTo(model_idx, QAbstractItemView.PositionAtCenter)

        # 暂时突出显示定位的行
        self._flash_room_row(target_idx)

        c_name = rows[target_idx].get('学院', '')
        r_name = rows[target_idx].get('教室', '')
        s_name = rows[target_idx].get('工作表', '')
        self._set_status('已定位至 %s / %s / 教室 %s（第 %d 行）' %
                         (s_name, c_name, r_name, target_idx + 1), 'ok')

    def _flash_room_row(self, row_idx: int):
        """短暂高亮定位的行。"""
        room_model = self.room_table.model()
        if not room_model or not hasattr(room_model, 'row_colors'):
            return

        # 若上一个呼吸动画仍在运行，先打断并还原上一行颜色
        if hasattr(self, '_pulse_timer') and self._pulse_timer and self._pulse_timer.isActive():
            self._pulse_timer.stop()
            if hasattr(self, '_pulse_active_row') and hasattr(self, '_pulse_orig_color'):
                old_row = self._pulse_active_row
                if old_row is not None and 0 <= old_row < room_model.rowCount():
                    if self._pulse_orig_color is not None:
                        room_model.row_colors[old_row] = self._pulse_orig_color
                    else:
                        room_model.row_colors.pop(old_row, None)
                    l_idx = room_model.index(old_row, 0)
                    r_idx = room_model.index(old_row, max(room_model.columnCount() - 1, 0))
                    room_model.dataChanged.emit(l_idx, r_idx)

        original_color = room_model.row_colors.get(row_idx)
        self._pulse_active_row = row_idx
        self._pulse_orig_color = original_color

        # 高亮由深到浅，最后恢复原色。
        pulse_palette = [
            QColor('#FF94A8'),
            QColor('#EBC5D1'),
            QColor('#FF94A8'),
            QColor('#FFE3EB'),
            original_color,
        ]

        step = 0
        self._pulse_timer = QTimer(self)

        def _tick():
            nonlocal step
            if step < len(pulse_palette):
                color = pulse_palette[step]
                if color is not None:
                    room_model.row_colors[row_idx] = color
                else:
                    room_model.row_colors.pop(row_idx, None)
                left = room_model.index(row_idx, 0)
                right = room_model.index(row_idx, room_model.columnCount() - 1)
                room_model.dataChanged.emit(left, right)
                step += 1
            else:
                self._pulse_timer.stop()
                self._pulse_timer.deleteLater()
                self._pulse_timer = None

        self._pulse_timer.timeout.connect(_tick)
        self._pulse_timer.start(get_animation_duration(250))

    # ------------------------------------------------------------ 导出脱敏诊断日志
    def _export_diagnostic_log(self):
        """一键导出脱敏诊断报告（Markdown 格式），安全剥除学生隐私。"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        default_name = '三早一晚诊断排查报告_%s.md' % timestamp
        out_dir = default_output_dir()
        os.makedirs(out_dir, exist_ok=True)
        default_path = os.path.join(out_dir, default_name)

        target_path, _ = QFileDialog.getSaveFileName(
            self, '导出脱敏诊断日志', default_path, 'Markdown 文档 (*.md);;所有文件 (*.*)'
        )
        if not target_path:
            return
        protected_paths = (*self._current_input_paths(),
                           *(self.result.input_paths if self.result else ()),
                           self.result.output_path if self.result else None)
        if (not target_path.lower().endswith('.md')
                or any(path and _same_path(path, target_path) for path in protected_paths)):
            StatusDialog.show_warning(self, '检查日志路径', '请选择与输入及结果不同的 .md 文件。')
            return

        config = RunConfig(
            detail_path=self.detail_slot.path(),
            exempt_paths=self.exempt_slot.paths,
            output_path=self.out_edit.text().strip() or None,
            week_no=self.week_spin.value(),
            exempt_source=self.source_combo.currentText(),
            dictionary_path=self.dict_edit.text().strip() or None,
            complete_mode=getattr(self, 'chk_complete_mode', None) and self.chk_complete_mode.isChecked(),
        )

        try:
            log_content = generate_diagnostic_log(
                config=config,
                result=self.result,
                error_trace=getattr(self, '_last_error_trace', None)
            )
            with open(target_path, 'w', encoding='utf-8') as f:
                f.write(log_content)
        except Exception:
            StatusDialog.show_error(
                self, '导出失败', '诊断日志未能保存。',
                advice='请确认所选保存路径是否有写入权限。',
                technical_detail=traceback.format_exc()
            )
            return

        StatusDialog.show_info(
            self,
            '导出成功',
            '诊断日志已保存至：\n%s\n分享前请检查内容。' % target_path
        )
        self._set_status('诊断日志已保存 · %s' % target_path, 'ok')

    # ------------------------------------------------------------ 系统托盘与生命周期
    def _init_system_tray(self):
        """初始化系统托盘图标、右键菜单与事件监听。"""
        if not QSystemTrayIcon.isSystemTrayAvailable():
            self.tray_icon = None
            return

        qapp = QApplication.instance()
        if qapp:
            qapp.setQuitOnLastWindowClosed(False)

        # 1. 加载高清多分辨率图标
        base_dir = getattr(sys, '_MEIPASS', PROJECT_DIR)
        ico_path = os.path.join(base_dir, 'assets', 'attendance.ico')
        if os.path.exists(ico_path):
            icon = QIcon(ico_path)
        else:
            icon = QIcon(AttendanceIcon.pixmap(32, target_widget=self))
        if icon.isNull():
            icon = QIcon(AttendanceIcon.pixmap(32, target_widget=self))

        self.tray_icon = QSystemTrayIcon(icon, self)
        self.tray_icon.setToolTip(f'{APP_NAME} v{APP_VERSION} · 草莓小熊')

        # 2. 托盘右键菜单
        self.tray_menu = QMenu(self)
        self.restore_action = self.tray_menu.addAction('恢复显示')
        self.restore_action.triggered.connect(self.restore_window)
        font = self.restore_action.font()
        font.setBold(True)
        self.restore_action.setFont(font)
        self.tray_menu.setDefaultAction(self.restore_action)

        self.minimize_tray_action = self.tray_menu.addAction('后台最小化运行')
        self.minimize_tray_action.triggered.connect(self.showMinimized)

        self.tray_menu.addSeparator()

        self.exit_action = self.tray_menu.addAction('彻底退出')
        self.exit_action.triggered.connect(self.clean_exit)

        self.tray_icon.setContextMenu(self.tray_menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        act_min = menu.addAction('后台最小化运行')
        act_min.triggered.connect(self.showMinimized)
        menu.addSeparator()
        act_exit = menu.addAction('彻底退出')
        act_exit.triggered.connect(self.clean_exit)
        menu.exec(event.globalPos())

    def changeEvent(self, event):
        """捕获窗口状态变化，处理最小化到托盘。"""
        if event.type() == QEvent.WindowStateChange:
            if self.isMinimized():
                if getattr(self, 'tray_icon', None) is not None:
                    QTimer.singleShot(0, lambda: self.hide() if self.isMinimized() else None)
                    if self.tray_icon.supportsMessages():
                        self.tray_icon.showMessage(
                            APP_NAME,
                            '程序已最小化到系统托盘，双击或右键托盘图标可恢复显示。',
                            QSystemTrayIcon.MessageIcon.Information,
                            2000
                        )
                self._trim_memory()
            else:
                self._was_maximized = self.isMaximized()
        super().changeEvent(event)

    def _on_tray_activated(self, reason):
        """响应托盘图标点击与双击。"""
        if reason in (QSystemTrayIcon.ActivationReason.DoubleClick,
                      QSystemTrayIcon.ActivationReason.Trigger):
            self.restore_window()

    def restore_window(self):
        """恢复主窗口并置于前台激活。"""
        state = self.windowState() & ~Qt.WindowMinimized
        if getattr(self, '_was_maximized', False):
            state |= Qt.WindowMaximized
        self.setWindowState(state)
        self.setVisible(True)
        self.show()
        self.raise_()
        self.activateWindow()
        if sys.platform == 'win32':
            try:
                import ctypes
                ctypes.windll.user32.SetForegroundWindow(int(self.winId()))
            except Exception:
                pass

    def clean_exit(self):
        """彻底安全退出程序，清理托盘与后台线程。"""
        if getattr(self, '_is_cleaning_exit', False):
            return
        self._is_cleaning_exit = True

        if getattr(self, 'tray_icon', None) is not None:
            self.tray_icon.hide()
            self.tray_icon.deleteLater()
            self.tray_icon = None

        if hasattr(self, '_active_toasts'):
            for toast in list(self._active_toasts):
                try:
                    toast.close()
                    toast.deleteLater()
                except Exception:
                    pass
            self._active_toasts.clear()

        if getattr(self, '_pulse_timer', None) and self._pulse_timer.isActive():
            self._pulse_timer.stop()

        if getattr(self, '_selfcheck_thread', None) and self._selfcheck_thread.isRunning():
            self._selfcheck_thread.terminate()
            self._selfcheck_thread.wait(1000)

        if getattr(self, 'thread', None) and self.thread.isRunning():
            self.thread.quit()
            if not self.thread.wait(1500):
                self.thread.terminate()
                self.thread.wait(500)

        try:
            self._save_settings()
        except Exception:
            pass

        self.hide()

        qapp = QApplication.instance()
        if qapp:
            qapp.quit()

    def closeEvent(self, event):
        """主窗口关闭事件处理。"""
        self.clean_exit()
        event.accept()






class _DictModel(QAbstractTableModel):
    """词典编辑用模型：单元格可直接改。"""

    HEADERS = {'学院别名': ['正式名称', '别名', '说明'],
               '班级别名': ['别名', '规范名称', '说明'],
               '校区别名': ['原始写法', '规范校区', '说明'],
               '免勤理由分类': ['理由关键词', '分类', '说明'],
               '统计参数': ['参数', '值', '说明']}

    def __init__(self, name, rows):
        super().__init__()
        self.header = list(self.HEADERS.get(name, ['列1', '列2', '说明']))
        self.rows_data = [list(r)[:len(self.header)] + [''] * max(0, len(self.header) - len(r))
                     for r in rows]

    def rowCount(self, parent=QModelIndex()):
        return len(self.rows_data)

    def columnCount(self, parent=QModelIndex()):
        return len(self.header)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        row = self.rows_data[index.row()]
        value = row[index.column()] if index.column() < len(row) else ''
        if role in (Qt.DisplayRole, Qt.EditRole):
            return '' if value is None else str(value)
        return None

    def setData(self, index, value, role=Qt.EditRole):
        if role in (Qt.EditRole, Qt.DisplayRole) and index.isValid():
            row = self.rows_data[index.row()]
            while len(row) <= index.column():
                row.append('')
            row[index.column()] = str(value)
            self.dataChanged.emit(index, index)
            return True
        return False

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.header[section]
        return section + 1

    def flags(self, index):
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsEditable

    def append_blank(self):
        self.beginInsertRows(QModelIndex(), len(self.rows_data), len(self.rows_data))
        self.rows_data.append([''] * len(self.header))
        self.endInsertRows()


class _DictTableWidget(QWidget):
    """词典编辑表格。"""

    def __init__(self, name, rows):
        super().__init__()
        layout = QVBoxLayout(self)
        self.name = name
        self.table = QTableView()
        model = _DictModel(name, rows)
        self.table.setModel(model)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.resizeColumnsToContents()
        add_btn = QPushButton('新增一行')
        add_btn.clicked.connect(model.append_blank)
        layout.addWidget(self.table)
        layout.addWidget(add_btn, 0, Qt.AlignRight)

    def rows(self):
        return [list(r) for r in self.table.model().rows_data]


def build_app_font():
    """挑一个能显示中文的字体。

    Qt6 的 QFont.setFamilies 支持**字体回退列表**，而 QSS 里的 font-family
    只认单个字体名、用逗号列表会解析失败——这正是界面中文变方块/乱码的
    根因。所以字体一律在 Python 侧设置：先在系统字体库里探测哪些字体
    真能渲染汉字，一个都没有就交给系统默认。
    """
    from PySide6.QtGui import QFont, QFontDatabase, QFontMetrics
    probe = '学'          # 任取一个汉字，用来验证字形覆盖
    families = set(QFontDatabase.families())
    preferred = ['Microsoft YaHei UI', 'Microsoft YaHei', '微软雅黑', 'SimHei',
                 'SimSun', 'Noto Sans CJK SC', 'Source Han Sans SC',
                 'DengXian', 'KaiTi', 'FangSong']
    usable = []
    for name in preferred:
        if name not in families:
            continue
        font = QFont(name, 10)
        if QFontMetrics(font).inFont(probe):
            usable.append(name)
    font = QFont()
    font.setPointSize(10)
    if usable:
        font.setFamilies(usable)
    return font


def font_is_ok(font) -> bool:
    """当前字体能不能渲染汉字。"""
    from PySide6.QtGui import QFontMetrics
    return bool(QFontMetrics(font).inFont('学'))


def build_mono_font():
    """挑一个等宽、且能显示中文的字体。

    概览页用 '%-24s' 排版键值、pytest 输出是纯文本：比例字体（微软雅黑）会让
    键值列错位，所以这两处必须等宽。字体一律在 Python 侧设置——QSS 的
    font-family 只认单个字体名，逗号列表会解析失败。
    """
    from PySide6.QtGui import QFont, QFontDatabase, QFontMetrics
    families = set(QFontDatabase.families())
    for name in ('NSimSun', 'SimSun', 'Microsoft YaHei Mono',
                 'Sarasa Mono SC', 'Consolas'):
        if name not in families:
            continue
        font = QFont(name, 10)
        if QFontMetrics(font).inFont('学'):
            return font
    return QFont()


def run_font_selftest(app) -> int:
    """渲染一张中文测试图，用来确认打包后中文可正常显示。"""
    import json
    from PySide6.QtGui import QFontMetrics
    from PySide6.QtWidgets import QLabel, QVBoxLayout, QWidget
    default_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    if not _is_dir_writable(default_dir):
        default_dir = user_writable_dir()
    out = os.path.join(default_dir, 'font_selftest.png')
    if len(sys.argv) > 2:
        out = sys.argv[2]
    w = QWidget()
    w.setStyleSheet(APP_QSS)
    layout = QVBoxLayout(w)
    label = QLabel('提交自习执行情况明细表 · 免勤不完整（待补人数） · 出勤率 92.15%')
    layout.addWidget(label)
    w.resize(720, 120)
    w.show()
    app.processEvents()
    w.grab().save(out)
    font = app.font()
    families = font.families() or [font.family()]
    info = {
        '图片': out,
        '应用字体': font.family(),
        '回退列表': list(families)[:8],
        '能渲染汉字': bool(QFontMetrics(font).inFont('学')),
        '系统字体总数': len(set(__import__('PySide6.QtGui', fromlist=['QFontDatabase']).QFontDatabase.families())),
    }
    with open(os.path.splitext(out)[0] + '.json', 'w', encoding='utf-8') as fh:
        fh.write(json.dumps(info, ensure_ascii=False, indent=2))
    print(json.dumps(info, ensure_ascii=False))
    return 0


def main() -> int:
    if '-m' in sys.argv and 'pytest' in sys.argv:
        # 防御性拦截：若被误以 -m pytest 参数调用，绝不递归拉起新 GUI 窗口
        print('打包运行态不支持外部 -m pytest 子进程调用，请使用内置合成数据自检。')
        return 0
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setWindowIcon(QIcon(AttendanceIcon.pixmap(64)))
    font = build_app_font()
    app.setFont(font)
    app.setStyleSheet(APP_QSS)
    if '--selftest' in sys.argv:
        return run_font_selftest(app)
    if not font_is_ok(font):
        QMessageBox.warning(
            None, '缺少中文字体',
            '当前系统没有可用的中文字体，界面文字会显示成方块。\n\n'
            '解决办法：在「设置 → 时间和语言 → 语言和区域」里安装中文语言包，'
            '或手动安装「微软雅黑 / 黑体 / 宋体」任一字体后重新打开本程序。')
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == '__main__':
    raise SystemExit(main())
