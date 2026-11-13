# -*- coding: utf-8 -*-
"""UI 设计令牌、颜色常量、全局 QSS 样式表与字体引擎。"""

from __future__ import annotations

import os
from typing import List
from PySide6.QtGui import QFont, QFontDatabase, QFontMetrics

from core.calc import ESTIMATE_RGB          # R17 估算标记色（唯一事实源）
WARN_RGB = 'FFF6DC'                         # SL4 异常行警示色不变量


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

/* ---------- 四联排置顶流水线与卡片容器 ---------- */
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
QLabel#slotTitle { font-size:15px; font-weight:bold; color:%(text_strong)s; }
QLabel#slotPath { color:%(muted)s; background:#FFFFFF; border:1px solid %(border)s;
                  border-radius:6px; padding:4px 9px; }
QFrame#dropSlot[ready="true"] { border:1.5px solid #146F64; background:#FAFCFB; }
QFrame#dropSlot[ready="true"]:hover { border-color:#0E564D; background:#F2FAF8; }
QFrame#dropSlot[ready="true"] QLabel#slotPath { color:%(green_fg)s; font-weight:600; }
QFrame#dropSlot[celebration="true"] { border:1.5px solid %(green_fg)s; background:#F2FAF8; }
QFrame#brandDivider { background:#C78E9C; }
QLabel#workspaceTitle { font-size:26px; font-weight:bold; color:%(text_strong)s; }
QLabel#sectionTitle { font-size:16px; font-weight:bold; color:%(text_strong)s; }
QLabel#settingLabel { font-weight:600; color:%(text_strong)s; }
QFrame#metricCard { background:%(bg)s; border:1px solid %(border)s; border-radius:10px; }
QFrame#metricCard:hover { border:1px solid %(primary)s; background:#FFF9F7; }
QLabel#metricLabel { color:%(text)s; font-size:13px; font-weight:600; }
QLabel#metricValue { font-size:24px; font-weight:bold; color:%(text_strong)s; }
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

/* ---------- 豪华大气可爱：双表结果卡片与通用卡槽 ---------- */
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
QLabel#conflictBanner {
    color:%(warn_fg)s; background:#FFF6DC; border:1px solid #E3B95E;
    border-left:4px solid %(warn_fg)s; border-radius:8px; padding:8px 12px;
    font-size:12px; font-weight:600;
}

QStatusBar { background:%(card)s; color:%(text_strong)s; border-top:1px solid %(divider)s; }

/* ---------- 集中式对话框规则 ---------- */
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

/* ---------- 空状态体验 (防字体压缩细绳) ---------- */
QFrame#emptyStateCard { background:%(card)s; border:1.5px solid %(border)s; border-radius:12px; }
QScrollArea#emptyStateScroll, QWidget#emptyStateScrollContent { background:transparent; border:none; }
QFrame#emptyStepCard { background:%(bg)s; border:1px solid %(border_subtle)s; border-radius:10px; min-height:46px; }
QFrame#emptyStepCard:hover { background:%(primary_light)s; border-color:%(control)s; }
QLabel#emptyStepBadge { background:%(primary)s; color:#FFFFFF; font-weight:bold; font-size:13px;
                        border-radius:12px; min-width:24px; max-width:24px; min-height:24px; max-height:24px;
                        qproperty-alignment:AlignCenter; }
QLabel#emptyStepTitle { font-weight:bold; color:#391525; font-size:13px; }
QLabel#emptyStepDesc { color:#693C4C; font-size:12px; }
QLabel#emptyWelcomeTitle { font-size:16px; font-weight:bold; color:%(text_strong)s; }
QLabel#emptyWelcomeSubtitle { font-size:13px; color:%(muted)s; }
QLabel#emptyTipText { font-size:12px; color:%(slate)s; background:%(surface)s;
                      border:1px dashed %(border)s; border-radius:6px; padding:8px 12px; }

/* ---------- 浮动轻量通知 ---------- */
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


def build_app_font() -> QFont:
    """选取系统支持汉字渲染的标准字体。"""
    probe = '学'
    families = set(QFontDatabase.families())
    preferred = ['Microsoft YaHei UI', 'Microsoft YaHei', '微软雅黑', 'SimHei',
                 'SimSun', 'Noto Sans CJK SC', 'Source Han Sans SC',
                 'DengXian', 'KaiTi', 'FangSong']
    usable: List[str] = []
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


def font_is_ok(font: QFont) -> bool:
    """检验当前字体能否渲染汉字。"""
    return bool(QFontMetrics(font).inFont('学'))


def build_mono_font() -> QFont:
    """选取等宽且支持汉字的中英文字体。"""
    families = set(QFontDatabase.families())
    for name in ('NSimSun', 'SimSun', 'Microsoft YaHei Mono',
                 'Sarasa Mono SC', 'Consolas'):
        if name not in families:
            continue
        font = QFont(name, 10)
        if QFontMetrics(font).inFont('学'):
            return font
    return QFont()
