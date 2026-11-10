# -*- coding: utf-8 -*-
"""自定义 UI 组件库。

包含：
- _GenericTableModel: 表格模型渲染
- CallableStr: 字符串/函数双重兼容类型
- PathPreview: 路径截断悬停预览
- ToastNotification: 浮动非阻塞微通知
- CuteBearAvatar: 草莓小熊互动头像
- MetricCard: 概览与指标卡片
- DropSlot: 拖拽导入与文件卡槽
- ResultSlot: 结果产物双卡片容器
- EmptyStepCard: 横向工作流指示卡（防字体压缩）
- EmptyStateCard: 空状态友好引导卡
"""

from __future__ import annotations

import os
import shutil
from typing import Optional

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QPointF,
    QRectF,
    QSize,
    QTimer,
    QUrl,
    Qt,
    Signal,
)
from PySide6.QtGui import (
    QBrush,
    QColor,
    QDesktopServices,
    QDragEnterEvent,
    QDropEvent,
    QFont,
    QGuiApplication,
    QPainter,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from core.excel_safety import classify_excel_file
from core.export import _same_path, atomic_output, check_output_path, unique_workbook_paths
from .icons import ActionIcon, AttendanceIcon
from .theme import TOKENS


def _format_file_size(size_bytes: int | float) -> str:
    """格式化字节大小为人类可读字符串。"""
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


def get_animation_duration(default_ms: int) -> int:
    """返回动效时长；在 offscreen / pytest 环境下返回 0ms 避免测试超时。"""
    if os.environ.get("QT_QPA_PLATFORM") == "offscreen":
        return 0
    try:
        if QGuiApplication.platformName() == "offscreen":
            return 0
    except Exception:
        pass
    return int(default_ms)


class _GenericTableModel(QAbstractTableModel):
    """把 dict 列表渲染成表格数据，供 QTableView 使用。"""

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


class CallableStr(str):
    """兼具字符串与可调用无参函数特性的兼容类型。"""

    def __call__(self) -> str:
        return str(self)


class PathPreview(QLabel):
    """在窄栏中保留路径两端，悬停时显示完整路径。"""

    def set_path(self, path: str):
        self.setToolTip(path)
        self.setText(self.fontMetrics().elidedText(path, Qt.ElideMiddle, max(50, self.contentsRect().width())))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.set_path(self.toolTip())


class ToastNotification(QFrame):
    """现代化浮动非阻塞轻量 Toast 通知。"""

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
    """可爱的草莓小熊互动头像：支持悬停提示、点击彩蛋反馈。"""

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
    """指标展示卡片，支持悬停提升与点击下钻跳转。"""

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


class DropSlot(QFrame):
    """文件提交槽位：支持点击选择与直接拖入。"""

    picked = Signal(str)

    def __init__(self, step: str, title: str, hint: str, *, multiple=False):
        super().__init__()
        self.setAcceptDrops(True)
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName('dropSlot')
        self.setProperty('dragOver', False)
        self.setProperty('ready', False)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.MinimumExpanding)
        self.setCursor(Qt.PointingHandCursor)
        self.paths = ()
        self.multiple = multiple
        self._hint = hint

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(6)

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
        self.path_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Expanding)
        self.path_label.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        layout.addWidget(self.path_label, 1)

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
        if len(paths) >= 2 or self.multiple:
            if hasattr(top_win, '_handle_dropped_files'):
                self.set_paths(self.paths)
                top_win._handle_dropped_files(paths, target='EXEMPT' if self.multiple else None)
                event.acceptProposedAction()
                return

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
        self.setProperty('justDropped', True)
        self._repolish()
        QTimer.singleShot(450, lambda: [
            self.setProperty('justDropped', False),
            self._repolish()
        ])


class ResultSlot(QFrame):
    """计算结果工作簿槽位：回填明细表与终稿汇总表双卡片操作。"""

    def __init__(self, selected_paths=lambda: ()):
        super().__init__()
        self.setFrameShape(QFrame.StyledPanel)
        self.setObjectName('dropSlot')
        self.setProperty('ready', False)
        self.setProperty('celebration', False)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.MinimumExpanding)
        self._path = ''
        self._detail_path = ''
        self._summary_path = ''
        self._input_paths = ()
        self._selected_paths = selected_paths

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)
        layout.setSpacing(6)

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

        self.state_label = QLabel('完成计算后可另存或打开目录')
        self.state_label.setObjectName('slotPath')
        self.state_label.setWordWrap(True)
        self.state_label.setTextFormat(Qt.PlainText)
        self.state_label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Fixed)
        self.state_label.setFixedHeight(28)
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

        # 产物 2：执行情况终稿（内置官方规范模板生成）
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


class EmptyStepCard(QFrame):
    """纵向工作流步骤指示卡（绝对高度锁定，杜绝纵向挤压与字体细绳）。"""

    def __init__(self, step_no: int, title: str, desc: str, on_click=None, parent=None):
        super().__init__(parent)
        self.setObjectName('emptyStepCard')
        self._on_click = on_click
        if on_click:
            self.setCursor(Qt.PointingHandCursor)

        self.setFixedHeight(48)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(14, 5, 14, 5)
        layout.setSpacing(12)

        badge = QLabel(str(step_no))
        badge.setObjectName('emptyStepBadge')
        badge.setFixedSize(24, 24)
        badge.setAlignment(Qt.AlignCenter)
        layout.addWidget(badge)

        text_layout = QVBoxLayout()
        text_layout.setContentsMargins(0, 0, 0, 0)
        text_layout.setSpacing(1)

        t_lbl = QLabel(title)
        t_lbl.setObjectName('emptyStepTitle')
        t_lbl.setTextFormat(Qt.PlainText)
        t_lbl.setFixedHeight(17)
        text_layout.addWidget(t_lbl)

        d_lbl = QLabel(desc)
        d_lbl.setObjectName('emptyStepDesc')
        d_lbl.setTextFormat(Qt.PlainText)
        d_lbl.setFixedHeight(15)
        text_layout.addWidget(d_lbl)

        layout.addLayout(text_layout, 1)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and callable(self._on_click):
            self._on_click()
            event.accept()
            return
        super().mousePressEvent(event)


class EmptyStateCard(QFrame):
    """友好空状态引导卡：草莓小熊问候 + 纵向三步流水线卡片（内置弹性滚动防护）。"""

    def __init__(self, main_win, parent=None):
        super().__init__(parent)
        self.setObjectName('emptyStateCard')
        self._win = main_win

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setSpacing(0)

        scroll = QScrollArea(self)
        scroll.setObjectName('emptyStateScroll')
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)

        container = QWidget()
        container.setObjectName('emptyStateScrollContent')
        root_layout = QVBoxLayout(container)
        root_layout.setContentsMargins(16, 12, 16, 12)
        root_layout.setSpacing(8)

        # 头部：吉祥物小熊 + 欢迎文案
        header = QHBoxLayout()
        header.setSpacing(12)

        mascot_lbl = CuteBearAvatar(46, container)
        mascot_lbl.setObjectName('emptyMascot')
        header.addWidget(mascot_lbl)

        header_text = QVBoxLayout()
        header_text.setSpacing(2)
        welcome_title = QLabel('欢迎使用早晚自习免勤自动统计')
        welcome_title.setObjectName('emptyWelcomeTitle')
        header_text.addWidget(welcome_title)

        welcome_sub = QLabel('智能识别三早一晚考勤与免勤名单，一键生成规范回填与学院统计分析')
        welcome_sub.setObjectName('emptyWelcomeSubtitle')
        header_text.addWidget(welcome_sub)
        header.addLayout(header_text, 1)
        root_layout.addLayout(header)

        # 纵向三步流水线引导卡片（定高 48px，严格保障上下非重叠与文字可读）
        steps_layout = QVBoxLayout()
        steps_layout.setSpacing(6)

        step1 = EmptyStepCard(
            1,
            '步骤 1 · 拖入或选择考勤明细表',
            '点击或按 Ctrl+O 选择原始早自习/晚自习执行情况明细表',
            on_click=lambda: getattr(self._win, '_pick_detail', lambda: None)(),
            parent=container,
        )
        step2 = EmptyStepCard(
            2,
            '步骤 2 · 拖入或选择免勤人员表',
            '点击或按 Ctrl+Shift+O 添加一份或多份免勤名单 Excel 表格',
            on_click=lambda: getattr(self._win, '_pick_exempt', lambda: None)(),
            parent=container,
        )
        step3 = EmptyStepCard(
            3,
            '步骤 3 · 一键开启智能统计运算',
            '全自动匹配免勤并回填，按 Ctrl+Enter 即刻生成明细表与终稿汇总表',
            on_click=lambda: getattr(self._win, '_start_run', lambda: None)(),
            parent=container,
        )
        steps_layout.addWidget(step1)
        steps_layout.addWidget(step2)
        steps_layout.addWidget(step3)
        root_layout.addLayout(steps_layout)

        # 底部提示语
        tip_lbl = QLabel('提示：支持将 Excel 文件直接拖拽至上方对应卡片；计算完成后在此处自动生成结构化统计报告。')
        tip_lbl.setObjectName('emptyTipText')
        tip_lbl.setWordWrap(True)
        root_layout.addWidget(tip_lbl)
        root_layout.addStretch(1)

        scroll.setWidget(container)
        outer_layout.addWidget(scroll)
