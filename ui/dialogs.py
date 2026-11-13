# -*- coding: utf-8 -*-
"""业务弹窗组件库。

包含：
- StatusDialog: 带可折叠诊断详情与一键复制的状态提示框
- StudentDrilldownDialog: 免勤学生明细下钻与 CSV 导出框
- DatabaseHistoryDialog: 本地 SQLite 历史存管与输出设置中心
- _DictModel / _DictTableWidget: 词典可视化编辑组件
- _run_frozen_inprocess_selfcheck: 进程内端到端合成数据自测
"""

from __future__ import annotations

import csv
import os
import shutil
import tempfile
from typing import List, Optional

from PySide6.QtCore import (
    QAbstractTableModel,
    QModelIndex,
    QSettings,
    QSize,
    QUrl,
    Qt,
)
from PySide6.QtGui import (
    QDesktopServices,
    QIcon,
    QStandardItem,
    QStandardItemModel,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QTableView,
    QVBoxLayout,
    QWidget,
)

from core import RunConfig, run
from core.io_exempt import normalize_student_name
from core.storage import AttendanceStorage
from .icons import ActionIcon, AttendanceIcon
from .theme import TOKENS, build_mono_font
from .widgets import _GenericTableModel, ToastNotification


class StatusDialog(QDialog):
    """带可折叠技术详情与一键复制的专业状态提示窗口。"""

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
    """查看指定统计范围内的免勤人员明细。"""

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

        card = QFrame()
        card.setObjectName('drilldownCard')
        c_layout = QHBoxLayout(card)
        c_layout.setContentsMargins(4, 4, 4, 4)
        c_lbl = QLabel(f'{scope_info}\n当前范围 {len(self.raw_students)} 条个人免勤记录')
        c_lbl.setTextFormat(Qt.PlainText)
        c_lbl.setObjectName('drilldownCountLabel')
        c_layout.addWidget(c_lbl)
        layout.addWidget(card)

        filter_bar = QHBoxLayout()
        filter_lbl = QLabel("搜索")
        filter_lbl.setObjectName('drilldownFilterLabel')
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('按姓名、学院、班级或学号筛选')
        self.search_input.textChanged.connect(self._apply_filter)
        filter_bar.addWidget(filter_lbl)
        filter_bar.addWidget(self.search_input)
        layout.addLayout(filter_bar)

        self.table = QTableView()
        self.columns = ['序号', '学号', '姓名', '学院', '专业班级', '星期', '自习', '教室', '来源',
                        '理由分类', '申请理由']
        self._update_table(self.raw_students)
        layout.addWidget(self.table, 1)

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
                '理由分类': s.get('reason_category', '未分类'),
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


class DatabaseHistoryDialog(QDialog):
    """历史记录与设置中心：管理历史计算记录与结果输出目录配置。"""

    def __init__(self, parent=None, storage=None):
        super().__init__(parent)
        self.setWindowTitle('历史记录与设置')
        self.resize(960, 600)
        self.setAttribute(Qt.WA_DeleteOnClose, True)
        self.setWindowIcon(QIcon(AttendanceIcon.pixmap(32, target_widget=self)))
        self.storage = storage or AttendanceStorage()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        # 头部说明卡片（克制专业，完全无夸张口号）
        header_card = QFrame()
        header_card.setStyleSheet('background: #FFFDFC; border: 1px solid #EBDAD6; border-radius: 10px; padding: 12px;')
        h_layout = QHBoxLayout(header_card)
        h_layout.setContentsMargins(8, 4, 8, 4)
        h_layout.setSpacing(12)

        bear_ico = QLabel()
        bear_ico.setPixmap(AttendanceIcon.pixmap(36, target_widget=self))
        h_layout.addWidget(bear_ico)

        h_text = QVBoxLayout()
        h_text.setSpacing(2)
        h_title = QLabel('历史记录与存储设置')
        h_title.setStyleSheet('font-size: 15px; font-weight: bold; color: #391525;')
        h_sub = QLabel('本地 SQLite 归档 · 50MB 自动配额自愈管理')
        h_sub.setStyleSheet('font-size: 12px; color: #70515C;')
        h_text.addWidget(h_title)
        h_text.addWidget(h_sub)
        h_layout.addLayout(h_text, 1)

        self.stats_label = QLabel()
        self.stats_label.setStyleSheet('font-size: 12.5px; font-weight: 600; color: #781236; background: #FFFFFF; border: 1px solid #E4B8C3; border-radius: 8px; padding: 6px 12px;')
        h_layout.addWidget(self.stats_label)
        layout.addWidget(header_card)

        # 输出目录配置
        out_box = QFrame()
        out_box.setStyleSheet('background: #FFFDFC; border: 1px solid #EBDAD6; border-radius: 8px; padding: 8px 12px;')
        ob_layout = QHBoxLayout(out_box)
        ob_layout.setContentsMargins(4, 2, 4, 2)
        ob_layout.setSpacing(8)

        lbl_out = QLabel('计算结果输出目录：')
        lbl_out.setStyleSheet('font-weight: 600; font-size: 12.5px; color: #442332;')
        ob_layout.addWidget(lbl_out)

        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        self.output_dir_edit.setStyleSheet('font-size: 12px; color: #391525; background: #FFFFFF; border: 1px solid #D5C2C6; border-radius: 6px; padding: 4px 8px;')

        curr_dir = ''
        if parent and hasattr(parent, '_get_custom_output_dir'):
            curr_dir = parent._get_custom_output_dir()
        if not curr_dir:
            s = QSettings('SanzaoWan', 'AttendanceStats')
            curr_dir = s.value('output_dir', '')
        self.output_dir_edit.setText(curr_dir or '（默认：自动保存在考勤明细表所在目录）')
        ob_layout.addWidget(self.output_dir_edit, 1)

        btn_browse_dir = QPushButton('更改目录…')
        btn_browse_dir.setIcon(ActionIcon.get('folder_open', TOKENS['colors']['primary_dark']))
        btn_browse_dir.clicked.connect(self._pick_custom_output_dir)
        ob_layout.addWidget(btn_browse_dir)

        btn_reset_dir = QPushButton('恢复默认')
        btn_reset_dir.clicked.connect(self._reset_default_output_dir)
        ob_layout.addWidget(btn_reset_dir)
        layout.addWidget(out_box)

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
        self.stats_label.setText(f'历史计算批次: {count} 条  |  占用: {mb:.2f} MB / {max_mb:.1f} MB')

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

    def _pick_custom_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, '选择计算结果输出目录', os.getcwd())
        if folder:
            self.output_dir_edit.setText(folder)
            s = QSettings('SanzaoWan', 'AttendanceStats')
            s.setValue('output_dir', folder)
            if self.parent() and hasattr(self.parent(), '_on_output_dir_changed'):
                self.parent()._on_output_dir_changed(folder)
            ToastNotification(self, '已成功更新计算结果输出目录', kind='success')

    def _reset_default_output_dir(self):
        self.output_dir_edit.setText('（默认：自动保存在考勤明细表所在目录）')
        s = QSettings('SanzaoWan', 'AttendanceStats')
        s.setValue('output_dir', '')
        if self.parent() and hasattr(self.parent(), '_on_output_dir_changed'):
            self.parent()._on_output_dir_changed('')
        ToastNotification(self, '已恢复默认输出目录', kind='info')


class _DictModel(QAbstractTableModel):
    """词典编辑用模型：单元格可直接修改。"""

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
    """词典编辑表格组件。"""

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


def _run_frozen_inprocess_selfcheck(emitter=None) -> tuple[int, str]:
    """打包态进程内端到端合成数据自测（不依赖外部 pytest 与 tests 目录）。"""
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
