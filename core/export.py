# -*- coding: utf-8 -*-
"""结果导出：在**原始明细表副本**上回填免勤人数与学院出勤率。

硬约束：
  * 只保留三张原始工作表：雁山（早）、雁山（晚）、屏风；不新增任何表。
  * 屏风表逐单元格零改动。
  * 绝不覆盖输入文件。
  * 出勤率写入同学院纵向合并单元格；合并区按学院行块归一化
    （修掉模板自带的 N69:N70 漏合并、O57:O59 与 N57:N58 不齐）。
  * 出勤率列加宽，避免 Excel 显示为 ####。
"""

from __future__ import annotations

import os
from contextlib import contextmanager
from tempfile import NamedTemporaryFile
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .calc import ESTIMATE_RGB, estimate_source
from .model import CollegeRate, DetailModel
from .parsing import strip_illegal

ALLOWED_SHEETS = ('雁山（早）', '雁山（晚）', '屏风')
RATE_WIDTH = 12.0


class ExportResult:
    def __init__(self):
        self.output_path: Optional[str] = None
        self.written_exempt_cells = 0
        self.written_rate_cells = 0
        self.merged_normalized = 0
        self.removed_sheets: List[str] = []
        self.notes: List[str] = []
        self.formulas_frozen = 0
        self.strings_sanitized = 0


def _same_path(a: str, b: str) -> bool:
    try:
        pa = os.path.normcase(os.path.realpath(os.path.abspath(a)))
        pb = os.path.normcase(os.path.realpath(os.path.abspath(b)))
        return pa == pb or os.path.samefile(a, b)
    except OSError:
        return False


def unique_workbook_paths(paths):
    selected = []
    for path in paths:
        if path and not any(_same_path(path, other) for other in selected):
            selected.append(os.fspath(path))
    return tuple(selected)


def check_output_path(path: str, protected_paths) -> None:
    if not str(path).lower().endswith('.xlsx'):
        raise ValueError('输出文件须使用 .xlsx 扩展名')
    if any(source and _same_path(source, path) for source in protected_paths):
        raise ValueError('输出路径不能与受保护的输入或结果文件相同')


@contextmanager
def atomic_output(path: str):
    """在目标目录创建独立临时文件，完整写入后替换目标。"""
    out_dir = os.path.dirname(os.path.abspath(path))
    os.makedirs(out_dir, exist_ok=True)
    with NamedTemporaryFile(dir=out_dir, prefix='.attendance-', suffix='.xlsx',
                            delete=False) as handle:
        tmp_path = handle.name
    try:
        yield tmp_path
        os.replace(tmp_path, path)
    finally:
        try:
            os.remove(tmp_path)
        except FileNotFoundError:
            pass


def _merge_ranges_in_col(ws, col: int):
    return [r for r in ws.merged_cells.ranges
            if r.min_col <= col <= r.max_col and r.max_col >= col]


def _college_cell(ws, lo, hi, column, result):
    for area in _merge_ranges_in_col(ws, column):
        if area.max_row < lo or area.min_row > hi:
            continue
        if area.min_col != column or area.max_col != column:
            raise ValueError(f'{ws.title} 的 {area} 合并跨越不同指标，无法安全回填')
        if (area.min_row, area.max_row) == (lo, hi):
            return ws.cell(lo, column)
        ws.unmerge_cells(str(area))
        result.merged_normalized += 1
    if hi > lo:
        ws.merge_cells(start_row=lo, start_column=column, end_row=hi, end_column=column)
        result.merged_normalized += 1
    return ws.cell(lo, column)


# 估算标记：教室级免勤是按比例分摊出来的估算值，必须在格子上留下固定痕迹
ESTIMATE_FILL = PatternFill(start_color=ESTIMATE_RGB, end_color=ESTIMATE_RGB,
                            fill_type='solid')
ESTIMATE_FONT = Font(italic=True)


def export(detail_path: str,
           output_path: str,
           detail: DetailModel,
           room_exempt: Dict[Tuple[str, int, str], int],
           rates: List[CollegeRate],
           average_rates: Dict[Tuple[str, str], float],
           estimated_cells=None,
           college_exempt=None, append_report=None) -> ExportResult:
    """生成结果文件。

    estimated_cells：{(工作表, 行号, 星期)}，这些格的免勤数是估算值，
    导出时加固定斜体 + 固定填充色（不改数字、不加批注、不加文字后缀）。
    """
    result = ExportResult()

    if _same_path(detail_path, output_path):
        raise ValueError('输出路径不能与输入明细表相同，以免覆盖原始文件')

    wb = load_workbook(detail_path)          # 保留公式与格式
    result.output_path = output_path

    # ---- 保真：把带缓存值的公式固化为数值 ----
    # openpyxl 不写公式的缓存值，若直接保存，所有公式格在按值读取时会变成空，
    # 既破坏原始数据，也让下游工具读到空值。这里用带缓存的值替换公式。
    frozen = 0
    sanitized = 0
    if college_exempt is None:
        wb_values = load_workbook(detail_path, data_only=True)
        for name in (name for name in wb.sheetnames if name in ALLOWED_SHEETS[:2]):
            ws_f = wb[name]
            ws_v = wb_values[name]
            for row in ws_f.iter_rows():
                for cell in row:
                    if cell.column in detail.protected_cols.get(name, ()):
                        continue
                    if cell.data_type == 'f':
                        cached = ws_v.cell(cell.row, cell.column).value
                        if cached is not None:
                            cell.value = cached
                            frozen += 1
                    elif isinstance(cell.value, str):
                        cleaned = strip_illegal(cell.value)
                        if cleaned != cell.value:
                            cell.value = cleaned
                            sanitized += 1
        wb_values.close()
    result.formulas_frozen = frozen
    result.strings_sanitized = sanitized

    # ---- 清理多余工作表（只留三张原表） ----
    for name in list(wb.sheetnames):
        if name not in ALLOWED_SHEETS:
            del wb[name]
            result.removed_sheets.append(name)

    avg_map = {(r.sheet, r.college, r.day): r for r in rates}

    for sheet_name in list(wb.sheetnames):
        if sheet_name not in ALLOWED_SHEETS or sheet_name.startswith('屏风'):
            continue
        ws = wb[sheet_name]
        rate_cols = detail.rate_cols_by_sheet.get(sheet_name, {})
        rate_label_row = detail.rate_label_rows.get(sheet_name)
        if rate_label_row:
            for label, column in rate_cols.items():
                ws.cell(rate_label_row, column).value = '周平均' if label == '平均' else label

        # ---- 1. 写教室级免勤人数 ----
        for room in detail.rooms_of(sheet_name) if college_exempt is None else ():
            for day, col in room.exempt_cols.items():
                if col is None:
                    continue
                value = room_exempt.get((sheet_name, room.row, day), 0)
                cell = ws.cell(room.row, col)
                cell.value = int(value)
                cell.number_format = '0'
                cell.alignment = Alignment(horizontal='center')
                if estimate_source(estimated_cells, (sheet_name, room.row, day)) == '估算':
                    # 估算值标记：固定斜体 + 固定填充色（数字本身保持纯整数）
                    # 判定来源唯一：core.calc.estimate_source（界面复用同一处）
                    cell.font = ESTIMATE_FONT
                    cell.fill = ESTIMATE_FILL
                result.written_exempt_cells += 1

        # ---- 2. 归一化出勤率区合并 + 写入出勤率 ----
        blocks = detail.blocks_of(sheet_name)
        written_colleges = set()
        for block in blocks:
            rows = block.rows if college_exempt is not None else block.merge_rows
            if not rows:
                continue
            lo, hi = min(rows), max(rows)
            if college_exempt is not None:
                for day, column in detail.exempt_cols_by_sheet.get(sheet_name, {}).items():
                    cell = _college_cell(ws, lo, hi, column, result)
                    value = college_exempt.get((sheet_name, block.name, day))
                    cell.value = value if block.name not in written_colleges else None
                    cell.number_format = '0'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    result.written_exempt_cells += cell.value is not None
                written_colleges.add(block.name)
            for label, col in rate_cols.items():
                if col is None:
                    continue
                rate_lo = max(lo, rate_label_row + 1) if rate_label_row else lo
                if rate_lo > hi:
                    continue
                cell = _college_cell(ws, rate_lo, hi, col, result)
                if label == '平均':
                    value = average_rates.get((sheet_name, block.name))
                else:
                    rec = avg_map.get((sheet_name, block.name, label))
                    value = rec.rate if rec else None
                if value is None:
                    cell.value = None
                else:
                    cell.value = float(value)
                    cell.number_format = '0.00%'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                result.written_rate_cells += 1

        # ---- 3. 列宽，避免 #### ----
        for label, col in rate_cols.items():
            if col:
                letter = get_column_letter(col)
                ws.column_dimensions[letter].width = RATE_WIDTH

    try:
        if append_report:
            append_report(wb)
        with atomic_output(output_path) as tmp_path:
            wb.save(tmp_path)
    finally:
        wb.close()
    return result
