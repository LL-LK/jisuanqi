# -*- coding: utf-8 -*-
"""结果导出：在**原始明细表副本**上回填免勤人数与学院出勤率，并施加防篡改保护与移动端样式优化。

硬约束：
  * 只保留三张原始可见工作表：雁山（早）、雁山（晚）、屏风（另含 veryHidden 防篡改校验元数据表）。
  * 屏风表逐单元格零改动。
  * 绝不覆盖输入文件。
  * 出勤率写入同学院纵向合并单元格；合并区按学院行块归一化
    （修掉模板自带的 N69:N70 漏合并、O57:O59 与 N57:N58 不齐）。
  * 出勤率列加宽（>= 13.5），避免 Excel 及移动端显示为 ####。
"""

from __future__ import annotations

import hashlib
import hmac
import os
from contextlib import contextmanager
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Dict, List, Optional, Tuple, Union

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter

from .calc import ESTIMATE_RGB, estimate_source
from .model import CollegeRate, DetailModel
from .parsing import strip_illegal

ALLOWED_SHEETS = ('雁山（早）', '雁山（晚）', '屏风')
TAMPER_META_SHEET = '_系统校验元数据'
DEFAULT_SHEET_PASSWORD = 'sanzaowan2026'
TAMPER_HMAC_KEY = b'jisuanqi-sanzaowan-hmac-sha256-key-2026'
RATE_WIDTH = 13.5
MOBILE_FONT_NAME = 'Microsoft YaHei'
_THIN_SIDE = Side(style='thin', color='FF000000')
_MERGED_BORDER = Border(left=_THIN_SIDE, right=_THIN_SIDE, top=_THIN_SIDE, bottom=_THIN_SIDE)


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


def _ensure_merged_range_borders(ws, lo: int, hi: int, column: int) -> None:
    """确保合并区域内的每个子单元格都带有显式细边框，防止 Apple Numbers / 移动端表格边框断裂。"""
    for r in range(lo, hi + 1):
        sub_cell = ws.cell(r, column)
        b = sub_cell.border
        if b is None or b.left is None or b.left.style is None or b.right is None or b.right.style is None:
            sub_cell.border = Border(
                left=b.left if (b and b.left and b.left.style) else _THIN_SIDE,
                right=b.right if (b and b.right and b.right.style) else _THIN_SIDE,
                top=b.top if (b and b.top and b.top.style) else _THIN_SIDE,
                bottom=b.bottom if (b and b.bottom and b.bottom.style) else _THIN_SIDE,
            )


def _college_cell(ws, lo, hi, column, result):
    for area in _merge_ranges_in_col(ws, column):
        if area.max_row < lo or area.min_row > hi:
            continue
        if area.min_col != column or area.max_col != column:
            raise ValueError(f'{ws.title} 的 {area} 合并跨越不同指标，无法安全回填')
        if (area.min_row, area.max_row) == (lo, hi):
            _ensure_merged_range_borders(ws, lo, hi, column)
            return ws.cell(lo, column)
        ws.unmerge_cells(str(area))
        result.merged_normalized += 1
    if hi > lo:
        ws.merge_cells(start_row=lo, start_column=column, end_row=hi, end_column=column)
        result.merged_normalized += 1
    _ensure_merged_range_borders(ws, lo, hi, column)
    return ws.cell(lo, column)


# 估算标记：教室级免勤是按比例分摊出来的估算值，必须在格子上留下固定痕迹
ESTIMATE_FILL = PatternFill(start_color=ESTIMATE_RGB, end_color=ESTIMATE_RGB,
                            fill_type='solid')
ESTIMATE_FONT = Font(name=MOBILE_FONT_NAME, italic=True)


def _standardize_mobile_sheet_style(ws) -> None:
    """将工作表字体统一为跨平台移动端无衬线字体（Microsoft YaHei），同时保留原有的粗体/斜体/字号/颜色，并设置冻结窗格与网格线。"""
    ws.freeze_panes = 'D4'
    if ws.views and ws.views.sheetView:
        ws.views.sheetView[0].showGridLines = True

    # 确保 B 列已有的学院合并块也具备完整的子单元格边框
    for rng in _merge_ranges_in_col(ws, 2):
        if rng.min_col == 2 and rng.max_col == 2:
            _ensure_merged_range_borders(ws, rng.min_row, rng.max_row, 2)

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            f = cell.font
            if f is not None:
                if f.name != MOBILE_FONT_NAME:
                    cell.font = Font(
                        name=MOBILE_FONT_NAME,
                        size=f.size,
                        bold=f.bold,
                        italic=f.italic,
                        vertAlign=f.vertAlign,
                        underline=f.underline,
                        strike=f.strike,
                        color=f.color,
                    )
            else:
                cell.font = Font(name=MOBILE_FONT_NAME)


def apply_sheet_protection(
    ws,
    password: str = DEFAULT_SHEET_PASSWORD,
    detail: Optional[DetailModel] = None,
) -> None:
    """启用 OOXML 工作表级密码保护并锁定关键考勤数据列（应到人数、实到人数、免勤人数、出勤率）。"""
    locked_prot = Protection(locked=True)
    unlocked_prot = Protection(locked=False)

    sheet_name = ws.title
    locked_cols = set(range(1, 6))  # 序号、学院、教室、专业班级、应到人数 (Cols 1..5)
    if detail is not None:
        for room in detail.rooms_of(sheet_name):
            for col in room.actual_cols.values():
                if col:
                    locked_cols.add(col)
            for col in room.exempt_cols.values():
                if col:
                    locked_cols.add(col)
        for col in detail.exempt_cols_by_sheet.get(sheet_name, {}).values():
            if col:
                locked_cols.add(col)
        for col in detail.rate_cols_by_sheet.get(sheet_name, {}).values():
            if col:
                locked_cols.add(col)
        for col in detail.protected_cols.get(sheet_name, ()):
            if col:
                locked_cols.add(col)
    else:
        locked_cols.update(range(1, ws.max_column + 1))

    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            if r <= 3 or c in locked_cols:
                cell.protection = locked_prot
            else:
                cell.protection = unlocked_prot

    ws.protection.sheet = True
    if password:
        ws.protection.set_password(password)


def _canonical_cell_value(val) -> str:
    if val is None:
        return ''
    if isinstance(val, bool):
        return 'TRUE' if val else 'FALSE'
    if isinstance(val, int):
        return f'I:{val}'
    if isinstance(val, float):
        if val.is_integer():
            return f'I:{int(val)}'
        return f'F:{val:.8f}'
    return f'S:{str(val).strip()}'


def _compute_cell_hmac(sheet_name: str, row: int, col: int, val_repr: str, secret_key: bytes = TAMPER_HMAC_KEY) -> str:
    payload = f'{sheet_name}|{row}|{col}|{val_repr}'.encode('utf-8')
    return hmac.new(secret_key, payload, hashlib.sha256).hexdigest()


def generate_tamper_metadata(
    wb: Workbook,
    detail: Optional[DetailModel] = None,
    rates: Optional[List[CollegeRate]] = None,
    secret_key: bytes = TAMPER_HMAC_KEY,
):
    """生成隐藏防篡改校验工作表 `_系统校验元数据`（veryHidden），写入关键单元格 HMAC-SHA256 签名。"""
    if TAMPER_META_SHEET in wb.sheetnames:
        del wb[TAMPER_META_SHEET]

    ws_meta = wb.create_sheet(TAMPER_META_SHEET)
    ws_meta.sheet_state = 'veryHidden'
    ws_meta.append(['sheet_name', 'row', 'col', 'val_repr', 'hmac_sha256'])

    for sheet_name in ('雁山（早）', '雁山（晚）'):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(r, c).value
                val_repr = _canonical_cell_value(val)
                sig = _compute_cell_hmac(sheet_name, r, c, val_repr, secret_key=secret_key)
                ws_meta.append([sheet_name, r, c, val_repr, sig])

    return ws_meta


def verify_workbook_integrity(
    file_path_or_wb: Union[str, Path, Workbook],
    secret_key: bytes = TAMPER_HMAC_KEY,
) -> Tuple[bool, List[str]]:
    """校验导出工作簿的防篡改签名完整性，使用 hmac.compare_digest 恒定时间比对。"""
    close_after = False
    if isinstance(file_path_or_wb, Workbook):
        wb = file_path_or_wb
    else:
        wb = load_workbook(file_path_or_wb, data_only=True)
        close_after = True

    violations: List[str] = []
    try:
        if TAMPER_META_SHEET not in wb.sheetnames:
            return False, [f'缺失防篡改校验元数据表: {TAMPER_META_SHEET}']

        ws_meta = wb[TAMPER_META_SHEET]
        checked_rows = 0
        for row_idx in range(2, ws_meta.max_row + 1):
            sheet_name = ws_meta.cell(row_idx, 1).value
            r = ws_meta.cell(row_idx, 2).value
            c = ws_meta.cell(row_idx, 3).value
            expected_repr = ws_meta.cell(row_idx, 4).value or ''
            expected_sig = ws_meta.cell(row_idx, 5).value or ''

            if not sheet_name or r is None or c is None:
                continue
            r_int, c_int = int(r), int(c)
            checked_rows += 1

            if sheet_name not in wb.sheetnames:
                violations.append(f'工作表缺失: {sheet_name}')
                continue

            ws = wb[sheet_name]
            actual_val = ws.cell(r_int, c_int).value
            actual_repr = _canonical_cell_value(actual_val)
            actual_sig = _compute_cell_hmac(sheet_name, r_int, c_int, actual_repr, secret_key=secret_key)

            if not hmac.compare_digest( str(actual_sig), str(expected_sig) ):
                coord = f'{get_column_letter(c_int)}{r_int}'
                violations.append(
                    f'{sheet_name}!{coord} 数据被篡改（原始摘要值={expected_repr}, 当前值={actual_repr}）'
                )

        if checked_rows == 0:
            violations.append('防篡改元数据表为空，未包含有效签名记录')

        return len(violations) == 0, violations
    finally:
        if close_after:
            wb.close()


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

    with open(detail_path, 'rb') as f_detail:
        wb = load_workbook(f_detail)          # 保留公式与格式
    result.output_path = output_path

    try:
        # ---- 保真：把带缓存值的公式固化为数值 ----
        # openpyxl 不写公式的缓存值，若直接保存，所有公式格在按值读取时会变成空，
        # 既破坏原始数据，也让下游工具读到空值。这里用带缓存的值替换公式。
        frozen = 0
        sanitized = 0
        if college_exempt is None:
            with open(detail_path, 'rb') as f_values:
                wb_values = load_workbook(f_values, data_only=True)
            try:
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
            finally:
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

            # ---- 4. 移动端兼容排版（字体归一化、冻结窗格、显示网格线、合并区边框） ----
            _standardize_mobile_sheet_style(ws)

            # ---- 5. 工作表级防篡改密码锁定 ----
            apply_sheet_protection(ws, password=DEFAULT_SHEET_PASSWORD, detail=detail)

        if append_report:
            append_report(wb)

        # ---- 6. 写入隐藏防篡改校验元数据表 `_系统校验元数据` ----
        generate_tamper_metadata(wb, detail=detail, rates=rates)

        with atomic_output(output_path) as tmp_path:
            wb.save(tmp_path)
    finally:
        wb.close()
    return result
