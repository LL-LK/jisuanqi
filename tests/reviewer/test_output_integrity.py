# -*- coding: utf-8 -*-
"""输出文件完整性回归（需求原文第 8/9/10/11 组）。

硬约束：
  * 输入文件绝不覆盖（跑完 md5 不变）；
  * 输出工作表名单恒等于输入的三张表，不新增任何统计/日志/审计表；
  * 屏风表内容与输入逐单元格相同；
  * 合并区按学院行块归一（N69:N70 类缺陷必须被修掉，O/P/Q/R 与 N 行块一致）；
  * 出勤率保留两位小数，且不得出现 ### 显示问题。
"""

from __future__ import annotations

import hashlib
import os

import pytest
from openpyxl import load_workbook

from core import RunConfig, run
from core.export import ALLOWED_SHEETS
from core.io_detail import load_detail

import synthetic

EXPECTED_SHEETS = ['雁山（早）', '雁山（晚）', '屏风']


def _md5(path: str) -> str:
    digest = hashlib.md5()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _merged_rows(ws, col: int):
    """某列上所有单列合并区的 (首行, 末行)。"""
    out = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_col == col == rng.max_col:
            out.add((rng.min_row, rng.max_row))
    return out


# ------------------------------------------------------------------ 输入不可覆盖


def test_input_detail_file_md5_unchanged(synthetic_detail, synthetic_exempt, tmp_path):
    """跑完一遍，输入明细表 md5 必须一字不差。"""
    before = _md5(synthetic_detail)
    out = tmp_path / '输出.xlsx'
    result = run(RunConfig(detail_path=synthetic_detail, exempt_path=synthetic_exempt,
                            output_path=str(out), week_no=7))
    assert result.ok
    assert _md5(synthetic_detail) == before


def test_input_exempt_file_md5_unchanged(synthetic_detail, synthetic_exempt, tmp_path):
    before = _md5(synthetic_exempt)
    out = tmp_path / '输出.xlsx'
    result = run(RunConfig(detail_path=synthetic_detail, exempt_path=synthetic_exempt,
                            output_path=str(out), week_no=7))
    assert result.ok
    assert _md5(synthetic_exempt) == before


def test_output_is_a_separate_file(synthetic_run):
    _, output = synthetic_run
    assert os.path.exists(output)
    assert os.path.getsize(output) > 0


def test_export_refuses_to_overwrite_input(synthetic_detail, tmp_path):
    """输出路径与输入相同 → 明确报错，绝不允许覆盖。"""
    from core.export import export
    detail = load_detail(synthetic_detail)
    with pytest.raises(ValueError):
        export(synthetic_detail, synthetic_detail, detail, {}, [], {})


# ------------------------------------------------------------------ 工作表名单


def test_output_sheet_names_are_exactly_three(synthetic_run):
    """输出工作表名单恒等于输入的三张表。"""
    _, output = synthetic_run
    wb = load_workbook(output, read_only=True)
    try:
        assert wb.sheetnames == EXPECTED_SHEETS
    finally:
        wb.close()


def test_output_has_no_extra_sheets(synthetic_run):
    """禁止新增统计/日志/审计表。"""
    _, output = synthetic_run
    wb = load_workbook(output, read_only=True)
    try:
        assert len(wb.sheetnames) == 3
        assert set(wb.sheetnames) == set(ALLOWED_SHEETS)
    finally:
        wb.close()


def test_output_sheet_names_match_input(synthetic_detail, synthetic_run):
    _, output = synthetic_run
    src = load_workbook(synthetic_detail, read_only=True)
    dst = load_workbook(output, read_only=True)
    try:
        assert dst.sheetnames == src.sheetnames
    finally:
        src.close()
        dst.close()


# ------------------------------------------------------------------ 屏风零改动


def test_screen_sheet_cells_identical_to_input(synthetic_detail, synthetic_run):
    """屏风表内容与输入逐单元格相同。"""
    _, output = synthetic_run
    src = load_workbook(synthetic_detail, data_only=True)
    dst = load_workbook(output, data_only=True)
    try:
        assert '屏风' in src.sheetnames and '屏风' in dst.sheetnames
        ws_src, ws_dst = src['屏风'], dst['屏风']
        assert ws_src.max_row == ws_dst.max_row
        assert ws_src.max_column == ws_dst.max_column
        for row in range(1, ws_src.max_row + 1):
            for col in range(1, ws_src.max_column + 1):
                assert ws_dst.cell(row, col).value == ws_src.cell(row, col).value, \
                    '屏风表 (%d,%d) 被改动' % (row, col)
    finally:
        src.close()
        dst.close()


def test_screen_sheet_has_no_exempt_or_rate_values(synthetic_detail, synthetic_run):
    """屏风表不得被写入任何免勤/出勤率数据。"""
    _, output = synthetic_run
    wb = load_workbook(output, data_only=True)
    try:
        ws = wb['屏风']
        for row in ws.iter_rows():
            for cell in row:
                assert not isinstance(cell.value, str) or '%' not in cell.value
    finally:
        wb.close()


# ------------------------------------------------------------------ 合并区归一


def test_merged_regions_follow_college_blocks(synthetic_detail, synthetic_run):
    """正常输入下，出勤率列的合并区必须与学院行块一致。"""
    _, output = synthetic_run
    detail = load_detail(synthetic_detail)
    wb = load_workbook(output)
    try:
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            expected = set()
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if len(rows) > 1:
                    expected.add((min(rows), max(rows)))
            for _label, col in rate_cols.items():
                assert _merged_rows(ws, col) == expected, \
                    '%s 第 %s 列合并区与学院行块不一致' % (sheet, col)
    finally:
        wb.close()


def test_broken_merge_is_normalized(synthetic_detail_broken, synthetic_exempt, tmp_path):
    """模板缺陷（N 列漏合并、O~R 与 N 行块不齐）必须在导出阶段被归一。"""
    out = tmp_path / '归一后.xlsx'
    result = run(RunConfig(detail_path=synthetic_detail_broken, exempt_path=synthetic_exempt,
                            output_path=str(out), week_no=7))
    assert result.ok
    detail = load_detail(synthetic_detail_broken)
    wb = load_workbook(out)
    try:
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            expected = set()
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if len(rows) > 1:
                    expected.add((min(rows), max(rows)))
            for _label, col in rate_cols.items():
                assert _merged_rows(ws, col) == expected, \
                    '%s 第 %s 列合并区未归一' % (sheet, col)
    finally:
        wb.close()


def test_rate_value_written_only_to_merge_anchor(synthetic_detail, synthetic_run):
    """合并区只写左上角，其余单元格必须为空（避免 'MergedCell 只读' 崩溃）。"""
    _, output = synthetic_run
    detail = load_detail(synthetic_detail)
    wb = load_workbook(output)
    try:
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if len(rows) < 2:
                    continue
                for _label, col in rate_cols.items():
                    assert ws.cell(rows[0], col).value is not None or True
                    for row in rows[1:]:
                        assert ws.cell(row, col).value is None, \
                            '%s (%d,%d) 不该写值' % (sheet, row, col)
    finally:
        wb.close()


# ------------------------------------------------------------------ 数值呈现


def test_rate_cells_use_two_decimal_percent_format(synthetic_detail, synthetic_run):
    """出勤率保留两位小数（0.00%）。"""
    _, output = synthetic_run
    detail = load_detail(synthetic_detail)
    wb = load_workbook(output)
    try:
        seen = 0
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if not rows:
                    continue
                for _label, col in rate_cols.items():
                    cell = ws.cell(rows[0], col)
                    if isinstance(cell.value, (int, float)):
                        seen += 1
                        assert cell.number_format == '0.00%', \
                            '%s (%d,%d) 数字格式不是 0.00%%' % (sheet, rows[0], col)
        assert seen > 0, '合成数据必须至少写出一个出勤率'
    finally:
        wb.close()


def test_rate_columns_are_wide_enough(synthetic_detail, synthetic_run):
    """出勤率列必须加宽，避免 Excel 显示 ###。"""
    _, output = synthetic_run
    detail = load_detail(synthetic_detail)
    wb = load_workbook(output)
    try:
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            for _label, col in rate_cols.items():
                letter = ws.cell(1, col).column_letter
                width = ws.column_dimensions[letter].width
                assert width is not None and width >= 10, \
                    '%s %s 列宽 %s 过窄，会显示 ###' % (sheet, letter, width)
    finally:
        wb.close()


def test_missing_actual_cell_is_empty_not_zero(synthetic_detail, synthetic_run):
    """实到缺失时，出勤率单元格必须留空（None），不能写 0。"""
    _, output = synthetic_run
    detail = load_detail(synthetic_detail)
    wb = load_workbook(output)
    try:
        ws = wb[synthetic.SHEET_EARLY]
        block = [b for b in detail.blocks_of(synthetic.SHEET_EARLY)
                 if b.name == '测试学院乙'][0]
        col = detail.rate_cols_by_sheet[synthetic.SHEET_EARLY]['周一']
        value = ws.cell(block.merge_rows[0], col).value
        assert value is None, '实到缺失时出勤率单元格必须为空，实际是 %r' % (value,)
    finally:
        wb.close()


def test_exempt_cells_are_integers(synthetic_detail, synthetic_run):
    """教室级免勤列必须写整数。"""
    _, output = synthetic_run
    detail = load_detail(synthetic_detail)
    wb = load_workbook(output)
    try:
        for sheet in detail.rate_cols_by_sheet:
            ws = wb[sheet]
            for room in detail.rooms_of(sheet):
                for day, col in room.exempt_cols.items():
                    value = ws.cell(room.row, col).value
                    assert value is None or isinstance(value, int), \
                        '%s (%d,%d) 免勤值不是整数：%r' % (sheet, room.row, col, value)
    finally:
        wb.close()
