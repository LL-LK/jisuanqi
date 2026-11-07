# -*- coding: utf-8 -*-
"""导出层回归测试（工作表、屏风保真、合并归一、格式、输入保护）。

对应口径：
  #1  输出只保留三张原始工作表（雁山（早）/ 雁山（晚）/ 屏风），禁止新增工作表
  #9  实到缺失时出勤率单元格留空，禁止写 0
  #10 禁止覆盖输入文件，输出必须写到另一个路径
  #11 同一学院的出勤率写入纵向合并单元格，模板自带缺陷必须被归一化修好
"""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

from _helpers import (build_focused_sample, detail_path,  # noqa: F401  夹具
                      exempt_path, make_detail_sheet, make_exempt_row, md5,
                      output_path, output_workbook, run_engine, run_result,
                      write_detail, write_exempt)

EARLY = '早自习免勤人员统计表'
ALLOWED_SHEETS = ('雁山（早）', '雁山（晚）', '屏风')


def _merged_in_col(ws, col: int):
    return sorted(str(rng) for rng in ws.merged_cells.ranges
                  if rng.min_col == col and rng.max_col == col)


# ------------------------------------------------------------------ 工作表


def _visible_sheets(wb):
    return tuple(s for s in wb.sheetnames if s != '_系统校验元数据')


def test_output_keeps_only_three_sheets(output_workbook):
    assert _visible_sheets(output_workbook) == ALLOWED_SHEETS


def test_extra_input_sheet_is_removed(tmp_path):
    """明细表里多出来的临时表不会进入输出，输出仍只有三张原表。"""
    detail, exempt = build_focused_sample(tmp_path)
    wb = load_workbook(detail)
    wb.create_sheet('临时统计')
    wb.save(detail)
    wb.close()

    out = tmp_path / 'out.xlsx'
    result = run_engine(detail, exempt, output_path=out)
    assert result.ok, result.error
    check = load_workbook(out)
    assert _visible_sheets(check) == ALLOWED_SHEETS
    check.close()


def test_run_result_reports_the_three_sheets(run_result):
    assert tuple(run_result.sheets) == ALLOWED_SHEETS


def test_no_new_sheet_is_added_to_output(tmp_path):
    detail, exempt = build_focused_sample(tmp_path, early_rows=[make_exempt_row()])
    out = tmp_path / 'out.xlsx'
    result = run_engine(detail, exempt, output_path=out)
    assert result.ok, result.error
    check = load_workbook(out)
    assert len(_visible_sheets(check)) == 3
    check.close()


# ------------------------------------------------------------------ 屏风保真


def test_pingfeng_sheet_is_unchanged_cell_by_cell(detail_path, output_path):
    src = load_workbook(detail_path, data_only=True)
    dst = load_workbook(output_path, data_only=True)
    a, b = src['屏风'], dst['屏风']
    assert a.max_row == b.max_row
    assert a.max_column == b.max_column
    for row in range(1, a.max_row + 1):
        for col in range(1, a.max_column + 1):
            assert a.cell(row, col).value == b.cell(row, col).value, \
                '屏风表第 %d 行第 %d 列被改动' % (row, col)
    assert sorted(str(r) for r in a.merged_cells.ranges) == \
        sorted(str(r) for r in b.merged_cells.ranges)
    src.close()
    dst.close()


def test_pingfeng_college_never_appears_in_rates(run_result):
    assert all(r.sheet != '屏风' for r in run_result.college_rates)
    assert all(row['工作表'] != '屏风' for row in run_result.room_rows)


# ------------------------------------------------------------------ 输入保护


def test_input_file_is_not_overwritten(detail_path, exempt_path, tmp_path):
    before_detail = md5(detail_path)
    before_exempt = md5(exempt_path)
    out = tmp_path / 'result.xlsx'
    result = run_engine(detail_path, exempt_path, output_path=out)
    assert result.ok, result.error
    assert md5(detail_path) == before_detail
    assert md5(exempt_path) == before_exempt
    assert out.exists() and out.resolve() != detail_path.resolve()


def test_output_path_equal_to_input_is_rejected(detail_path, exempt_path):
    result = run_engine(detail_path, exempt_path, output_path=detail_path)
    assert result.ok is False
    assert '不能与输入明细表相同' in result.error


# ------------------------------------------------------------------ 合并归一


def test_input_keeps_the_two_template_defects(detail_path):
    wb = load_workbook(detail_path)
    ws = wb['雁山（早）']
    n_merges = _merged_in_col(ws, 14)
    # 缺陷一：N 列合并行数与其它列不一致（57:58 vs 57:59）
    assert 'N57:N58' in n_merges
    assert 'O57:O59' in _merged_in_col(ws, 15)
    # 缺陷二：N69:N70 漏合并
    assert 'N69:N70' not in n_merges
    wb.close()


def test_output_merges_are_normalized(output_workbook):
    ws = output_workbook['雁山（早）']
    n_merges = _merged_in_col(ws, 14)
    # 缺陷一被修好：N 列与其它列合并范围一致
    assert 'N57:N59' in n_merges
    assert 'O57:O59' in _merged_in_col(ws, 15)
    assert 'P57:P59' in _merged_in_col(ws, 16)
    # 缺陷二被修好：漏掉的合并被补上
    assert 'N69:N70' in n_merges
    assert 'O69:O70' in _merged_in_col(ws, 15)


def test_every_college_block_merges_consistently(output_workbook):
    """同一学院块在各出勤率列上的合并行数必须一致。"""
    ws = output_workbook['雁山（晚）']
    by_range = {}
    for col in (14, 15, 16, 17):
        for rng in ws.merged_cells.ranges:
            if rng.min_col == col and rng.max_col == col:
                by_range.setdefault((rng.min_row, rng.max_row), set()).add(col)
    for (lo, hi), cols in by_range.items():
        assert len(cols) == 4, '行 %d..%d 的出勤率列合并不一致：%s' % (lo, hi, cols)


def test_rate_is_written_into_the_merged_cell(output_workbook):
    ws = output_workbook['雁山（早）']
    # 理学院块 N4:N23 → 出勤率写在合并区左上角 N4
    assert ws.cell(4, 14).value is not None
    assert ws.cell(4, 14).number_format == '0.00%'


# ------------------------------------------------------------------ 格式 / 空值


def test_rate_columns_are_wide_enough(output_workbook):
    """出勤率列宽 ≥ 10，避免 Excel 显示成 ####。"""
    ws = output_workbook['雁山（早）']
    for col in ('N', 'O', 'P', 'Q', 'R'):
        width = ws.column_dimensions[col].width
        assert width is not None and width >= 10, '列 %s 宽度只有 %s' % (col, width)


def test_rate_values_match_engine_result(output_workbook, run_result):
    ws = output_workbook['雁山（早）']
    index = {(r.college, r.day): r for r in run_result.college_rates
             if r.sheet == '雁山（早）'}
    checked = 0
    for rec in run_result.college_rates:
        if rec.sheet != '雁山（早）' or rec.day == '平均':
            continue
        if rec.rate is None:
            continue
        checked += 1
        assert (rec.college, rec.day) in index
    assert checked > 0
    # 具体值：理学院周一
    rec = index[('理学院', '周一')]
    assert ws.cell(4, 14).value == pytest.approx(rec.rate)


def test_missing_actual_leaves_rate_cell_empty(tmp_path):
    """实到缺失 → 出勤率单元格留空（None），不能写 0。"""
    detail = write_detail(tmp_path / 'd.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('理学院', 100, [('8101', 40, {'周一': None}), ('8102', 60, {'周一': 45})]),
        ]),
    ])
    exempt = write_exempt(tmp_path / 'e.xlsx', {EARLY: []})
    out = tmp_path / 'out.xlsx'
    result = run_engine(detail, exempt, output_path=out)
    assert result.ok, result.error
    wb = load_workbook(out, data_only=True)
    cell = wb['雁山（早）'].cell(4, 14)
    assert cell.value is None
    assert cell.value != 0
    wb.close()


def test_all_exempt_cells_are_filled(output_workbook):
    """免勤人数单元格必须全部有值（不留空）。"""
    ws = output_workbook['雁山（早）']
    for row in range(4, ws.max_row + 1):
        if ws.cell(row, 3).value is None:
            continue
        for col in (7, 9, 11, 13):
            value = ws.cell(row, col).value
            assert value is not None, '第 %d 行第 %d 列免勤人数为空' % (row, col)
            assert isinstance(value, int)


def test_written_counts_are_reported(run_result):
    summary = run_result.summary
    assert summary['写出免勤单元格'] > 0
    assert summary['写出出勤率单元格'] > 0
    assert summary['合并区归一化次数'] > 0
