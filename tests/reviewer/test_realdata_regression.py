# -*- coding: utf-8 -*-
"""真实工作簿回归（只读；产物写到 pytest 临时目录，绝不碰原始文件）。

真实文件只作**只读输入**，输出写到 tmp_path；
本模块不打印、不导出任何学号与姓名。
"""

from __future__ import annotations

import hashlib
import os

import pytest
from openpyxl import load_workbook

from core import RunConfig, run
from core.export import ALLOWED_SHEETS
from core.io_detail import load_detail

pytestmark = pytest.mark.realdata


def _md5(path: str) -> str:
    digest = hashlib.md5()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b''):
            digest.update(chunk)
    return digest.hexdigest()


def _merged_rows(ws, col: int):
    return {(rng.min_row, rng.max_row) for rng in ws.merged_cells.ranges
            if rng.min_col == col == rng.max_col}


@pytest.fixture(scope='module')
def real_run(real_detail_path, real_exempt_path, tmp_path_factory):
    out_dir = tmp_path_factory.mktemp('real')
    out = out_dir / '真实数据输出.xlsx'
    before = _md5(real_detail_path)
    result = run(RunConfig(detail_path=real_detail_path, exempt_path=real_exempt_path,
                            output_path=str(out), week_no=7))
    assert result.ok, '真实数据流水线必须跑通：%s' % result.error
    assert _md5(real_detail_path) == before, '输入明细表被改动了'
    return result, str(out), real_detail_path


def test_real_output_has_exactly_three_sheets(real_run):
    _result, output, _detail = real_run
    wb = load_workbook(output, read_only=True)
    try:
        assert wb.sheetnames == list(ALLOWED_SHEETS)
        assert len(wb.sheetnames) == 3
    finally:
        wb.close()


def test_real_screen_sheet_is_byte_identical_in_content(real_run):
    """屏风表逐单元格零改动。"""
    _result, output, detail = real_run
    src = load_workbook(detail, data_only=True)
    dst = load_workbook(output, data_only=True)
    try:
        ws_src, ws_dst = src['屏风'], dst['屏风']
        assert (ws_src.max_row, ws_src.max_column) == (ws_dst.max_row, ws_dst.max_column)
        for row in range(1, ws_src.max_row + 1):
            for col in range(1, ws_src.max_column + 1):
                assert ws_dst.cell(row, col).value == ws_src.cell(row, col).value
    finally:
        src.close()
        dst.close()


def test_real_merge_regions_follow_college_blocks(real_run):
    """真实文件里的 N69:N70 漏合并、O/P/Q/R 行块不齐，必须在导出阶段被归一。"""
    _result, output, detail_path = real_run
    detail = load_detail(detail_path)
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
                actual = _merged_rows(ws, col)
                assert actual == expected, \
                    '%s 第 %s 列合并区 %s 与学院行块 %s 不一致' % (sheet, col, actual, expected)
    finally:
        wb.close()


def test_real_college_rates_are_recomputable(real_run):
    """真实数据算出的每条出勤率都必须能被公式复算出来。"""
    result, _output, _detail = real_run
    checked = 0
    for rec in result.college_rates:
        if rec.rate is None:
            continue
        denominator = rec.expected - rec.exempt
        assert denominator > 0
        assert rec.rate == pytest.approx(rec.actual / denominator, abs=1e-12)
        checked += 1
    assert checked > 0


def test_real_pingfeng_not_in_any_college_rate(real_run):
    """屏风不得出现在任何学院出勤率里。"""
    result, _output, _detail = real_run
    for rec in result.college_rates:
        assert '屏风' not in rec.sheet
        assert rec.sheet in ('雁山（早）', '雁山（晚）')


def test_real_missing_actual_yields_blank_not_zero(real_run):
    """实到缺失时输出单元格必须为空，不能写 0。

    注意：实到**确实为 0** 时，出勤率 0.00% 是合法结果，不能误判为「缺失写 0」。
    判定依据是学院出勤率记录本身：actual 为 None → 单元格必须为空。
    """
    result, output, detail_path = real_run
    detail = load_detail(detail_path)
    by_key = {(r.sheet, r.college, r.day): r for r in result.college_rates}
    wb = load_workbook(output)
    try:
        blanks = 0
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if not rows:
                    continue
                for label, col in rate_cols.items():
                    if label == '平均':
                        continue
                    rec = by_key.get((sheet, block.name, label))
                    if rec is None:
                        continue
                    value = ws.cell(rows[0], col).value
                    if rec.actual is None:
                        assert value is None, \
                            '实到缺失却写了 %r（%s / %s / %s）' % (value, sheet, block.name, label)
                        blanks += 1
                    else:
                        assert value == pytest.approx(rec.rate, abs=1e-9), \
                            '写出的出勤率与计算值不一致：%s / %s / %s' % (sheet, block.name, label)
        assert blanks > 0, '真实数据里存在实到缺失的格子'
    finally:
        wb.close()


def test_real_output_file_is_written_next_to_nothing(real_run):
    """输出必须是独立文件，不得覆盖任何输入。"""
    _result, output, detail_path = real_run
    assert os.path.exists(output)
    assert os.path.abspath(output) != os.path.abspath(detail_path)
