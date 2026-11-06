# -*- coding: utf-8 -*-
"""界面层验收门（SL1–SL4，标记 `uigate`）。

v1.0.5 入口条件 3「界面行为中性对拍 0 差异」由两部分组成：
  * **数字侧**（在默认闸门内）：`test_acceptance_baseline.py::test_output_content_digest_unchanged`
    —— 产出内容摘要逐项 sha256 锁定，界面改动一旦碰到任何数字/合并区立刻转红；
  * **显示侧**（本模块，`uigate`）：SL1–SL4 四条显示层裁定的可执行验收。

按需求方/Chief 裁定，docgate/uigate 属**发版前条件**、不是构建条件：
  * 默认闸门（`pytest tests -q`）不含本模块；
  * 发版前单独执行 `python -X utf8 -m pytest tests -m uigate -q`，须全绿。
"""

from __future__ import annotations

import os
import sys

import pytest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

pytest.importorskip('PySide6', reason='未安装 PySide6，跳过界面层用例')

pytestmark = pytest.mark.uigate

from PySide6.QtCore import Qt                              # noqa: E402
from PySide6.QtGui import QBrush, QColor, QFont           # noqa: E402
from PySide6.QtWidgets import QApplication                 # noqa: E402

import app as gui                                         # noqa: E402

ESTIMATE_RGB = 0xFFF3CD            # _bg_rgb 返回去掉 alpha 的 RGB
WARN_RGB = 0xFFF6DC                # 异常行警示底色


@pytest.fixture(scope='module')
def qapp():
    return QApplication.instance() or QApplication(sys.argv[:1])


@pytest.fixture(scope='module')
def rendered(qapp, real_detail_path, real_exempt_path, tmp_path_factory):
    """跑一次真实流水线并把结果渲染进主窗口，返回 (window, result)。"""
    from core import RunConfig, run
    out = tmp_path_factory.mktemp('uigate') / '界面验收输出.xlsx'
    result = run(RunConfig(detail_path=real_detail_path, exempt_path=real_exempt_path,
                            output_path=str(out), week_no=7))
    assert result.ok, result.error
    win = gui.MainWindow()
    win._render(result)
    yield win, result
    win.close()


def _headers(view):
    model = view.model()
    return [str(model.headerData(i, Qt.Horizontal, Qt.DisplayRole) or '')
            for i in range(model.columnCount())]


def _row_color(model, row):
    return model.row_colors.get(row)


def _role(model, row, col, role):
    return model.data(model.index(row, col), role)


def _bg_rgb(model, row, col):
    """取单元格背景色（QColor 或 QBrush 都要兼容）。"""
    value = _role(model, row, col, Qt.BackgroundRole)
    if isinstance(value, QBrush):
        value = value.color()
    if isinstance(value, QColor):
        return value.rgb() & 0xFFFFFF
    return None


def _is_italic(model, row, col):
    value = _role(model, row, col, Qt.FontRole)
    if isinstance(value, QFont):
        return bool(value.italic())
    if isinstance(value, str):
        return 'italic' in value.lower()
    return False


# ------------------------------------------------------------------ SL1


def test_room_detail_table_shows_room_level_exempt_actual_rate(rendered):
    """SL1：教室明细页必须增列**教室级免勤 / 实到 / 出勤率**（数据侧 room_rows 已有）。"""
    win, _result = rendered
    headers = _headers(win.room_table)
    assert len(headers) > 6, '教室明细页仍是 6 列：%s' % headers
    assert any('免勤' in h for h in headers), '教室明细页没有教室级免勤列：%s' % headers
    assert any('实到' in h for h in headers), '教室明细页没有实到列：%s' % headers


def test_room_detail_table_shows_actual_room_exempt_values(rendered):
    """SL1 续：教室级免勤要真的显示数值（不是只加表头）。"""
    win, result = rendered
    model = win.room_table.model()
    headers = _headers(win.room_table)
    col = next((i for i, h in enumerate(headers) if '周一免勤' in h), None)
    assert col is not None, '缺少「周一免勤」列：%s' % headers
    values = [model.data(model.index(r, col)) for r in range(model.rowCount())]
    numeric = [v for v in values if isinstance(v, (int, float))]
    assert numeric, '教室级免勤列没有任何数值'
    expected = [r.get('周一免勤') for r in result.room_rows
                if isinstance(r.get('周一免勤'), (int, float))]
    assert sum(numeric) == sum(expected)


# ------------------------------------------------------------------ SL2


def test_estimate_mark_is_same_color_and_italic_in_ui(rendered):
    """SL2：界面里的估算值必须与导出工作簿**同色同形态**（FFFFF3CD + 斜体）。"""
    win, _result = rendered
    model = win.room_table.model()
    headers = _headers(win.room_table)
    col = next((i for i, h in enumerate(headers) if '周一免勤' in h), None)
    assert col is not None, '缺少「周一免勤」列：%s' % headers
    found_rgb = found_italic = False
    for row in range(model.rowCount()):
        value = model.data(model.index(row, col))
        if not isinstance(value, (int, float)) or not value:
            continue
        rgb = _bg_rgb(model, row, col)
        if rgb == ESTIMATE_RGB:
            found_rgb = True
        if _is_italic(model, row, col):
            found_italic = True
    assert found_rgb, '界面未用 R17 固定色 FFFFF3CD 标记估算值'
    assert found_italic, '界面未用斜体标记估算值'


# ------------------------------------------------------------------ SL3


def test_missing_actual_displays_dash_not_blank_not_none(rendered):
    """SL3：缺失实到在界面显示「—」（不是空白、也不是字面量 'None'）；导出仍为真空白。"""
    win, result = rendered
    model = win.rate_table.model()
    headers = _headers(win.rate_table)
    col = next((i for i, h in enumerate(headers) if '学院总实到' in h), None)
    assert col is not None, '缺少「学院总实到」列：%s' % headers
    missing_rows = [r for r in result.college_rates if r.actual is None]
    assert missing_rows, '真实数据里应存在实到缺失的学院×天组合'
    shown = set()
    for row in range(model.rowCount()):
        value = model.data(model.index(row, col))
        if value is None:
            continue
        text = str(value)
        assert text != 'None', '界面把 None 显示成字面量 "None"'
        shown.add(text)
    assert '—' in shown, '缺失实到未用「—」占位，实际取值集合：%s' % shown


# ------------------------------------------------------------------ SL4


def test_incomplete_status_row_is_visually_distinguished(rendered):
    """SL4：状态为「免勤不完整（待补人数）」的行必须有视觉区分（整行底色）。"""
    win, result = rendered
    model = win.rate_table.model()
    headers = _headers(win.rate_table)
    status_col = next((i for i, h in enumerate(headers) if h == '状态'), None)
    assert status_col is not None, headers
    hit = 0
    for row in range(model.rowCount()):
        if str(model.data(model.index(row, status_col)) or '') != '免勤不完整（待补人数）':
            continue
        hit += 1
        color = _row_color(model, row)
        assert color is not None, '「免勤不完整（待补人数）」行没有行级底色'
        rgb = color.rgb() & 0xFFFFFF if isinstance(color, QColor) else None
        assert rgb == WARN_RGB, '异常行底色 %s，应为 %s' % (rgb, WARN_RGB)
    assert hit, '真实数据里应存在「免勤不完整（待补人数）」的学院×天组合'


def test_late_friday_trace_rows_are_visually_distinct(rendered):
    """SL4 续：晚表周五留痕行必须与「出数」行**视觉可分**（否则 R22 的「有值 ≠ 出数」被误读）。"""
    win, result = rendered
    model = win.room_table.model()
    headers = _headers(win.room_table)
    fri_col = next((i for i, h in enumerate(headers) if '周五免勤' in h), None)
    assert fri_col is not None, '缺少「周五免勤」列：%s' % headers
    sheet_col = next((i for i, h in enumerate(headers) if h == '工作表'), None)

    trace_rows, normal_rows = [], []
    for row in range(model.rowCount()):
        sheet = str(model.data(model.index(row, sheet_col)) or '')
        value = model.data(model.index(row, fri_col))
        if not isinstance(value, (int, float)) or not value:
            continue
        (trace_rows if '晚' in sheet else normal_rows).append(row)
    assert trace_rows, '晚表周五留痕行应存在'
    # 视觉可分：要么该行有专属底色，要么表格提供了「参与/留痕/来源」类标识列
    marked = any(_row_color(model, row) is not None for row in trace_rows)
    flag_col = any(any(k in h for k in ('参与', '留痕', '来源', '是否计入'))
                   for h in headers)
    assert marked or flag_col, \
        '晚表周五留痕行与出数行在界面上不可区分（既无行底色，也无标识列）'


def test_ui_does_not_change_reported_numbers(rendered):
    """入口条件 3 的界面侧底线：界面渲染不得改动任何上报数字。"""
    win, result = rendered
    model = win.rate_table.model()
    headers = _headers(win.rate_table)
    rate_col = next((i for i, h in enumerate(headers) if h == '出勤率'), None)
    status_col = next((i for i, h in enumerate(headers) if h == '状态'), None)
    exempt_col = next((i for i, h in enumerate(headers) if h == '当次总免勤'), None)
    expected = {}
    for rec in result.college_rates:
        expected[(rec.sheet, rec.college, rec.day)] = rec
    checked = 0
    for row in range(model.rowCount()):
        sheet = str(model.data(model.index(row, 0)) or '')
        college = str(model.data(model.index(row, 1)) or '')
        day = str(model.data(model.index(row, 2)) or '')
        rec = expected.get((sheet, college, day))
        if rec is None:
            continue
        if exempt_col is not None:
            # 表格模型按显示口径返回字符串，比较时统一按字符串比对
            assert str(model.data(model.index(row, exempt_col))) == str(rec.exempt)
        if status_col is not None:
            assert str(model.data(model.index(row, status_col)) or '') == rec.status
        if rate_col is not None and rec.rate is not None:
            shown = str(model.data(model.index(row, rate_col)) or '')
            assert shown == '%.2f%%' % (rec.rate * 100), \
                '界面出勤率与引擎值不一致：%s vs %r' % (shown, rec.rate)
        checked += 1
    assert checked > 0
