# -*- coding: utf-8 -*-
"""计算层回归测试（分摊 + 学院出勤率）。

对应口径：
  #2  屏风校区整体排除（不进学院汇总）
  #3  周一 / 周三 / 周四 / 周五 四天独立核算，禁止跨日串算
  #4  教室出勤率 = 实到 ÷ (应到 − 免勤)
  #5  学院出勤率 = 学院总实到 ÷ (学院总应到 − 当次自习学院总免勤)，
      学院总应到以明细表学院名括号内人数为准
  #8  教室级免勤按应到人数比例分摊，分摊之和恒等于学院总免勤
  #9  实到人数缺失一律留空，禁止用 0 代替；分母 ≤ 0 不产出出勤率
"""

from __future__ import annotations

import pytest

from _helpers import (build_focused_sample, make_detail_sheet,  # noqa: F401  夹具
                      make_exempt_row, write_detail, write_exempt)
from core.calc import Calculator, allocate
from core.io_detail import load_detail
from core.io_exempt import aggregate_exemptions, load_exemptions

EARLY = '早自习免勤人员统计表'
LATE = '晚自习免勤人员统计表'


def _compute(tmp_path, blocks, rows, week_no=7, sheets=None):
    """跑一遍「读取 → 聚合 → 分摊 → 算率」，返回 (calc, detail, agg)。"""
    detail = write_detail(tmp_path / 'detail_calc.xlsx',
                           sheets or [make_detail_sheet('雁山（早）', blocks)])
    exempt = write_exempt(tmp_path / 'exempt_calc.xlsx', {EARLY: rows})
    model = load_detail(str(detail))
    records = load_exemptions(str(exempt))
    agg = aggregate_exemptions(records, week_no, campus='雁山')
    calc = Calculator(model, agg, alias={})
    return calc, model, agg


def _rate(calc, sheet='雁山（早）', college='理学院', day='周一'):
    for rec in calc.college_rates():
        if rec.sheet == sheet and rec.college == college and rec.day == day:
            return rec
    raise AssertionError('找不到 %s / %s / %s 的出勤率记录' % (sheet, college, day))


# ------------------------------------------------------------------ 分摊


def test_allocate_sum_equals_total():
    assert sum(allocate(7, [2, 0, 5])) == 7
    assert sum(allocate(5, [3, 7])) == 5
    assert sum(allocate(10, [1, 1, 1])) == 10
    assert sum(allocate(100, [33, 33, 34])) == 100


def test_allocate_largest_remainder():
    """最大余数法：和恒等，各项之间最多差 1。"""
    assert sorted(allocate(5, [3, 7])) == [2, 3]
    equal = allocate(10, [1, 1, 1])
    assert sum(equal) == 10
    assert max(equal) - min(equal) <= 1
    assert allocate(4, [0, 0, 4]) == [0, 0, 4]


def test_allocate_edge_cases():
    assert allocate(0, [1, 2]) == [0, 0]
    assert allocate(-3, [1, 2]) == [0, 0]
    assert allocate(5, []) == []
    # 没有有效权重时全部落到第一项，保证和相等
    assert allocate(5, [0, 0]) == [5, 0]


def test_allocate_is_proportional():
    share = allocate(10, [40, 60])
    assert share == [4, 6]
    assert sum(share) == 10


def test_room_exempt_split_is_proportional(tmp_path):
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30}), ('8102', 60, {'周一': 45})]),
    ], [make_exempt_row(sid='99%011d' % i) for i in range(1, 11)])
    room_exempt = calc.room_exemptions()
    assert room_exempt[('雁山（早）', 4, '周一')] == 4
    assert room_exempt[('雁山（早）', 5, '周一')] == 6
    assert sum(room_exempt.values()) == 10


def test_room_exempt_sum_equals_college_exempt(tmp_path):
    calc, _model, agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30}), ('8102', 60, {'周一': 45})]),
    ], [make_exempt_row(sid='99%011d' % i) for i in range(1, 8)])
    room_exempt = calc.room_exemptions()
    total = sum(v for (sheet, row, day), v in room_exempt.items() if day == '周一')
    assert total == agg.exempt_count('理学院', '周一', '早自习')
    assert total == 7


# ------------------------------------------------------------------ 学院出勤率


def test_college_rate_formula(tmp_path):
    """学院出勤率 = 学院总实到 ÷ (学院总应到 − 当次自习学院总免勤)。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30}), ('8102', 60, {'周一': 45})]),
    ], [make_exempt_row(sid='99%011d' % i) for i in range(1, 11)])
    rec = _rate(calc)
    assert rec.expected == 100
    assert rec.actual == 75
    assert rec.exempt == 10
    assert rec.rate == pytest.approx(75 / 90)
    assert rec.status == '完整'


def test_college_expected_uses_bracket_label(tmp_path):
    """学院总应到以明细表学院名括号内人数为准，教室应到合计只作交叉核对。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 120, [('8101', 40, {'周一': 30}), ('8102', 60, {'周一': 45})]),
    ], [])
    rec = _rate(calc)
    assert rec.expected == 120           # 括号内人数
    assert rec.room_expected_sum == 100  # 教室应到合计仅作交叉核对
    assert rec.rate == pytest.approx(75 / 120)
    assert any('学院标注总人数' in a for a in rec.anomalies)


def test_single_room_college_rate_equals_room_formula(tmp_path):
    """只有一间教室时，学院出勤率就是教室出勤率 = 实到 ÷ (应到 − 免勤)。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 40, [('8101', 40, {'周一': 30})]),
    ], [make_exempt_row(sid='99%011d' % i) for i in range(1, 6)])
    rec = _rate(calc)
    assert rec.rate == pytest.approx(30 / (40 - 5))


def test_missing_actual_gives_no_rate(tmp_path):
    """实到缺失 → 不产出出勤率（不是 0）。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': None}), ('8102', 60, {'周一': 45})]),
    ], [])
    rec = _rate(calc)
    assert rec.rate is None
    assert rec.actual is None
    assert rec.status == '实到数据未完整填写'


def test_text_actual_counts_as_missing(tmp_path):
    """实到写成文字「免勤」按缺失处理。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': '免勤'}), ('8102', 60, {'周一': 45})]),
    ], [])
    rec = _rate(calc)
    assert rec.rate is None
    assert rec.status == '实到数据未完整填写'


def test_zero_actual_is_not_missing(tmp_path):
    """实到为 0 是有效数据（分母为正时照常算率）。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 0}), ('8102', 60, {'周一': 45})]),
    ], [])
    rec = _rate(calc)
    assert rec.actual == 45
    assert rec.rate == pytest.approx(45 / 100)


def test_nonpositive_denominator_gives_no_rate(tmp_path):
    """分母 ≤ 0 不产出出勤率。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 10, [('8101', 10, {'周一': 8})]),
    ], [make_exempt_row(sid='99%011d' % i) for i in range(1, 11)])
    rec = _rate(calc)
    assert rec.rate is None
    assert rec.status == '分母非正数，无法计算'
    assert rec.exempt == 10


def test_days_are_independent(tmp_path):
    """周一的免勤不得串到周三。"""
    calc, _model, agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30, '周三': 28}),
                         ('8102', 60, {'周一': 45, '周三': 44})]),
    ], [make_exempt_row(sid='99%011d' % i, day_text='周一')
        for i in range(1, 11)])
    monday = _rate(calc, day='周一')
    wednesday = _rate(calc, day='周三')
    assert monday.exempt == 10
    assert wednesday.exempt == 0
    assert wednesday.rate == pytest.approx(72 / 100)
    assert monday.rate != wednesday.rate


def test_average_blank_when_any_day_missing(tmp_path):
    """口径（对话口径）：平均出勤率只在**所有应统计天**都算出时才填。

    缺任一天即留空，不用部分天的平均冒充整周平均。
    """
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30, '周三': 28, '周四': 29, '周五': None}),
                         ('8102', 60, {'周一': 45, '周三': 44, '周四': 46, '周五': 47})]),
    ], [])
    rates = calc.college_rates()
    averages = calc.average_rates(rates)
    assert ('雁山（早）', '理学院') not in averages


def test_average_present_when_all_days_present(tmp_path):
    """所有天都算出时，平均值必须是各天出勤率的算术平均。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30, '周三': 28, '周四': 29, '周五': 31}),
                         ('8102', 60, {'周一': 45, '周三': 44, '周四': 46, '周五': 47})]),
    ], [])
    rates = calc.college_rates()
    averages = calc.average_rates(rates)
    daily = [r.rate for r in rates if r.rate is not None]
    assert len(daily) == 4
    assert averages[('雁山（早）', '理学院')] == pytest.approx(sum(daily) / len(daily))


def test_average_blank_when_nothing_computed(tmp_path):
    """全部天都算不出时，平均值必须留空（不写 0）。"""
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 100, {'周一': None, '周三': None, '周四': None, '周五': None})]),
    ], [])
    rates = calc.college_rates()
    averages = calc.average_rates(rates)
    assert ('雁山（早）', '理学院') not in averages


def test_average_present_when_all_days_computed(tmp_path):
    calc, _model, _agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30, '周三': 28, '周四': 29, '周五': 30}),
                         ('8102', 60, {'周一': 45, '周三': 44, '周四': 46, '周五': 47})]),
    ], [])
    rates = calc.college_rates()
    averages = calc.average_rates(rates)
    assert ('雁山（早）', '理学院') in averages
    assert averages[('雁山（早）', '理学院')] == pytest.approx(
        sum(r.rate for r in rates if r.day != '平均') / 4)


def test_pingfeng_is_excluded_from_rates(tmp_path):
    """屏风表不参与计算：屏风学院不出现在学院汇总里。"""
    detail = write_detail(tmp_path / 'd.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('理学院', 100, [('8101', 40, {'周一': 30})]),
        ]),
        make_detail_sheet('屏风', [
            ('化工学院', 50, [('9201', 50, {'周一': 45})]),
        ]),
    ])
    exempt = write_exempt(tmp_path / 'e.xlsx', {EARLY: []})
    model = load_detail(str(detail))
    agg = aggregate_exemptions(load_exemptions(str(exempt)), 7)
    calc = Calculator(model, agg, alias={})
    colleges = {r.college for r in calc.college_rates()}
    assert colleges == {'理学院'}
    assert all(r.sheet != '屏风' for r in calc.college_rates())
    assert all(r.sheet != '屏风' for r in model.rooms)


def test_college_exempt_uses_unique_sid_count(tmp_path):
    """学院总免勤按唯一学号去重后的计数，不重复累加分摊值。"""
    rows = [make_exempt_row(sid='9900000000001'),
            make_exempt_row(sid='9900000000001', name='学生甲 1'),
            make_exempt_row(sid='9900000000002', name='学生乙02')]
    calc, _model, agg = _compute(tmp_path, [
        ('理学院', 100, [('8101', 40, {'周一': 30}), ('8102', 60, {'周一': 45})]),
    ], rows)
    rec = _rate(calc)
    assert agg.exempt_count('理学院', '周一', '早自习') == 2
    assert rec.exempt == 2
    assert rec.rate == pytest.approx(75 / 98)
