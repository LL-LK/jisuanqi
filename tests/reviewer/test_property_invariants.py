# -*- coding: utf-8 -*-
"""性质化不变量（hypothesis）。

六条不变量：
  I1 任意脏输入不抛未捕获异常（fail fast 也必须是**明确**异常，不得是裸崩溃）；
  I2 同一学号在同一 (星期, 自习) 最多计 1 次；
  I3 教室分摊值之和恒等于学院总免勤；
  I4 分母 ≤0 时不产出任何出勤率数值；
  I5 实到缺失时产出空值而非 0；
  I6 输出工作表名单恒等于输入三张表。
"""

from __future__ import annotations

import os

import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st

from core import RunConfig, run
from core.calc import Calculator, allocate
from core.io_exempt import Aggregate, aggregate_exemptions

import synthetic

COLLEGE = '测试学院甲'
CLASS_A = '测试专业甲2026-1班'
DAYS = ('周一', '周三', '周四', '周五')

dirty_text = st.text(max_size=40)
messy_week = st.sampled_from([
    '第7周', '第7-18周', '第6-7，9-13周', '第7—12周', '第七周，第八周',
    '第7、8、9周', '6-10周', '全学期', '', '   ', '第', '周', 'abc', '第周',
])
messy_time = st.sampled_from([
    '周一', '周一、周三、周四、周五', '周一 ，周三 ，周四', '周一到周五',
    '周一周二周三周四周五', '周一、二、四、五', '第一 三 四', '', '今天',
])


# ------------------------------------------------------------------ I1 不崩


@given(text=dirty_text)
def test_dirty_text_never_raises_unexpected_exception(text):
    """任意脏文本都不得触发未捕获异常。"""
    from core.parsing import (extract_group_size, normalize_campus, normalize_class,
                              normalize_room, normalize_text, parse_college_label,
                              parse_int, parse_session, parse_weekdays, parse_weeks)
    assert isinstance(parse_weeks(text), set)
    assert isinstance(parse_weekdays(text), set)
    assert isinstance(normalize_campus(text), str)
    assert isinstance(normalize_class(text), str)
    assert isinstance(normalize_room(text), str)
    assert isinstance(normalize_text(text), str)
    assert parse_int(text) is None or isinstance(parse_int(text), int)
    assert parse_session(text) in (None, '早自习', '晚自习')
    assert extract_group_size(text) is None or isinstance(extract_group_size(text), int)
    name, total = parse_college_label(text)
    assert isinstance(name, str) and (total is None or isinstance(total, int))


@given(week=messy_week, time=messy_time, campus=st.sampled_from(
    ['雁山校区', '雁山校园', '燕山校区', '屏风校区', '屏风', '雁门校区', '未知校区', '']))
def test_dirty_exempt_record_never_raises(week, time, campus):
    """脏免勤记录只能变成「计入」或「记异常」，不得抛异常。"""
    records = [synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE,
                                     class_text=CLASS_A, weeks=week, time=time, campus=campus)]
    agg = aggregate_exemptions(records, 7, campus='雁山')
    assert agg.exempt_count(COLLEGE, '周一', '早自习') >= 0


def test_missing_input_fails_fast_with_explicit_error(tmp_path):
    """fail fast 也必须是明确异常（run() 返回 ok=False + 异常类型名，而不是崩溃）。"""
    result = run(RunConfig(detail_path=str(tmp_path / '不存在.xlsx'), week_no=7))
    assert result.ok is False
    assert 'FileNotFoundError' in result.error
    assert result.output_path is None


# ------------------------------------------------------------------ I2 去重


@st.composite
def _exempt_rows(draw):
    sids = draw(st.lists(st.sampled_from(['SYN20260001', 'SYN20260002', 'SYN20260003']),
                          min_size=1, max_size=3, unique=True))
    days = draw(st.lists(st.sampled_from(list(DAYS)), min_size=1, max_size=4, unique=True))
    sessions = draw(st.lists(st.sampled_from(['早自习', '晚自习']), min_size=1, max_size=2, unique=True))
    weeks = draw(st.sampled_from(['第7周', '第7-18周', '第6-7，9-13周', '第七周，第八周', '6-10周']))
    rows = []
    for _ in range(draw(st.integers(min_value=0, max_value=12))):
        rows.append((draw(st.sampled_from(sids)), draw(st.sampled_from(days)),
                     draw(st.sampled_from(sessions)), weeks))
    return rows


@given(rows=_exempt_rows())
def test_same_sid_counted_at_most_once_per_day_and_session(rows):
    """同一学号在同一 (星期, 自习) 最多计 1 次。"""
    records = [synthetic.make_record(sid=sid, college=COLLEGE, class_text=CLASS_A,
                                     weeks=week, time=day, session=session)
               for sid, day, session, week in rows]
    agg = aggregate_exemptions(records, 7, campus='雁山')
    for day in DAYS:
        for session in ('早自习', '晚自习'):
            expected = len({sid for sid, d, s, _w in rows if d == day and s == session})
            assert agg.exempt_count(COLLEGE, day, session) == expected


# ------------------------------------------------------------------ I3 分摊恒等


@given(total=st.integers(min_value=0, max_value=5000),
       weights=st.lists(st.integers(min_value=0, max_value=200),
                        min_size=1, max_size=8))
def test_allocation_sum_always_equals_total(total, weights):
    """教室分摊值之和恒等于要分配的总人数。"""
    assert sum(allocate(total, weights)) == total


@given(total=st.integers(min_value=0, max_value=200),
       rooms=st.integers(min_value=1, max_value=6))
def test_allocation_sum_equals_college_exempt_without_cap(total, rooms):
    """不触发教室应到封顶时，分摊之和 == 学院总免勤（分摊值不得回加学院总免勤）。

    每间教室的应到被设成足够大，保证任何一间都不会成为封顶瓶颈。
    """
    each = max(total, 1) + 10
    detail = synthetic.make_detail_model([
        (COLLEGE, each * rooms,
         [('T%03d' % i, each, {'周一': 0}) for i in range(rooms)]),
    ])
    agg = Aggregate()
    agg.personal[(COLLEGE, '周一', '早自习')] = {synthetic.fake_sid(i) for i in range(total)}
    calc = Calculator(detail, agg)
    room_exempt = calc.room_exemptions()
    allocated = sum(v for (sheet, row, day), v in room_exempt.items()
                    if day == '周一' and sheet == synthetic.SHEET_EARLY)
    assert allocated == total == agg.exempt_count(COLLEGE, '周一', '早自习')


# ------------------------------------------------------------------ I4 分母非正


@given(expected=st.integers(min_value=1, max_value=200),
       excess=st.integers(min_value=0, max_value=100))
def test_nonpositive_denominator_yields_no_rate(expected, excess):
    """分母 ≤0 时不产出任何出勤率数值。"""
    exempt = expected + excess
    detail = synthetic.make_detail_model([
        (COLLEGE, expected, [('T001', expected, {'周一': 0})]),
    ])
    agg = Aggregate()
    agg.personal[(COLLEGE, '周一', '早自习')] = {synthetic.fake_sid(i) for i in range(exempt)}
    rates = Calculator(detail, agg).college_rates()
    assert rates
    for rec in rates:
        if rec.day == '周一':
            assert rec.rate is None
            assert rec.status != '完整'


# ------------------------------------------------------------------ I5 缺失即空值


@given(actual=st.integers(min_value=0, max_value=99),
       missing=st.booleans())
def test_missing_actual_produces_none_not_zero(actual, missing):
    """实到缺失时产出空值（None）而非 0。"""
    values = {'周一': (None if missing else actual)}
    detail = synthetic.make_detail_model([
        (COLLEGE, 50, [('T001', 50, values)]),
    ])
    agg = Aggregate()
    rates = Calculator(detail, agg).college_rates()
    rec = [r for r in rates if r.day == '周一'][0]
    if missing:
        assert rec.actual is None
        assert rec.rate is None
    else:
        assert rec.actual == actual
        assert rec.rate == pytest.approx(actual / 50)


# ------------------------------------------------------------------ I6 工作表名单


@settings(max_examples=8, deadline=None,
          suppress_health_check=[HealthCheck.too_slow, HealthCheck.function_scoped_fixture])
@given(extra=st.text(min_size=1, max_size=12).filter(
    lambda s: (s not in (synthetic.SHEET_EARLY, synthetic.SHEET_LATE, synthetic.SHEET_PINGFENG))
    and not (set(s) & set('[]:*?/\\'))
    and not any(ord(ch) < 32 or ord(ch) == 127 for ch in s)))
def test_output_sheet_names_always_three(extra, tmp_path):
    """无论输入里多出什么表，输出恒等于三张原始工作表。"""
    from openpyxl import load_workbook
    detail_path = str(tmp_path / '带杂表.xlsx')
    synthetic.build_detail_workbook(detail_path)
    wb = load_workbook(detail_path)
    wb.create_sheet(extra)
    wb.save(detail_path)
    wb.close()

    out = str(tmp_path / '输出.xlsx')
    result = run(RunConfig(detail_path=detail_path, output_path=out, week_no=7))
    assert result.ok
    wb = load_workbook(out, read_only=True)
    try:
        assert wb.sheetnames == [synthetic.SHEET_EARLY, synthetic.SHEET_LATE,
                                 synthetic.SHEET_PINGFENG]
    finally:
        wb.close()
    assert os.path.exists(out)
