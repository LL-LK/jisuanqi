# -*- coding: utf-8 -*-
"""性质化不变量测试（hypothesis）。

覆盖的不变量：

  * 任意脏周次 / 星期 / 校区字符串都不抛未捕获异常；
  * 同一学号在同一 (星期, 自习) 最多计 1 次；
  * 教室分摊值之和恒等于学院总免勤；
  * 分母 ≤ 0 时不产出任何出勤率数值；
  * 实到缺失时产出为空值而非 0；
  * 输出工作表名单恒等于三张原表；
  * 跑批后输入文件 md5 不变。
"""

from __future__ import annotations

import itertools

import pytest
from hypothesis import HealthCheck, given, settings
from hypothesis import strategies as st
from openpyxl import load_workbook

from _helpers import (build_focused_sample, exempt_records,  # noqa: F401  夹具
                      make_detail_sheet, make_exempt_row, md5, run_engine,
                      write_detail, write_exempt)
from core.calc import allocate
from core.io_detail import load_detail
from core.io_exempt import aggregate_exemptions, load_exemptions
from core.parsing import (WEEKDAY_NAMES, normalize_campus, parse_weekdays,
                          parse_weeks)

EARLY = '早自习免勤人员统计表'
ALLOWED_SHEETS = ('雁山（早）', '雁山（晚）', '屏风')

DIRTY_TEXT = st.text(max_size=40)


@pytest.fixture(scope='session')
def work_dir(tmp_path_factory):
    """为每个 hypothesis 例子生成独立的工作目录（避免函数级夹具冲突）。"""
    counter = itertools.count()

    def make(prefix='case'):
        return tmp_path_factory.mktemp('%s_%d' % (prefix, next(counter)))

    return make


# ------------------------------------------------------------------ 鲁棒性


@given(DIRTY_TEXT)
def test_fuzzy_week_text_never_raises(text):
    weeks = parse_weeks(text)
    assert isinstance(weeks, set)
    assert all(isinstance(w, int) and 0 <= w <= 99 for w in weeks)


@given(DIRTY_TEXT)
def test_fuzzy_weekday_text_never_raises(text):
    days = parse_weekdays(text)
    assert isinstance(days, set)
    assert days <= set(WEEKDAY_NAMES)


@given(DIRTY_TEXT)
def test_fuzzy_campus_never_raises(text):
    assert isinstance(normalize_campus(text), str)


# ------------------------------------------------------------------ 分摊


@given(st.integers(min_value=0, max_value=5000),
       st.lists(st.integers(min_value=0, max_value=200),
                min_size=1, max_size=8))
@settings(max_examples=200, deadline=None)
def test_allocate_sum_always_equals_total(total, weights):
    assert sum(allocate(total, weights)) == max(total, 0)


@given(st.integers(min_value=0, max_value=5000),
       st.lists(st.integers(min_value=0, max_value=200),
                min_size=1, max_size=8))
@settings(max_examples=200, deadline=None)
def test_allocate_never_negative(total, weights):
    assert all(v >= 0 for v in allocate(total, weights))


@given(st.integers(min_value=0, max_value=300),
       st.lists(st.integers(min_value=1, max_value=100),
                min_size=1, max_size=5))
@settings(max_examples=100, deadline=None)
def test_room_allocation_sum_equals_college_exempt(total, weights):
    """教室分摊值之和恒等于学院总免勤（未触发封顶时）。"""
    share = allocate(total, weights)
    assert sum(share) == total
    assert len(share) == len(weights)


# ------------------------------------------------------------------ 去重


@given(st.lists(st.tuples(st.integers(min_value=0, max_value=5),
                          st.sampled_from(['周一', '周三', '周四', '周五']),
                          st.sampled_from(['早', '晚'])),
                min_size=1, max_size=25))
@settings(max_examples=60, deadline=None)
def test_same_sid_counted_at_most_once_per_day_session(work_dir, combos):
    """同一学号在同一 (星期, 自习) 最多计 1 次。"""
    rows = []
    for sid_idx, day, session in combos:
        rows.append(make_exempt_row(sid='99%011d' % (sid_idx + 1),
                                    day_text=day, session=session))
    base = work_dir('dedup')
    records = exempt_records(base, rows)
    agg = aggregate_exemptions(records, 7, campus='雁山')
    for day, session in {(c[1], c[2]) for c in combos}:
        session_name = '早自习' if session == '早' else '晚自习'
        distinct = {rows[i].sid for i, c in enumerate(combos)
                    if c[1] == day and c[2] == session}
        assert agg.exempt_count('理学院', day, session_name) == len(distinct)


# ------------------------------------------------------------------ 分母 / 空值


@given(st.integers(min_value=1, max_value=200))
@settings(max_examples=20, deadline=None)
def test_no_rate_when_denominator_nonpositive(work_dir, expected):
    """分母 ≤ 0 时不产出任何出勤率数值。"""
    base = work_dir('denom')
    detail = write_detail(base / 'd.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('理学院', expected, [('8101', expected, {'周一': max(expected - 1, 0)})]),
        ]),
    ])
    exempt = write_exempt(base / 'e.xlsx', {
        EARLY: [make_exempt_row(sid='99%011d' % i) for i in range(1, expected + 1)],
    })
    model = load_detail(str(detail))
    agg = aggregate_exemptions(load_exemptions(str(exempt)), 7, campus='雁山')
    from core.calc import Calculator
    rates = Calculator(model, agg, alias={}).college_rates()
    assert rates
    for rec in rates:
        if rec.exempt >= rec.expected:
            assert rec.rate is None
            assert rec.status == '分母非正数，无法计算'


@given(st.integers(min_value=1, max_value=200))
@settings(max_examples=20, deadline=None)
def test_missing_actual_yields_no_number(work_dir, expected):
    """实到缺失 → 出勤率为空值（None），绝不是 0。"""
    base = work_dir('missing')
    detail = write_detail(base / 'd.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('理学院', expected, [('8101', expected, {'周一': None})]),
        ]),
    ])
    exempt = write_exempt(base / 'e.xlsx', {EARLY: []})
    out = base / 'out.xlsx'
    result = run_engine(detail, exempt, output_path=out)
    assert result.ok, result.error
    for rec in result.college_rates:
        if rec.rate is None:
            assert rec.actual is None
    wb = load_workbook(out, data_only=True)
    cell = wb['雁山（早）'].cell(4, 14)
    assert cell.value is None
    assert cell.value != 0
    wb.close()


# ------------------------------------------------------------------ 输出结构 / 输入保护


@given(st.lists(st.text(alphabet='临时表甲乙丙丁', min_size=1, max_size=4),
                min_size=0, max_size=3))
@settings(max_examples=6, deadline=None, suppress_health_check=[HealthCheck.too_slow])
def test_output_sheets_always_three(work_dir, extra_names):
    """输出工作表名单恒等于三张原表。"""
    base = work_dir('sheets')
    detail, exempt = build_focused_sample(base)
    wb = load_workbook(detail)
    for name in extra_names:
        if name not in wb.sheetnames:
            wb.create_sheet(name)
    wb.save(detail)
    wb.close()

    out = base / 'out.xlsx'
    result = run_engine(detail, exempt, output_path=out)
    assert result.ok, result.error
    check = load_workbook(out)
    assert tuple(check.sheetnames) == ALLOWED_SHEETS
    check.close()


@given(st.integers(min_value=1, max_value=20))
@settings(max_examples=5, deadline=None)
def test_input_md5_unchanged_after_run(work_dir, week_no):
    """跑批后输入文件 md5 不变。"""
    base = work_dir('md5')
    detail, exempt = build_focused_sample(base, early_rows=[make_exempt_row()])
    before = (md5(detail), md5(exempt))
    out = base / ('out_%d.xlsx' % week_no)
    result = run_engine(detail, exempt, output_path=out, week_no=week_no)
    assert result.ok, result.error
    assert (md5(detail), md5(exempt)) == before
    assert out.exists()
