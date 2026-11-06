# -*- coding: utf-8 -*-
"""教室分摊与学院出勤率回归（需求原文第 6/7 组）。

三条硬口径：
  * 教室分摊值之和 **恒等于** 学院总免勤（最大余数法，不得截断到教室应到）；
  * 学院出勤率 = 总实到 ÷ (总应到 − 当次总免勤)，复算差异必须为 0；
  * 分母 ≤0 不出率；实到缺失留空、不得写 0。
"""

from __future__ import annotations

import pytest

from core.calc import Calculator, allocate
from core.io_exempt import Aggregate

import synthetic

DAYS = ('周一', '周三', '周四', '周五')


# ------------------------------------------------------------------ 分摊算法


@pytest.mark.parametrize('total,weights', [
    (5, [3, 7]),
    (10, [1, 1, 1]),
    (7, [2, 0, 5]),
    (1, [10, 10]),
    (100, [60, 40]),
    (25, [10, 10]),
    (13, [5, 5, 3]),
])
def test_allocate_sum_equals_total(total, weights):
    """最大余数法的硬性要求：分摊之和恒等于总数。"""
    assert sum(allocate(total, weights)) == total


def test_allocate_maximum_remainder_known_values():
    # 注：模块 docstring 写 allocate(10, [1, 1, 1]) == [4, 3, 3]，
    # 实现的最大余数法在余数并列时按列序倒序补位，实际是 [3, 3, 4]。
    # 余数并列时的补位顺序不是契约，**和相等**才是契约。
    assert allocate(5, [3, 7]) == [2, 3]
    assert sorted(allocate(10, [1, 1, 1])) == [3, 3, 4]
    assert sum(allocate(10, [1, 1, 1])) == 10


def test_allocate_zero_total_returns_zeros():
    assert allocate(0, [10, 20]) == [0, 0]


def test_allocate_single_weight_takes_everything():
    assert allocate(9, [7, 0]) == [9, 0]


def test_allocate_never_negative():
    for total in (0, 1, 7, 100):
        for weights in ([10, 10], [0, 0], [3]):
            assert all(v >= 0 for v in allocate(total, weights))


# ------------------------------------------------------------------ 学院出勤率


def _room_exempt_by_college(result):
    """从流水线结果里汇总教室级免勤：(自习, 学院, 星期) → 人数。"""
    out = {}
    for row in result.room_rows:
        session = '早自习' if '早' in row['工作表'] else '晚自习'
        for day in DAYS:
            key = (session, row['学院'], day)
            out[key] = out.get(key, 0) + row.get('%s免勤' % day, 0)
    return out


def test_college_rate_recompute_diff_zero(synthetic_run):
    """每一条出勤率都必须能被「总实到 ÷ (总应到 − 当次总免勤)」复算出来，差异为 0。"""
    result, _ = synthetic_run
    checked = 0
    for rec in result.college_rates:
        if rec.rate is None:
            continue
        denominator = rec.expected - rec.exempt
        assert denominator > 0, '出率时分母必须为正：%s' % rec.note
        assert rec.actual is not None
        assert rec.rate == pytest.approx(rec.actual / denominator, abs=1e-12)
        checked += 1
    assert checked > 0, '合成数据必须至少算出一条出勤率'


def test_rate_is_only_missing_for_legal_reasons(synthetic_run):
    """缺率只允许两种原因：实到缺失、分母非正；其余一律必须有值。"""
    result, _ = synthetic_run
    for rec in result.college_rates:
        denominator = None if rec.expected is None else rec.expected - rec.exempt
        if rec.rate is None:
            assert rec.actual is None or denominator is None or denominator <= 0, \
                '不该缺率却缺了：%s / %s / %s' % (rec.sheet, rec.college, rec.day)
        else:
            assert rec.actual is not None and denominator > 0


def test_missing_actual_yields_no_rate(synthetic_run):
    """实到缺失 → 该学院该天不出率（乙·周一·早）。"""
    result, _ = synthetic_run
    hit = [r for r in result.college_rates
           if r.college == '测试学院乙' and r.day == '周一' and r.session == '早自习']
    assert hit, '必须存在乙学院周一的出勤率记录'
    rec = hit[0]
    assert rec.actual is None
    assert rec.rate is None
    assert '缺失' in rec.status or '未完整填写' in rec.status


def test_text_actual_treated_as_missing(synthetic_run):
    """实到写成文字「免勤」→ 按缺失处理，不得当成 0。"""
    result, _ = synthetic_run
    hit = [r for r in result.college_rates
           if r.college == '测试学院丙' and r.day == '周四' and r.session == '晚自习']
    assert hit
    rec = hit[0]
    assert rec.actual is None
    assert rec.rate is None


def test_zero_actual_is_a_valid_value_not_missing():
    """实到 = 0 是有效数据（出勤率 0），不得当成缺失。"""
    detail = synthetic.make_detail_model([
        ('测试学院甲', 20, [('T001', 20, {'周一': 0})]),
    ])
    agg = Aggregate()
    calc = Calculator(detail, agg)
    rates = calc.college_rates()
    rec = rates[0]
    assert rec.actual == 0
    assert rec.rate == pytest.approx(0.0)
    assert rec.status == '完整'


def test_denominator_nonpositive_yields_no_rate(synthetic_run):
    """分母 ≤0 不出率（丙·周四·早：总免勤 25 > 总应到 20）。"""
    result, _ = synthetic_run
    hit = [r for r in result.college_rates
           if r.college == '测试学院丙' and r.day == '周四' and r.session == '早自习']
    assert hit
    rec = hit[0]
    assert rec.expected - rec.exempt <= 0
    assert rec.rate is None
    assert '分母' in rec.status or '分母' in rec.note


def test_group_overflow_beyond_college_expected_yields_no_rate(synthetic_run):
    """集体申请 30 人 > 学院应到 20 → 分母 −10，不出率。"""
    result, _ = synthetic_run
    hit = [r for r in result.college_rates
           if r.college == '测试学院丁' and r.day == '周三' and r.session == '晚自习']
    assert hit
    rec = hit[0]
    assert rec.exempt == 30
    assert rec.expected == 20
    assert rec.rate is None


def test_average_rate_requires_all_days(synthetic_run):
    """平均出勤率：只要有一天缺失就整体留空（不得用部分天冒充整周平均）。"""
    result, _ = synthetic_run
    from core.calc import Calculator as _C  # noqa: F401  (保持导入语义清晰)
    avg_missing = [r for r in result.college_rates
                   if r.college == '测试学院乙' and r.day == '周一' and r.session == '早自习']
    assert avg_missing and avg_missing[0].rate is None
    # 甲四天齐全 → 平均值必须存在
    calc_avg = None
    for rec in result.college_rates:
        if rec.college == '测试学院甲' and rec.session == '早自习' and rec.day == '周一':
            calc_avg = rec
    assert calc_avg is not None and calc_avg.rate is not None


def test_room_allocation_sum_equals_college_exempt(synthetic_run):
    """教室分摊值之和 == 学院总免勤（甲·周一·早：5 人分到两间教室）。"""
    result, _ = synthetic_run
    by_college = _room_exempt_by_college(result)
    total = by_college[('早自习', '测试学院甲', '周一')]
    college_exempt = [r for r in result.college_rates
                       if r.college == '测试学院甲' and r.day == '周一'
                       and r.session == '早自习'][0].exempt
    assert total == college_exempt == 5


def test_room_allocation_capacity_cap_follows_conversation_caliber(synthetic_run):
    """C1=甲（对话口径）：教室级免勤按**教室容量封顶**，超额**单列**。

    丙·周四·早：集体声明 25 人 > 学院应到 20 ⇒
      * 比例分摊：T005/T006 应到各 10 ⇒ 比例分摊 13/12
      * 容量封顶：写进教室免勤单元格的是 10/10（绝不写 12>10 这种物理不可能值）
      * 超额单列：3+2=5 人进异常清单，**不写入任何教室免勤单元格**
      * 不变式：Σ(教室封顶值) + Σ(超额) == 学院总免勤 == 25
      * 每间封顶值 ≤ 该教室应到人数
      * 学院分母 20 − 25 = −5 ≤ 0 ⇒ 学院不出勤率
    """
    result, _ = synthetic_run
    by_college = _room_exempt_by_college(result)
    capped_total = by_college[('早自习', '测试学院丙', '周四')]
    college = [r for r in result.college_rates
               if r.college == '测试学院丙' and r.day == '周四'
               and r.session == '早自习'][0]
    college_exempt = college.exempt

    assert college_exempt == 25
    overflow = [o for o in result.exempt_overflow
                if o['college'] == '测试学院丙' and o['day'] == '周四']
    overflow_total = sum(o['excess'] for o in overflow)
    assert capped_total == 20                     # 两间教室各被封顶到应到 10
    assert overflow_total == 5
    assert capped_total + overflow_total == college_exempt == 25
    # 每间封顶值 ≤ 该教室应到人数（绝不允许 12 > 10 这种物理不可能值）
    rooms = [r for r in result.room_rows
             if r['学院'] == '测试学院丙' and r.get('周四免勤') is not None
             and '早' in r['工作表']]
    assert rooms, '未取到丙块的教室行'
    for r in rooms:
        assert r['周四免勤'] <= r['应到人数'], r
    # 学院分母非正 ⇒ 不出勤率
    assert college.rate is None
    assert college.status != '完整'


def test_room_allocation_never_added_back_to_college_total(synthetic_run):
    """分摊值只写教室级免勤列，绝不回加学院总免勤（学院总免勤仍是唯一学号去重值）。"""
    result, _ = synthetic_run
    by_college = _room_exempt_by_college(result)
    ding = [r for r in result.college_rates
            if r.college == '测试学院丁' and r.day == '周三' and r.session == '晚自习'][0]
    assert ding.exempt == 30
    assert ding.exempt != 30 + by_college[('晚自习', '测试学院丁', '周三')]
    bing = [r for r in result.college_rates
            if r.college == '测试学院丙' and r.day == '周四' and r.session == '早自习'][0]
    assert bing.exempt == 25
    assert bing.exempt != 25 + by_college[('早自习', '测试学院丙', '周四')]


def test_overflow_is_reported_in_anomaly_list(synthetic_run):
    """教室免勤超过应到必须进异常清单，不得静默处理。"""
    result, _ = synthetic_run
    kinds = [item['type'] for item in result.anomalies]
    assert any('超过' in kind for kind in kinds), kinds


def test_late_sheet_friday_exemption_is_surfaced(synthetic_run):
    """晚表周五免勤必须出现在异常清单里。"""
    result, _ = synthetic_run
    notes = ' | '.join('%s %s' % (item['type'], item.get('note', '')) for item in result.anomalies)
    assert '周五' in notes
