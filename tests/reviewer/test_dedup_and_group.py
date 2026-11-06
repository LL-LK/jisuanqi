# -*- coding: utf-8 -*-
"""免勤去重、集体申请覆盖、校区排除回归（需求原文第 3/4/5 组）。

去重键 = (学号, 星期, 早晚自习)。三条铁律：
  * 同键只计一次；
  * 同一学号早自习与晚自习各计一次；
  * 跨日期（不同星期）分别计。
集体申请：覆盖同班同日同自习的个人申请，且**不与个人重复相加**。
"""

from __future__ import annotations

import pytest

from core.io_exempt import aggregate_exemptions

import synthetic

COLLEGE = '测试学院甲'
OTHER_COLLEGE = '测试学院乙'
CLASS_A = '测试专业甲2026-1班'
CLASS_C = '测试专业丙2026-1班'


def _aggregate(records, week_no=7):
    return aggregate_exemptions(records, week_no, campus='雁山')


# ------------------------------------------------------------------ 去重


def test_same_sid_same_day_session_counted_once():
    """同学号 + 同星期 + 同自习 → 只计一次。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1


def test_same_sid_different_week_notation_counted_once():
    """跨周次写法命中同一周 → 仍只计一次（真实数据里冗余 323 条的主要来源）。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              weeks='第7周'),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              weeks='第6-7，9-13周'),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              weeks='第7-18周'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1


def test_same_sid_early_and_late_counted_twice():
    """同一人周一早自习与周一晚自习各计一次。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1),
                              college=COLLEGE, class_text=CLASS_A, session='早自习'),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              session='晚自习'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1
    assert agg.exempt_count(COLLEGE, '周一', '晚自习') == 1


def test_same_sid_different_days_counted_separately():
    """跨日期分别计。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              time='周一'),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              time='周三'),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              time='周四'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1
    assert agg.exempt_count(COLLEGE, '周三', '早自习') == 1
    assert agg.exempt_count(COLLEGE, '周四', '早自习') == 1


def test_distinct_sids_counted_separately():
    records = [synthetic.make_record(sid=synthetic.fake_sid(i), college=COLLEGE,
                                     class_text=CLASS_A) for i in range(1, 6)]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 5


def test_records_outside_week_window_are_filtered():
    """周次窗口之外的申请不计入第 7 周。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              weeks='第6周'),
        synthetic.make_record(sid=synthetic.fake_sid(2), college=COLLEGE, class_text=CLASS_A,
                              weeks='第8周'),
        synthetic.make_record(sid=synthetic.fake_sid(3), college=COLLEGE, class_text=CLASS_A,
                              weeks='第6-7，9-13周'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1
    assert agg.week_filtered == 2


def test_pingfeng_records_are_excluded_entirely():
    """屏风校区完全排除：不计免勤、不进汇总。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              campus='屏风校区'),
        synthetic.make_record(sid=synthetic.fake_sid(2), college=COLLEGE, class_text=CLASS_A,
                              campus='屏风'),
        synthetic.make_record(sid=synthetic.fake_sid(3), college=COLLEGE, class_text=CLASS_A,
                              campus='雁山校区'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1
    assert agg.campus_filtered == 2


def test_campus_typo_variants_are_counted_as_yanshan():
    """'雁山校区'/'雁山校园'/'燕山校区' 归一后必须照常计入。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              campus='雁山校区'),
        synthetic.make_record(sid=synthetic.fake_sid(2), college=COLLEGE, class_text=CLASS_A,
                              campus='雁山校园'),
        synthetic.make_record(sid=synthetic.fake_sid(3), college=COLLEGE, class_text=CLASS_A,
                              campus='燕山校区'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 3
    assert agg.campus_filtered == 0


def test_unparsable_week_is_flagged_not_silently_dropped():
    """真正无法识别的周次必须进异常清单（不得静默消失）。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              weeks='待定'),
    ]
    agg = _aggregate(records)
    assert agg.unmatched_week == 1
    assert any('周次' in item['type'] for item in agg.anomalies)


def test_whole_semester_covers_window_week():
    """「全学期」按裁定 R4 覆盖统计窗口周，不算无法识别。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A,
                              weeks='全学期'),
    ]
    agg = _aggregate(records)
    assert agg.unmatched_week == 0
    assert agg.exempt_count(COLLEGE, '周一', '早自习') == 1


# ------------------------------------------------------------------ 集体申请


def test_group_application_covers_personal_and_is_not_added_twice():
    """集体 43 人覆盖同班个人 28 人 → 学院免勤 = 43（不是 71）。"""
    records = [synthetic.make_record(name='测试专业丙2026-1班全体学生（共43人）',
                                     college='测试学院丙', class_text=CLASS_C,
                                     time='周三', session='晚自习')]
    records += [synthetic.make_record(sid=synthetic.fake_sid(i), college='测试学院丙',
                                      class_text=CLASS_C, time='周三', session='晚自习')
                for i in range(1, 29)]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 43
    assert agg.covered_personal == 28


def test_group_application_does_not_cover_other_session():
    """集体申请只覆盖同班**同日同自习**的个人；早自习的个人照计。"""
    records = [synthetic.make_record(name='测试专业丙2026-1班全体学生（共43人）',
                                     college='测试学院丙', class_text=CLASS_C,
                                     time='周三', session='晚自习')]
    records += [synthetic.make_record(sid=synthetic.fake_sid(i), college='测试学院丙',
                                      class_text=CLASS_C, time='周三', session='早自习')
                for i in range(1, 8)]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 43
    assert agg.exempt_count('测试学院丙', '周三', '早自习') == 7


def test_group_application_does_not_cover_other_day():
    """集体申请不覆盖同班其它星期的个人。"""
    records = [synthetic.make_record(name='测试专业丙2026-1班全体学生（共43人）',
                                     college='测试学院丙', class_text=CLASS_C,
                                     time='周三', session='晚自习')]
    records += [synthetic.make_record(sid=synthetic.fake_sid(i), college='测试学院丙',
                                      class_text=CLASS_C, time='周四', session='晚自习')
                for i in range(1, 8)]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 43
    assert agg.exempt_count('测试学院丙', '周四', '晚自习') == 7


def test_group_application_covers_equivalent_class_writing():
    """集体申请写 '应用统计学2026-1班'、个人写 '应用统计学26-1班' 时也必须覆盖。"""
    records = [synthetic.make_record(name='测试专业丙2026-1班全体学生（共43人）',
                                     college='测试学院丙', class_text='测试专业丙2026-1班',
                                     time='周三', session='晚自习')]
    records += [synthetic.make_record(sid=synthetic.fake_sid(i), college='测试学院丙',
                                      class_text='测试专业丙26-1', time='周三', session='晚自习')
                for i in range(1, 6)]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 43
    assert agg.covered_personal == 5


def test_group_application_larger_declaration_wins():
    """同班同日同自习重复申报时取较大声明人数，并记异常。"""
    records = [
        synthetic.make_record(name='测试专业丙2026-1班全体学生（共43人）',
                              college='测试学院丙', class_text=CLASS_C,
                              time='周三', session='晚自习'),
        synthetic.make_record(name='测试专业丙2026-1班全体学生（共40人）',
                              college='测试学院丙', class_text=CLASS_C,
                              time='周三', session='晚自习'),
    ]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 43
    assert any('重复申报' in item['type'] for item in agg.anomalies)


def test_group_application_without_size_does_not_cover_personal():
    records = [synthetic.make_record(name='测试专业丙2026-1班全体学生',
                                     college='测试学院丙', class_text=CLASS_C,
                                     time='周三', session='晚自习')]
    records += [synthetic.make_record(sid=synthetic.fake_sid(i), college='测试学院丙',
                                      class_text=CLASS_C, time='周三', session='晚自习')
                for i in range(1, 4)]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 3
    assert agg.covered_personal == 0


def test_group_application_does_not_swallow_other_class_personal():
    """同学院非集体班级的个人申请必须照计。"""
    records = [synthetic.make_record(name='测试专业丙2026-1班全体学生（共43人）',
                                     college='测试学院丙', class_text=CLASS_C,
                                     time='周三', session='晚自习')]
    records += [synthetic.make_record(sid=synthetic.fake_sid(i), college='测试学院丙',
                                      class_text='测试专业丙2026-2班', time='周三',
                                      session='晚自习')
                for i in range(1, 4)]
    agg = _aggregate(records)
    assert agg.exempt_count('测试学院丙', '周三', '晚自习') == 46


def test_same_sid_across_colleges_counted_once():
    """同一学号出现在两个学院 → 全局只应计一次。"""
    records = [
        synthetic.make_record(sid=synthetic.fake_sid(1), college=COLLEGE, class_text=CLASS_A),
        synthetic.make_record(sid=synthetic.fake_sid(1), college=OTHER_COLLEGE, class_text=CLASS_A),
    ]
    agg = _aggregate(records)
    total = agg.exempt_count(COLLEGE, '周一', '早自习') + agg.exempt_count(OTHER_COLLEGE, '周一', '早自习')
    assert total == 1
