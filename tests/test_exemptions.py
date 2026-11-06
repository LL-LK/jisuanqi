# -*- coding: utf-8 -*-
"""免勤表读取 / 去重 / 聚合的回归测试。

对应口径：
  #2  屏风校区整体排除
  #6  去重键 = (学号, 星期, 早晚自习)，跨日期分别计，早晚各计一次
  #7  集体免勤申请覆盖同班同日同自习的个人申请，不重复相加
  #12 周次 / 星期无法解析的记录不计入，但必须出现在异常清单里

所有用例都走**真实链路**：先写成免勤表，再用 ``load_exemptions`` 读回来，
这样周次/星期/校区/自习的解析与归一化都在断言范围内。
"""

from __future__ import annotations

import pytest

from _helpers import (build_focused_sample, exempt_records,  # noqa: F401  夹具
                      make_exempt_row, write_exempt)
from core.io_exempt import aggregate_exemptions, load_exemptions

EARLY = '早自习免勤人员统计表'
LATE = '晚自习免勤人员统计表'


def _aggregate(tmp_path, rows, week_no=7, campus='雁山', sheet=EARLY):
    records = exempt_records(tmp_path, rows, sheet=sheet)
    return aggregate_exemptions(records, week_no, campus=campus)


# ------------------------------------------------------------------ 读取


def test_load_exemptions_reads_fictional_rows(tmp_path):
    path = write_exempt(tmp_path / 'exempt.xlsx', {
        EARLY: [make_exempt_row(sid='9900000000001', name='学生甲01')],
        LATE: [make_exempt_row(sid='9900000000002', name='学生乙02', session='晚')],
    })
    records = load_exemptions(str(path))
    assert len(records) == 2
    first = records[0]
    assert first.sid == '9900000000001'
    assert first.name == '学生甲01'
    assert first.college == '理学院'
    assert first.weeks == {7}
    assert first.weekdays == {'周一'}
    assert first.session == '早自习'
    assert first.campus == '雁山'
    assert first.is_personal is True
    assert first.is_group is False


def test_load_exemptions_detects_room_column(tmp_path):
    path = write_exempt(tmp_path / 'exempt_room.xlsx', {
        EARLY: [make_exempt_row(room='8101')],
    }, with_room=True)
    records = load_exemptions(str(path))
    assert records[0].room == '8101'


def test_load_exemptions_without_room_column(tmp_path):
    path = write_exempt(tmp_path / 'exempt_noroom.xlsx', {
        EARLY: [make_exempt_row()],
    })
    records = load_exemptions(str(path))
    assert records[0].room == ''


def test_load_exemptions_normalizes_dirty_formats(tmp_path):
    """脏周次 / 星期 / 校区写法都要被归一化。"""
    path = write_exempt(tmp_path / 'exempt_dirty.xlsx', {
        EARLY: [
            make_exempt_row(sid='9900000000001', week_text='第七周，第八周',
                            day_text='周一、二、四、五', campus='燕山校区'),
            make_exempt_row(sid='9900000000002', name='学生乙02',
                            week_text='第6-7，9-13周', day_text='周一到周五',
                            campus='雁山校园'),
        ],
    })
    records = load_exemptions(str(path))
    assert records[0].weeks == {7, 8}
    assert records[0].weekdays == {'周一', '周二', '周四', '周五'}
    assert records[0].campus == '雁山'
    assert records[1].weeks == {6, 7, 9, 10, 11, 12, 13}
    assert records[1].weekdays == {'周一', '周二', '周三', '周四', '周五'}


# ------------------------------------------------------------------ 去重


def test_same_sid_same_day_same_session_counts_once(tmp_path):
    rows = [make_exempt_row(sid='9900000000001'),
            make_exempt_row(sid='9900000000001', name='学生甲 1')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 1
    assert agg.duplicate_events == 1


def test_same_sid_different_days_counted_separately(tmp_path):
    rows = [make_exempt_row(sid='9900000000001', day_text='周一'),
            make_exempt_row(sid='9900000000001', day_text='周三')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 1
    assert agg.exempt_count('理学院', '周三', '早自习') == 1


def test_same_sid_early_and_late_counted_separately(tmp_path):
    rows = [make_exempt_row(sid='9900000000001', session='早'),
            make_exempt_row(sid='9900000000001', session='晚')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 1
    assert agg.exempt_count('理学院', '周一', '晚自习') == 1


def test_multi_day_record_contributes_to_each_day(tmp_path):
    rows = [make_exempt_row(sid='9900000000001', day_text='周一、周三')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 1
    assert agg.exempt_count('理学院', '周三', '早自习') == 1
    assert agg.exempt_count('理学院', '周四', '早自习') == 0


def test_different_sid_same_day_counted_independently(tmp_path):
    rows = [make_exempt_row(sid='9900000000001'),
            make_exempt_row(sid='9900000000002', name='学生乙02')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 2
    assert agg.unique_sids == 2


# ------------------------------------------------------------------ 集体申请


def test_group_application_covers_personal(tmp_path):
    rows = [
        make_exempt_row(sid='', name='应用统计学26-1班全体学生（共5人）',
                        class_text='应用统计学26-1班'),
        make_exempt_row(sid='9900000000001', class_text='应用统计学26-1班'),
        make_exempt_row(sid='9900000000002', name='学生乙02',
                        class_text='应用统计学2026-1班'),
    ]
    agg = _aggregate(tmp_path, rows)
    # 集体声明 5 人，个人申请被覆盖，不再重复相加
    assert agg.exempt_count('理学院', '周一', '早自习') == 5
    assert agg.group_applications == 1
    assert agg.covered_personal == 2


def test_group_and_personal_of_other_class_both_count(tmp_path):
    rows = [
        make_exempt_row(sid='', name='应用统计学26-1班全体学生（共5人）',
                        class_text='应用统计学26-1班'),
        make_exempt_row(sid='9900000000001', class_text='应用统计学26-2班'),
    ]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 6
    assert agg.covered_personal == 0


def test_duplicate_group_takes_larger_size_and_reports(tmp_path):
    rows = [
        make_exempt_row(sid='', name='应用统计学26-1班全体学生（共5人）',
                        class_text='应用统计学26-1班'),
        make_exempt_row(sid='', name='应用统计学26-1班全体学生（共9人）',
                        class_text='应用统计学26-1班'),
    ]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 9
    assert agg.duplicate_events == 1
    assert any(a['type'] == '集体免勤重复申报' for a in agg.anomalies)


def test_group_record_without_sid_is_not_personal(tmp_path):
    row = make_exempt_row(sid='', name='应用统计学26-1班全体学生（共5人）',
                          class_text='应用统计学26-1班')
    agg = _aggregate(tmp_path, [row])
    assert agg.exempt_count('理学院', '周一', '早自习') == 5
    assert agg.invalid_records == 0


# ------------------------------------------------------------------ 过滤


def test_pingfeng_campus_is_excluded(tmp_path):
    rows = [make_exempt_row(campus='屏风校区'),
            make_exempt_row(sid='9900000000002', name='学生乙02', campus='屏风')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 0
    assert agg.campus_filtered == 2


def test_other_campus_is_excluded(tmp_path):
    rows = [make_exempt_row(campus='雁门校区')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 0
    assert agg.campus_filtered == 1


def test_campus_variants_are_normalized_to_yanshan(tmp_path):
    rows = [make_exempt_row(sid='9900000000001', campus='雁山校园'),
            make_exempt_row(sid='9900000000002', name='学生乙02', campus='燕山校区')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 2
    assert agg.campus_filtered == 0


def test_other_weeks_are_filtered_out(tmp_path):
    rows = [make_exempt_row(week_text='第6周'),
            make_exempt_row(sid='9900000000002', name='学生乙02', week_text='第8周')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 0
    assert agg.week_filtered == 2


def test_week_range_containing_target_week_is_kept(tmp_path):
    rows = [make_exempt_row(week_text='第7-18周')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 1
    assert agg.week_filtered == 0


# ------------------------------------------------------------------ 无法解析


def test_unparsable_week_is_not_counted_but_reported(tmp_path):
    rows = [make_exempt_row(week_text='')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 0
    assert agg.unmatched_week == 1
    assert any(a['type'] == '免勤周次无法识别' for a in agg.anomalies)


def test_unparsable_day_is_not_counted_but_reported(tmp_path):
    rows = [make_exempt_row(day_text='')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 0
    assert agg.unmatched_day == 1
    assert any(a['type'] == '免勤星期无法识别' for a in agg.anomalies)


def test_record_without_sid_and_without_group_is_invalid(tmp_path):
    rows = [make_exempt_row(sid='', name='学生甲01')]
    agg = _aggregate(tmp_path, rows)
    assert agg.exempt_count('理学院', '周一', '早自习') == 0
    assert agg.invalid_records == 1
    assert any(a['type'] == '免勤记录无法定位' for a in agg.anomalies)


def test_aggregate_stats_are_consistent(tmp_path):
    rows = [make_exempt_row(sid='9900000000001'),
            make_exempt_row(sid='9900000000001', name='学生甲 1'),
            make_exempt_row(sid='9900000000002', name='学生乙02', week_text='第6周'),
            make_exempt_row(sid='9900000000003', name='学生丙03', campus='屏风校区'),
            make_exempt_row(sid='9900000000004', name='学生丁04', week_text='')]
    agg = _aggregate(tmp_path, rows)
    assert agg.total_records == 5
    assert agg.duplicate_events == 1
    assert agg.week_filtered == 1
    assert agg.campus_filtered == 1
    assert agg.unmatched_week == 1
    assert agg.exempt_count('理学院', '周一', '早自习') == 1
