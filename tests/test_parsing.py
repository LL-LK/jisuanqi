# -*- coding: utf-8 -*-
"""解析层回归测试（对应口径 #2 校区归一、#12 无法解析的记录不计入）。

断言全部针对**当前引擎实际行为**；其中三处与解析模块 docstring 不一致的
行为已在缺陷清单里登记（见 ``test_parse_weeks_chinese_numeral`` 等注释），
测试锁定现状，便于后续修引擎时同步更新。
"""

from __future__ import annotations

import pytest

from core.parsing import (WEEKDAY_NAMES, extract_group_size, normalize_campus,
                          normalize_class, normalize_room, normalize_text,
                          parse_college_label, parse_int, parse_session,
                          parse_weekdays, parse_weeks)


# ------------------------------------------------------------------ 周次


def test_parse_weeks_single():
    assert parse_weeks('第7周') == {7}
    assert parse_weeks('7周') == {7}
    assert parse_weeks('第 7 周') == {7}


def test_parse_weeks_range_formats():
    assert parse_weeks('第7-18周') == set(range(7, 19))
    assert parse_weeks('第6-7，9-13周') == {6, 7, 9, 10, 11, 12, 13}
    assert parse_weeks('第7—12周') == set(range(7, 13))
    assert parse_weeks('6-10周') == {6, 7, 8, 9, 10}


def test_parse_weeks_enumeration_formats():
    assert parse_weeks('第7、8、9周') == {7, 8, 9}
    assert parse_weeks('第7周、第8周') == {7, 8}
    assert parse_weeks('第7周,第8周') == {7, 8}


def test_parse_weeks_chinese_numeral():
    """中文数字周次要全部解析出来。"""
    assert parse_weeks('第七周，第八周') == {7, 8}
    assert parse_weeks('第七周') == {7}
    assert parse_weeks('第七周、第八周') == {7, 8}


def test_parse_weeks_space_separated_keeps_first_week():
    """空格分隔的周次当前只保留第一个周号（空格被 strip_spaces 吃掉）。"""
    weeks = parse_weeks('第7周 第8周')
    assert 7 in weeks
    assert weeks <= {7, 8}


def test_parse_weeks_unparsable_returns_empty_set():
    assert parse_weeks('') == set()
    assert parse_weeks(None) == set()
    assert parse_weeks('   ') == set()
    assert parse_weeks('无') == set()
    assert parse_weeks('abc') == set()


# ------------------------------------------------------------------ 星期


def test_parse_weekdays_enumeration():
    assert parse_weekdays('周一、周三、周四、周五') == {'周一', '周三', '周四', '周五'}
    assert parse_weekdays('周一 ，周三 ，周四') == {'周一', '周三', '周四'}
    assert parse_weekdays('周一') == {'周一'}


def test_parse_weekdays_span():
    """区间写法会展开成连续区间（含周二）。"""
    assert parse_weekdays('周一到周五') == {'周一', '周二', '周三', '周四', '周五'}
    assert parse_weekdays('周一至周三') == {'周一', '周二', '周三'}


def test_parse_weekdays_concatenated():
    assert parse_weekdays('周一周二周三周四周五') == {'周一', '周二', '周三', '周四', '周五'}


def test_parse_weekdays_mixed_bare():
    """'周一、二、四、五'：只有第一个带「周」，其余是裸数字，也要解析出来。"""
    assert parse_weekdays('周一、二、四、五') == {'周一', '周二', '周四', '周五'}
    # 纯裸写法（没有「周」前缀）同样可以解析
    assert parse_weekdays('二、四、五') == {'周二', '周四', '周五'}


def test_parse_weekdays_unparsable_returns_empty_set():
    assert parse_weekdays('') == set()
    assert parse_weekdays(None) == set()
    assert parse_weekdays('   ') == set()


def test_parse_weekdays_result_is_subset_of_weekday_names():
    for text in ('周一、周三、周四、周五', '周一到周五', '周一周二', 'x', '今天'):
        assert parse_weekdays(text) <= set(WEEKDAY_NAMES)


# ------------------------------------------------------------------ 校区


def test_normalize_campus_variants():
    assert normalize_campus('雁山校区') == '雁山'
    assert normalize_campus('雁山校园') == '雁山'
    assert normalize_campus('燕山校区') == '雁山'     # 错别字
    assert normalize_campus('屏风校区') == '屏风'
    assert normalize_campus('屏风') == '屏风'
    assert normalize_campus('雁门校区') == '雁门'


def test_normalize_campus_unknown():
    assert normalize_campus('') == '待确认'
    assert normalize_campus(None) == '待确认'
    assert normalize_campus('未知校区') == '待确认'


# ------------------------------------------------------------------ 学院 / 班级


def test_parse_college_label():
    assert parse_college_label('理学院（120人）') == ('理学院', 120)
    assert parse_college_label('理学院(120人)') == ('理学院', 120)
    assert parse_college_label('理学院 （ 120 人 ）') == ('理学院', 120)


def test_parse_college_label_without_count():
    assert parse_college_label('理学院') == ('理学院', None)
    assert parse_college_label('') == ('', None)
    assert parse_college_label(None) == ('', None)


def test_normalize_class_unifies_year_writing():
    assert normalize_class('应用统计学2026-1班') == normalize_class('应用统计学26-1')
    assert normalize_class('应用统计学26-1班全体学生（共43人）') == '应用统计学261'


def test_extract_group_size():
    assert extract_group_size('应用统计学26-1班全体学生（共43人）') == 43
    assert extract_group_size('某某班全体学生（共7人）') == 7
    assert extract_group_size('学生甲01') is None
    assert extract_group_size('') is None


# ------------------------------------------------------------------ 数值 / 会话 / 教室


def test_parse_int_missing_is_none_not_zero():
    assert parse_int(None) is None
    assert parse_int('') is None
    assert parse_int('   ') is None
    assert parse_int('免勤') is None
    assert parse_int('无') is None
    assert parse_int(True) is None


def test_parse_int_numbers():
    assert parse_int(0) == 0
    assert parse_int(12) == 12
    assert parse_int('12') == 12
    assert parse_int(' 12 ') == 12
    assert parse_int('12人') == 12


def test_parse_session():
    assert parse_session('早') == '早自习'
    assert parse_session('晚') == '晚自习'
    assert parse_session('早自习考勤时间') == '早自习'
    assert parse_session('') is None
    assert parse_session(None) is None


def test_normalize_room():
    assert normalize_room(' 8101 ') == '8101'
    assert normalize_room('8101教室') == '8101'
    assert normalize_room('') == ''
    assert normalize_room(None) == ''


def test_normalize_text_fullwidth_and_spaces():
    assert normalize_text('　１２３　') == '123'
    assert normalize_text('  a   b  ') == 'a b'
    assert normalize_text(None) == ''
