# -*- coding: utf-8 -*-
"""星期解析全枚举回归（需求原文第 2 组）。

不变量：**星期集合必须是并集**——'周一、周三、周四、周五' 不能退化成 {周五}。
"""

from __future__ import annotations

import pytest

from core.parsing import WEEKDAY_NAMES, parse_weekdays

#: 需求点名的 5 种写法 → 期望星期集合
MANDATED_WEEKDAY_CASES = [
    ('周一、周三、周四、周五', {'周一', '周三', '周四', '周五'}),
    ('周一 ，周三 ，周四', {'周一', '周三', '周四'}),
    ('周一到周五', {'周一', '周二', '周三', '周四', '周五'}),
    ('周一周二周三周四周五', {'周一', '周二', '周三', '周四', '周五'}),
    ('周一、二、四、五', {'周一', '周二', '周四', '周五'}),
]


@pytest.mark.parametrize('text,expected', MANDATED_WEEKDAY_CASES,
                         ids=[c[0] for c in MANDATED_WEEKDAY_CASES])
def test_parse_weekdays_mandated_cases(text, expected):
    """5 种星期写法逐一断言星期集合。"""
    assert parse_weekdays(text) == expected


def test_parse_weekdays_four_days_all_kept():
    """四天枚举不能只剩最后一天（旧写法取最后一个星期）。"""
    result = parse_weekdays('周一、周三、周四、周五')
    assert len(result) == 4
    for day in ('周一', '周三', '周四', '周五'):
        assert day in result


def test_parse_weekdays_mixed_bare_keeps_all_four():
    """'周一、二、四、五' 必须解析出 4 天（'一' 不能把后面的裸数字吃掉）。"""
    result = parse_weekdays('周一、二、四、五')
    assert len(result) == 4
    assert result == {'周一', '周二', '周四', '周五'}


def test_parse_weekdays_span_includes_middle_days():
    """区间写法必须补齐中间日（周一到周五含周二、周三、周四）。"""
    result = parse_weekdays('周一到周五')
    assert result == set(WEEKDAY_NAMES[:5])


def test_parse_weekdays_span_reversed_is_normalized():
    """反向区间与正向区间等价。"""
    assert parse_weekdays('周五到周一') == parse_weekdays('周一到周五')


def test_parse_weekdays_concatenated_keeps_five():
    """无分隔连续写必须切出 5 天。"""
    assert len(parse_weekdays('周一周二周三周四周五')) == 5


def test_parse_weekdays_whitespace_and_fullwidth_separator():
    """空格 + 全角逗号混排不得丢字段。"""
    assert parse_weekdays('周一 ，周三 ，周四') == {'周一', '周三', '周四'}
    assert parse_weekdays('周一,周三') == {'周一', '周三'}


def test_parse_weekdays_bare_chinese_without_zhou():
    """缺「周」字的中文星期（'第一 三 四'）必须解析出周一/周三/周四。"""
    assert parse_weekdays('第一 三 四') == {'周一', '周三', '周四'}


def test_parse_weekdays_empty_and_garbage():
    for text in ('', None, '   ', 'abc'):
        assert parse_weekdays(text) == set()


def test_parse_weekdays_does_not_treat_today_as_sunday():
    """「今天」「全天」这类说法不得被误判为周日免勤。"""
    assert parse_weekdays('今天') == set()
    assert parse_weekdays('全天') == set()


def test_parse_weekdays_result_is_subset_of_weekday_names():
    for text, _expected in MANDATED_WEEKDAY_CASES:
        assert parse_weekdays(text) <= set(WEEKDAY_NAMES)
