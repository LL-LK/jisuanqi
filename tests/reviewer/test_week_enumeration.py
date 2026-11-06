# -*- coding: utf-8 -*-
"""周次解析全枚举回归（需求原文第 1 组）。

不变量：**周次集合必须是并集，不允许出现「取最后一个区间/取最后一个周」的写法**。
因此每条用例既断言集合内容，也断言集合大小——只留第 9 周的实现会在这里红。
"""

from __future__ import annotations

import pytest

from core.parsing import parse_weeks

#: 需求点名的 7 种写法 → 期望周号集合（逐一断言，不做等价类合并）
MANDATED_WEEK_CASES = [
    ('第7周', {7}),
    ('第7-18周', set(range(7, 19))),
    ('第6-7，9-13周', {6, 7, 9, 10, 11, 12, 13}),
    ('第7—12周', set(range(7, 13))),
    ('第七周，第八周', {7, 8}),
    ('第7、8、9周', {7, 8, 9}),
    ('6-10周', {6, 7, 8, 9, 10}),
]


@pytest.mark.parametrize('text,expected', MANDATED_WEEK_CASES,
                         ids=[c[0] for c in MANDATED_WEEK_CASES])
def test_parse_weeks_mandated_cases(text, expected):
    """7 种周次写法逐一断言周号集合。"""
    assert parse_weeks(text) == expected


def test_parse_weeks_enumerated_keeps_all_three_weeks():
    """'第7、8、9周' 必须解析出三个周——旧写法只取最后一个，会退化成 {9}。"""
    result = parse_weeks('第7、8、9周')
    assert 7 in result and 8 in result and 9 in result
    assert len(result) == 3


def test_parse_weeks_range_keeps_both_ends():
    """区间写法必须含两端，且长度等于区间长度。"""
    result = parse_weeks('第7-18周')
    assert 7 in result and 18 in result
    assert len(result) == 12


def test_parse_weeks_multi_range_keeps_every_segment():
    """多区间写法不能只保留第一段或最后一段。"""
    result = parse_weeks('第6-7，9-13周')
    assert result == {6, 7} | set(range(9, 14))
    assert 13 in result and 6 in result


def test_parse_weeks_variants_of_same_meaning_agree():
    """同义写法必须得到同一个集合（防止不同写法走出不同分支）。"""
    assert parse_weeks('第7周') == parse_weeks('7周') == parse_weeks('第 7 周')
    assert parse_weeks('第7-12周') == parse_weeks('第7周-第12周')
    assert parse_weeks('第7到第12周') == parse_weeks('第7-12周')


def test_parse_weeks_chinese_numeral_range():
    """中文数字区间（'第七周到第八周'）必须展开为 {7, 8}。"""
    assert parse_weeks('第七周到第八周') == {7, 8}
    assert parse_weeks('第七周至第八周') == {7, 8}


def test_parse_weeks_empty_and_garbage():
    """无法识别时返回空集合（由上层记异常，不得抛异常）。"""
    for text in ('', None, '   ', '无', 'abc', '第周'):
        assert parse_weeks(text) == set()


def test_parse_weeks_never_raises_on_garbage():
    """脏输入不抛未捕获异常（性质化版本见 test_properties.py）。"""
    for text in ('第', '周', '第-周', '第7-', '-12周', '第7--8周', '第0周', '第99周'):
        assert isinstance(parse_weeks(text), set)


def test_parse_weeks_space_separated_enumeration():
    assert parse_weeks('第7周 第8周') == {7, 8}


def test_parse_weeks_chinese_numeral_range_without_prefix():
    assert parse_weeks('第六-十七周') == set(range(6, 18))
