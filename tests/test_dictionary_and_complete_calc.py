# -*- coding: utf-8 -*-
"""词典完善度与出勤率 100% 完整计算保障专项测试套件。"""
from __future__ import annotations

import openpyxl
import pytest

from _helpers import make_detail_sheet, write_detail
from core.calc import Calculator
from core.dictionary import Dictionary, create_dictionary
from core.io_detail import load_detail
from core.io_exempt import aggregate_exemptions
from core.model import ExemptRecord


# ---------------------------------------------------------------------------
# 1. 词典 GUT 学院全称、简称、错别字与最长模糊兜底
# ---------------------------------------------------------------------------
def test_dictionary_gut_colleges_all_recognized(tmp_path):
    dict_file = tmp_path / "dict_full.xlsx"
    create_dictionary(str(dict_file))
    d = Dictionary.load(str(dict_file))

    cases = {
        # 常见简称
        "化生学院": "化学与生物工程学院",
        "化生院": "化学与生物工程学院",
        "化生": "化学与生物工程学院",
        "土木学院": "土木工程学院",
        "土建学院": "土木工程学院",
        "土木与建筑工程学院": "土木工程学院",
        "地科学院": "地球科学学院",
        "地学学院": "地球科学学院",
        "地院": "地球科学学院",
        "机控学院": "机械与控制工程学院",
        "机电学院": "机械与控制工程学院",
        "计算机学院": "计算机科学与工程学院",
        "计科院": "计算机科学与工程学院",
        "计网学院": "计算机科学与工程学院",
        "材料学院": "材料科学与工程学院",
        "材工学院": "材料科学与工程学院",
        "测绘学院": "测绘地理信息学院",
        "物电学院": "物理与电子信息工程学院",
        "环境学院": "环境科学与工程学院",
        "环工学院": "环境科学与工程学院",
        "旅游学院": "旅游与风景园林学院",
        "旅风学院": "旅游与风景园林学院",
        "数统学院": "数学与统计学院",
        "外语学院": "外国语学院",
        "经管学院": "商学院",
        "艺术学院": "艺术学院",
        "亚太学院": "亚太信息技术学院",
        "马院": "马克思主义学院",
        "公管学院": "公共管理与传媒学院",
        # 常见填报漏字与错别字
        "化学与工程学院": "化学与生物工程学院",
        "化学与生物工程": "化学与生物工程学院",
        "地球学院": "地球科学学院",
        "材料科学与工程x": "材料科学与工程学院",
        "环境科学院工程学院": "环境科学与工程学院",
        # 智能最长模糊兜底（未在词典别名表直接写明的复合长文本）
        "某高校化生学院学生会": "化学与生物工程学院",
        "机控卓越班": "机械与控制工程学院",
        "测绘工程专升本": "测绘地理信息学院",
    }

    for input_name, expected_formal in cases.items():
        assert d.college(input_name) == expected_formal, f"Failed for {input_name}"


# ---------------------------------------------------------------------------
# 2. 校区别名与错别字识别
# ---------------------------------------------------------------------------
def test_dictionary_campus_variants(tmp_path):
    dict_file = tmp_path / "dict_full.xlsx"
    create_dictionary(str(dict_file))
    d = Dictionary.load(str(dict_file))

    yanshan_cases = ["雁山", "雁山校区", "雁山区", "雁山校园", "燕山校区", "雁三校区", "雁山县区", "雁山校區"]
    for c in yanshan_cases:
        assert d.campus(c) == "雁山", f"Failed for campus {c}"

    pingfeng_cases = ["屏风", "屏风校区", "屏峰校区"]
    for c in pingfeng_cases:
        assert d.campus(c) == "屏风", f"Failed for campus {c}"
    assert d.campus("雁门校区") == "雁门"


# ---------------------------------------------------------------------------
# 3. 免勤理由子串分类覆盖
# ---------------------------------------------------------------------------
def test_dictionary_reason_categories(tmp_path):
    dict_file = tmp_path / "dict_full.xlsx"
    create_dictionary(str(dict_file))
    d = Dictionary.load(str(dict_file))

    cases = {
        "本周二晚有选修课冲突": "课务学业",
        "上无机化学实验课": "课务学业",
        "第55届校运会体育代表队集训": "文体活动",
        "入党积极分子优推考核": "党团事务",
        "院学生会团委部门值班": "公共事务",
        "参加全校大学生体测": "公共事务",
        "因急性肠胃炎在医院挂水就医": "健康医疗",
        "市人民医院拔牙就诊": "健康医疗",
        "因家里有事请假返校车票冲突": "因事请假",
    }
    for reason, expected_cat in cases.items():
        assert d.reason_class_of(reason) == expected_cat, f"Failed for reason {reason}"


# ---------------------------------------------------------------------------
# 4. 实到栏填写“免勤”/“请假”文本折算为 0 实到参与出勤率计算
# ---------------------------------------------------------------------------
def test_calc_room_mianqin_text_as_zero_and_calculates(tmp_path):
    blocks = [
        ("亚太信息技术学院", 200, [
            ("4101", 100, {"周一": 90}),
            ("4102", 100, {"周一": "免勤"}),  # 填文字“免勤”，折算为 0 实到参与计算
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_text.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions([], week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=True)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]
    # 实到合计 = 90 + 0 = 90，分母 = 200 - 0 = 200，率 = 45.0%
    assert rate.rate == pytest.approx(0.45)
    assert rate.actual == 90
    assert "4102" in rate.note
    assert "已折算为0实到计入" in rate.note


# ---------------------------------------------------------------------------
# 5. 部分教室实到未录入（None）时按有效教室实到计算，出勤率恒定产出
# ---------------------------------------------------------------------------
def test_calc_partial_missing_rooms_calculates_with_warning(tmp_path):
    # 3 间教室，其中 4203 留空（未填实到）
    blocks = [
        ("地球科学学院", 300, [
            ("4201", 100, {"周一": 95}),
            ("4202", 100, {"周一": 85}),
            ("4203", 100, {"周一": None}),  # 留空未填
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_partial.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions([], week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=True)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]
    # 出勤率不为 None！实到 = 95 + 85 = 180，分母 = 300，出勤率 = 60.0%
    assert rate.rate == pytest.approx(0.60)
    assert rate.actual == 180
    assert "部分教室缺失（已按有效实到计算）" in rate.status
    assert "4203" in rate.note
    assert any("未录入" in a for a in rate.anomalies)


# ---------------------------------------------------------------------------
# 6. 分母 <= 0（全员免勤或超额免勤）时按 100.0% 满勤出勤计入
# ---------------------------------------------------------------------------
def test_calc_denominator_nonpositive_as_full_attendance(tmp_path):
    blocks = [
        ("化学与生物工程学院", 50, [
            ("4301", 50, {"周一": 0}),
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_full_exempt.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))

    # 全院总应到 50 人，免勤 50 人（或 60 人）
    recs = [
        ExemptRecord(
            sid=f"990001000{i:04d}", name=f"免勤生{i}", college="化学与生物工程学院", class_text="化生1班", class_key="化生1班",
            weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山", room="", reason="校运会",
            group_size=None, source="未分类", row=3 + i,
        )
        for i in range(50)
    ]
    agg = aggregate_exemptions(recs, week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=True)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]
    # 分母 = 50 - 50 = 0 <= 0 -> 100% 满勤出勤！
    assert rate.rate == 1.0
    assert "全员免勤（满勤）" in rate.status
    assert "100.00% 满勤出勤计入" in rate.note


# ---------------------------------------------------------------------------
# 7. 全院所有教室实到均未录入（纯空白）时安全置空
# ---------------------------------------------------------------------------
def test_calc_all_blank_rooms_returns_none(tmp_path):
    blocks = [
        ("外国语学院", 150, [
            ("4401", 50, {"周一": None}),
            ("4402", 50, {"周一": None}),
            ("4403", 50, {"周一": None}),
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_all_blank.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions([], week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=True)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]
    assert rate.rate is None
    assert rate.status == "实到数据未录入"
