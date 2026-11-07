# -*- coding: utf-8 -*-
"""班级与学号多维穿透、班级精准落位及双模出勤率计算专项自动化测试。"""

import pytest
from core.parsing import parse_class_feature, match_class_to_room_classes
from core.io_exempt import Aggregate, ExemptRecord, aggregate_exemptions
from core.model import DetailModel, RoomRow, CollegeBlock
from core.calc import Calculator


def test_parse_class_feature_variants():
    # 标准与简称对照
    assert parse_class_feature("地球化学2026-1班") == ("地球化学", "26-1")
    assert parse_class_feature("地球化学26-1") == ("地球化学", "26-1")
    assert parse_class_feature("地球化学2026拔尖班") == ("地球化学", "26-拔尖")
    assert parse_class_feature("资勘26_2") == ("资源勘查工程", "26-2")
    assert parse_class_feature("资源勘查工程2026-2班") == ("资源勘查工程", "26-2")
    assert parse_class_feature("地球信息科学与技术2026-1半") == ("地球信息科学与技术", "26-1")
    assert parse_class_feature("地求信息2026-2班") == ("地球信息科学与技术", "26-2")
    assert parse_class_feature("勘技2026创新班") == ("勘查技术与工程", "26-创新")
    assert parse_class_feature("环工26-1班") == ("环境工程", "26-1")
    assert parse_class_feature("给排水26-创新班") == ("给排水科学与工程", "26-创新")


def test_match_class_to_room_classes():
    room_classes = "地球信息2026-1班、勘技2026-1班、勘技2026-2班、资勘2026-1班、资勘2026-2班"
    
    # 精准命中
    assert match_class_to_room_classes("地球信息科学与技术2026-1半", room_classes) is True
    assert match_class_to_room_classes("勘查技术与工程26-1班", room_classes) is True
    assert match_class_to_room_classes("资源勘查工程2026-2班", room_classes) is True
    assert match_class_to_room_classes("资勘26_1", room_classes) is True

    # 不在教室内的班级
    assert match_class_to_room_classes("地质工程2026-1班", room_classes) is False
    assert match_class_to_room_classes("环境工程26-1班", room_classes) is False


def test_class_exact_room_allocation_and_conservation():
    """验证：学生班级命中教室优先定点扣减，且数学守恒。"""
    # 模拟明细表：某学院两间教室
    r1 = RoomRow(
        sheet="雁山（早）", row=3, seq=1, college_name="地科学院", college_total=100,
        room="6508", classes_text="地信26-1班、地信26-2班", expected=50,
        actuals={"周一": 45}, actual_raw={"周一": "45"}, exempt_cols={"周一": 7}
    )
    r2 = RoomRow(
        sheet="雁山（早）", row=4, seq=2, college_name="地科学院", college_total=100,
        room="6509", classes_text="地工26-1班", expected=50,
        actuals={"周一": 40}, actual_raw={"周一": "40"}, exempt_cols={"周一": 7}
    )
    blk = CollegeBlock(sheet="雁山（早）", name="地科学院", total=100, rows=[3, 4], data_rows=[3, 4])
    detail = DetailModel(
        path="mock.xlsx", sheet_names=["雁山（早）"], rooms=[r1, r2], blocks=[blk],
        rate_cols_by_sheet={"雁山（早）": {"周一": 8, "平均": 9}}
    )

    # 模拟免勤记录：
    # 2名学生来自地信26-1（应精准落入 6508）
    # 1名学生来自资勘26-1（未被抽查班级，属于余量，按容量分摊）
    rec1 = ExemptRecord(
        sid="9900010800001", name="张三", college="地科学院", class_text="地球信息科学与技术2026-1",
        class_key="地信261", weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山",
        room="", reason="选修课", group_size=None, source="全免勤大表", row=3
    )
    rec2 = ExemptRecord(
        sid="9900010800002", name="李四", college="地科学院", class_text="地球信息26-1班",
        class_key="地信261", weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山",
        room="", reason="社团", group_size=None, source="全免勤大表", row=4
    )
    rec3 = ExemptRecord(
        sid="9900010800003", name="王五", college="地科学院", class_text="资源勘查工程26-1班",
        class_key="资勘261", weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山",
        room="", reason="公假", group_size=None, source="全免勤大表", row=5
    )

    agg = aggregate_exemptions([rec1, rec2, rec3], week_no=7, campus="雁山")
    assert agg.exempt_count("地科学院", "周一", "早自习") == 3
    assert len(agg.get_students("地科学院", "周一", "早自习")) == 3

    # 执行计算
    calc = Calculator(detail, agg, complete_mode=False)
    room_ex = calc.room_exemptions()

    # 6508 应该有 2 人来自班级精准命中，另外 1 人来自余量分摊
    # 总免勤 3 人，两教室应到 50 vs 50
    # 6508 (原容量 50-2=48), 6509 (原容量 50-0=50) -> 余量 1 人落在 6509
    assert room_ex[("雁山（早）", 3, "周一")] == 2  # 6508 精准落位 2 人
    assert room_ex[("雁山（早）", 4, "周一")] == 1  # 6509 余量分摊 1 人
    # 守恒性检验
    assert sum(room_ex.values()) == 3

    # 穿透映射检查
    stus_6508 = calc.room_students_map[("雁山（早）", 3, "周一")]
    assert len(stus_6508) == 2
    assert {s["name"] for s in stus_6508} == {"张三", "李四"}


def test_dual_mode_rates():
    """验证严格契约模式 vs 容错模式的出勤率计算。"""
    # 某一教室实到缺失（文字“免勤”）
    r1 = RoomRow(
        sheet="雁山（早）", row=3, seq=1, college_name="测试学院", college_total=100,
        room="101", classes_text="测试班", expected=50,
        actuals={"周一": None}, actual_raw={"周一": "免勤"}, exempt_cols={"周一": 7}
    )
    blk = CollegeBlock(sheet="雁山（早）", name="测试学院", total=100, rows=[3], data_rows=[3])
    detail = DetailModel(
        path="mock.xlsx", sheet_names=["雁山（早）"], rooms=[r1], blocks=[blk],
        rate_cols_by_sheet={"雁山（早）": {"周一": 8, "平均": 9}}
    )
    agg = Aggregate()

    # 1. 严格模式：实到缺失留空，rate 为 None
    calc_strict = Calculator(detail, agg, complete_mode=False)
    rates_strict = calc_strict.college_rates()
    assert rates_strict[0].rate is None
    assert rates_strict[0].status == "实到数据未完整填写"

    # 2. 容错模式：文字“免勤”折为0，生成出勤率
    calc_comp = Calculator(detail, agg, complete_mode=True)
    rates_comp = calc_comp.college_rates()
    assert rates_comp[0].rate is not None
    assert rates_comp[0].rate == 0.0
