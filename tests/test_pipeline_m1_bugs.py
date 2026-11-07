# -*- coding: utf-8 -*-
"""Milestone 1 Bug 1 & Bug 2 专属回归与边界测试套件。

覆盖缺陷:
  * Bug 1: core/pipeline.py:126 单免勤表导入时身份核验、学院归一与「免勤核对」工作表旁路缺陷
  * Bug 2: core/calc.py:356-370 complete_mode 容错模式下有效教室考勤分母失真缺陷
"""

from __future__ import annotations

from pathlib import Path
import openpyxl
import pytest

from _helpers import (
    build_focused_sample,
    make_detail_sheet,
    make_exempt_row,
    write_detail,
    write_exempt,
)
from core.calc import Calculator
from core.io_detail import load_detail
from core.io_exempt import aggregate_exemptions, ExemptRecord
from core.pipeline import RunConfig, run


# ==============================================================================
# Bug 1 回归测试: 单免勤表身份核验与「免勤核对」工作表
# ==============================================================================

def test_single_exempt_generates_identity_review_sheet(tmp_path):
    """验证导入单份免勤表且存在异常时，成功生成「免勤核对」工作表并准确记录来源引用。"""
    detail_path = tmp_path / "detail_single.xlsx"
    exempt_path = tmp_path / "exempt_single.xlsx"
    out_path = tmp_path / "result_single.xlsx"

    # 1. 构造明细表
    early_sheet = make_detail_sheet("雁山（早）", [
        ("化学与生物工程学院", 100, [("4101", 100, {"周一": 90})]),
    ])
    write_detail(detail_path, [early_sheet])

    # 2. 构造单份免勤表: 注入 2 条必须触发审核的异常数据
    # A. 学号非 13 位 (异常学号)
    row_odd_sid = make_exempt_row(
        sid="990001001", college="化生院", name="测试生01",
        class_text="化生26-1", week_text="第7周", day_text="周一", session="早"
    )
    # B. 占位学号无人数 (待补人数)
    row_placeholder = make_exempt_row(
        sid="全班", college="化学与生物工程学院", name="全班免勤",
        class_text="化生26-2", week_text="第7周", day_text="周一", session="早"
    )
    write_exempt(exempt_path, {"早自习免勤人员统计表": [row_odd_sid, row_placeholder]})

    # 3. 运行流水线 (导入单份免勤表)
    config = RunConfig(
        detail_path=str(detail_path),
        exempt_path=str(exempt_path),
        output_path=str(out_path),
        week_no=7,
        campus="雁山",
        include_exempt_summary=True,  # 要求生成免勤报告与核对表
    )
    result = run(config)
    assert result.ok is True, result.error

    # 4. 校验 Excel 产物
    wb = openpyxl.load_workbook(str(out_path), data_only=True)
    try:
        assert "免勤核对" in wb.sheetnames, "单份免勤表导入且有异常时，必须生成「免勤核对」工作表！"
        ws_review = wb["免勤核对"]
        headers = [ws_review.cell(1, c).value for c in range(1, 4)]
        assert headers == ["问题", "来源", "处理说明"]

        rows_data = [
            [ws_review.cell(r, c).value for c in range(1, 4)]
            for r in range(2, ws_review.max_row + 1)
        ]
        types = [r[0] for r in rows_data]
        assert any("学号" in str(t) for t in types)
        assert any("待补" in str(t) for t in types)
    finally:
        wb.close()


def test_single_exempt_identity_check_and_audit(tmp_path):
    """验证单免勤表导入时，学院别名归一、冲突审计与免勤核对工作表正常生效。"""
    blocks = [
        ("化学与生物工程学院", 100, [("4101", 100, {"周一": 90})]),
        ("材料科学与工程学院", 100, [("4201", 100, {"周一": 85})]),
    ]
    detail = write_detail(tmp_path / "detail_identity.xlsx", [make_detail_sheet("雁山（早）", blocks)])

    exempt_rows = [
        make_exempt_row(sid="9900001000001", college="化生院", name="张三", day_text="周一", session="早"),
        make_exempt_row(sid="9900001000001", college="材工学院", name="张三", day_text="周一", session="早"),
    ]
    exempt = write_exempt(tmp_path / "exempt_identity.xlsx", {"早自习免勤人员统计表": exempt_rows})
    out = tmp_path / "out_identity.xlsx"

    cfg = RunConfig(
        detail_path=str(detail),
        exempt_path=str(exempt),
        output_path=str(out),
        week_no=7,
        include_exempt_summary=True,
    )
    res = run(cfg)
    assert res.ok, res.error

    # 1. 验证冲突被捕获
    assert any("学号对应学院或姓名冲突" in a.get("type", "") for a in res.anomalies)

    # 2. 验证生成了免勤核对工作表
    wb = openpyxl.load_workbook(str(out), data_only=True)
    try:
        assert "免勤核对" in wb.sheetnames
        ws_review = wb["免勤核对"]
        assert ws_review.max_row >= 2
    finally:
        wb.close()


def test_single_exempt_clean_data_no_spurious_review_sheet(tmp_path):
    """验证数据完全干净的单份免勤表不会生成多余的空「免勤核对」工作表。"""
    detail_path, exempt_path = build_focused_sample(tmp_path)
    out_path = tmp_path / "result_clean.xlsx"

    config = RunConfig(
        detail_path=str(detail_path),
        exempt_path=str(exempt_path),
        output_path=str(out_path),
        week_no=7,
        include_exempt_summary=True,
    )
    result = run(config)
    assert result.ok is True

    wb = openpyxl.load_workbook(str(out_path), data_only=True)
    try:
        # 无 review_issues 时不生成空表
        assert "免勤核对" not in wb.sheetnames
    finally:
        wb.close()


# ==============================================================================
# Bug 2 回归测试: 容错模式下有效教室分母正确性
# ==============================================================================

def test_complete_calc_denominator_excludes_missing_rooms(tmp_path):
    """
    场景: 地球科学学院共 3 间教室，每间应到 100 人，总应到 300 人。
    周一实到:
      - 4201: 95 人 (有效)
      - 4202: 85 人 (有效)
      - 4203: None (未录入)
    免勤: 0 人。

    旧算法 (Bug 2):
      分子 = 95 + 85 = 180
      分母 = 300 (全院应到)
      率 = 180 / 300 = 60.00% (严重失真，将 4203 视为全员缺勤)

    修复后新算法:
      分子 = 95 + 85 = 180
      有效教室应到 = 100 + 100 = 200
      有效分母 = 200 - 0 = 200
      率 = 180 / 200 = 90.00% (真实反映有效教室出勤水平)
    """
    blocks = [
        ("地球科学学院", 300, [
            ("4201", 100, {"周一": 95}),
            ("4202", 100, {"周一": 85}),
            ("4203", 100, {"周一": None}),  # 留空
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_partial_fix.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions([], week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=True)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]

    # 核心数学断言: 出勤率必须是 90.00% 而不是 60.00%
    assert rate.rate == pytest.approx(0.90, rel=1e-4)
    assert rate.actual == 180
    assert "部分教室缺失（已按有效实到计算）" in rate.status
    assert "4203" in rate.note
    assert any("未录入" in a for a in rate.anomalies)


def test_complete_calc_with_valid_room_exemptions(tmp_path):
    """验证部分教室缺失时，有效教室分母正确扣减有效教室对应免勤，防止出勤率畸变。"""
    blocks = [
        ("地球科学学院", 300, [
            ("4201", 100, {"周一": 85}),
            ("4202", 100, {"周一": 80}),
            ("4203", 100, {"周一": None}),  # 缺失未填
        ])
    ]
    detail = write_detail(tmp_path / "detail_exempt_calc.xlsx", [make_detail_sheet("雁山（早）", blocks)])

    recs = [
        ExemptRecord(
            sid=f"990001000{i:04d}", name=f"免勤生{i}", college="地球科学学院", class_text="地科1班", class_key="地科1班",
            weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山",
            room="4201" if i < 10 else ("4202" if i < 15 else "4203"),
            reason="公假", group_size=None, source="未分类", row=3 + i,
        )
        for i in range(30)
    ]
    model = load_detail(str(detail))
    agg = aggregate_exemptions(recs, week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=True)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]

    # 真实基数：
    # 有效教室应到 = 100 + 100 = 200
    # 有效教室免勤 = 10 + 5 = 15（4203 的 15 人绝不扣减有效教室分母）
    # 有效考勤分母 = 200 - 15 = 185
    # 有效实到合计 = 85 + 80 = 165
    # 正确出勤率 = 165 / 185 ≈ 89.189%
    expected_rate = 165.0 / 185.0
    assert rate.rate == pytest.approx(expected_rate, rel=1e-4)
    assert "部分教室缺失（已按有效实到计算）" in rate.status
    assert "4203" in rate.note


def test_strict_mode_partial_missing_remains_none(tmp_path):
    """对比验证: 严格契约模式 (complete_mode=False) 下，部分教室缺失仍严格保持为 None。"""
    blocks = [
        ("地球科学学院", 300, [
            ("4201", 100, {"周一": 95}),
            ("4202", 100, {"周一": 85}),
            ("4203", 100, {"周一": None}),
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_partial_strict.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions([], week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={}, complete_mode=False)

    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    rate = rates[0]
    assert rate.rate is None
    assert rate.status == "实到数据未完整填写"
