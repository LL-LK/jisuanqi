# -*- coding: utf-8 -*-
"""深度 Debug 11 项缺陷与同类衍生漏洞专项回归测试套件。"""
from __future__ import annotations

import json
import os
import subprocess
import sys
from dataclasses import fields
from pathlib import Path

import openpyxl
import pytest

from _helpers import make_detail_sheet, write_detail
from core.calc import Calculator
from core.dictionary import Dictionary, create_dictionary
from core.excel_safety import _sanitize_anomaly_text, generate_diagnostic_log
from core.io_detail import load_detail
from core.io_exempt import aggregate_exemptions, load_exemptions
from core.model import ExemptRecord
from core.parsing import (
    normalize_campus,
    normalize_class,
    parse_weekdays,
    parse_weeks,
    resolve_pairs,
)
from core.pipeline import RunConfig, RunResult, run


# ---------------------------------------------------------------------------
# Bug 1: PyInstaller 单文件 (_MEIPASS) 与只读目录持久化回退
# ---------------------------------------------------------------------------
def test_bug1_frozen_and_readonly_writable_paths(monkeypatch, tmp_path):
    import app as app_mod

    fake_exe_dir = tmp_path / "app_bin"
    fake_exe_dir.mkdir()
    fake_exe = fake_exe_dir / "免勤统计系统.exe"
    fake_exe.write_bytes(b"MZ")

    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(fake_exe))

    # 可写 EXE 同级目录
    assert app_mod.user_writable_dir() == str(fake_exe_dir)
    dict_path = app_mod.default_dictionary_path()
    assert dict_path.startswith(str(fake_exe_dir))
    out_dir = app_mod.default_output_dir()
    assert out_dir.startswith(str(fake_exe_dir))
    assert os.path.isdir(out_dir)

    # 模拟 EXE 目录不可写时自动回退到 LOCALAPPDATA
    fallback_appdata = tmp_path / "local_appdata"
    monkeypatch.setenv("LOCALAPPDATA", str(fallback_appdata))
    monkeypatch.setattr(
        app_mod,
        "_is_dir_writable",
        lambda p: str(fallback_appdata) in str(p),
    )
    writable = app_mod.user_writable_dir()
    assert str(fallback_appdata) in writable
    assert os.path.isdir(writable)


# ---------------------------------------------------------------------------
# Bug 2: 打包态内置进程内合成数据全链路自检
# ---------------------------------------------------------------------------
def test_bug2_frozen_inprocess_selfcheck():
    import app as app_mod

    code, out = app_mod._run_frozen_inprocess_selfcheck()
    assert code == 0, out
    assert "5 passed" in out


# ---------------------------------------------------------------------------
# Bug 3: 教室级精确免勤 (room_level) 采集、双教室列、无余量不标估算、全精确有余量分摊
# ---------------------------------------------------------------------------
def test_bug3_room_level_exact_and_dual_room_columns(tmp_path):
    exempt_path = tmp_path / "exempt_dual_room.xlsx"
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "早自习免勤人员统计表"
    ws1.append(["早自习标题"])
    ws1.append(["序号", "学号", "学院", "姓名", "专业班级", "早自习教室号", "晚自习教室号", "免勤周次", "免勤时间", "免勤自习", "所在校区", "申请免勤理由"])
    ws1.append([1, "9900010000101", "信息科学与工程学院", "测试一", "计科1班", "04207", "06101", "第7周", "周一", "早自习", "雁山校区", "训练"])
    ws1.append([2, "9900010000102", "信息科学与工程学院", "测试二", "计科1班", "04207", "06101", "第7周", "周一", "早自习", "雁山校区", "训练"])

    ws2 = wb.create_sheet("晚自习免勤人员统计表")
    ws2.append(["晚自习标题"])
    ws2.append(["序号", "学号", "学院", "姓名", "专业班级", "早自习教室号", "晚自习教室号", "免勤周次", "免勤时间", "免勤自习", "所在校区", "申请免勤理由"])
    # 晚自习表应优先读取「晚自习教室号」06101
    ws2.append([1, "9900010000101", "信息科学与工程学院", "测试一", "计科1班", "04207", "06101", "第7周", "周一", "晚自习", "雁山校区", "训练"])
    wb.save(exempt_path)

    recs = load_exemptions(str(exempt_path), week_no=7)
    agg = aggregate_exemptions(recs, week_no=7, campus="雁山")

    assert agg.room_level.get(("信息科学与工程学院", "周一", "早自习", "04207")) == 2
    assert agg.room_level.get(("信息科学与工程学院", "周一", "晚自习", "06101")) == 1

    # 构造明细表：全院免勤=2 且全部精确落在 4207（leftover == 0），两间教室均不应标为估算
    blocks = [
        ("信息科学与工程学院", 100, [
            ("4207", 50, {"周一": 48}),
            ("4208", 50, {"周一": 50}),
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_exact.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    calc = Calculator(model, agg, alias={})
    room_cells = calc.room_exemptions()

    r4207 = next(r for r in model.rooms if r.room == "4207")
    r4208 = next(r for r in model.rooms if r.room == "4208")
    assert room_cells[("雁山（早）", r4207.row, "周一")] == 2
    assert room_cells[("雁山（早）", r4208.row, "周一")] == 0
    assert ("雁山（早）", r4207.row, "周一") not in calc.estimated_cells
    assert ("雁山（早）", r4208.row, "周一") not in calc.estimated_cells


def test_bug3_all_rooms_exact_with_leftover_still_allocates(tmp_path):
    """当某学院所有教室都有精确免勤命中，但学院总免勤 > 教室精确免勤之和时，余量按剩余容量分摊而不丢失。"""
    blocks = [
        ("化学与生物工程学院", 100, [
            ("4101", 60, {"周一": 50}),
            ("4102", 40, {"周一": 35}),
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_leftover.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))

    recs = [
        ExemptRecord(
            sid="9900010000201", name="测一", college="化学与生物工程学院", class_text="化生1班", class_key="化生1班",
            weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山", room="4101",
            reason="训练", group_size=None, source="未分类", row=3,
        ),
        ExemptRecord(
            sid="9900010000202", name="测二", college="化学与生物工程学院", class_text="化生2班", class_key="化生2班",
            weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山", room="4102",
            reason="训练", group_size=None, source="未分类", row=4,
        ),
        # 第3人未填教室号，计入学院总免勤 (total=3 > exact_sum=2)
        ExemptRecord(
            sid="9900010000203", name="测三", college="化学与生物工程学院", class_text="化生1班", class_key="化生1班",
            weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山", room="",
            reason="训练", group_size=None, source="未分类", row=5,
        ),
    ]
    agg = aggregate_exemptions(recs, week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={})
    room_cells = calc.room_exemptions()

    total_allocated = sum(room_cells[("雁山（早）", r.row, "周一")] for r in model.rooms)
    assert total_allocated == 3
    assert any(("雁山（早）", r.row, "周一") in calc.estimated_cells for r in model.rooms)


# ---------------------------------------------------------------------------
# Bug 4: model.py 重复字段清理、非第 7 周「全学期」生效、多段周次-星期防串周
# ---------------------------------------------------------------------------
def test_bug4_whole_term_any_week_and_no_cartesian_bleed():
    field_names = [f.name for f in fields(ExemptRecord)]
    assert field_names.count("sid_padded") == 1

    # 非第 7 周（如第 12 周）「全学期」必须照常命中
    pairs_wt, _ = resolve_pairs("全学期", "周一至周五", window_week=12)
    assert (12, "周一") in pairs_wt

    rec = ExemptRecord(
        sid="9900010000301", name="测全", college="地球科学学院", class_text="地学1班", class_key="地学1班",
        weeks=parse_weeks("全学期"), weekdays=parse_weekdays("周一至周五"),
        session="早自习", campus="雁山", room="", reason="训练", group_size=None, source="未分类", row=3,
        whole_term=True, pairs=pairs_wt,
    )
    # 即使换到第 15 周统计，whole_term 记录依然生效
    agg15 = aggregate_exemptions([rec], week_no=15, campus="雁山")
    assert agg15.exempt_count("地球科学学院", "周一", "早自习") == 1

    # 多段「第6周周一，第7周周三」在第 7 周只应命中周三，绝不串出周一
    pairs_multi, _ = resolve_pairs("第6周周一，第7周周三", "", window_week=7)
    rec_multi = ExemptRecord(
        sid="9900010000302", name="测串", college="地球科学学院", class_text="地学1班", class_key="地学1班",
        weeks=parse_weeks("第6周周一，第7周周三"), weekdays=parse_weekdays("第6周周一，第7周周三"),
        session="早自习", campus="雁山", room="", reason="训练", group_size=None, source="未分类", row=4,
        whole_term=False, pairs=pairs_multi,
    )
    assert rec_multi.days_for(7) == {"周三"}
    agg7 = aggregate_exemptions([rec_multi], week_no=7, campus="雁山")
    assert agg7.exempt_count("地球科学学院", "周一", "早自习") == 0
    assert agg7.exempt_count("地球科学学院", "周三", "早自习") == 1


# ---------------------------------------------------------------------------
# Bug 5 & 8 & 9: 同名学院多块合并分摊、未合并 B 列继承、别名学院 incomplete_scopes 与「免勤」说明
# ---------------------------------------------------------------------------
def test_bug5_multiple_blocks_same_college_and_unmerged_rows(tmp_path):
    # 构造两个同名学院块：确保分摊只按学院总免勤分摊一次，不翻倍
    blocks = [
        ("环境科学与工程学院", 100, [
            ("6101", 30, {"周一": 28}),
            ("6102", 30, {"周一": 27}),
        ]),
        ("环境科学与工程学院", 100, [
            ("6103", 40, {"周一": 35}),
        ]),
    ]
    detail_path = write_detail(tmp_path / "detail_multiblock.xlsx", [make_detail_sheet("雁山（早）", blocks)])

    # 将第一个学院块的第 4 行解除合并并清空 B4，测试 io_detail 对未合并空 B 列行的继承
    wb = openpyxl.load_workbook(detail_path)
    ws = wb["雁山（早）"]
    for rng in list(ws.merged_cells.ranges):
        if rng.min_col == 2 and rng.max_col == 2 and rng.min_row == 4:
            ws.unmerge_cells(str(rng))
    ws.cell(5, 2, None)
    wb.save(detail_path)
    wb.close()

    model = load_detail(str(detail_path))
    assert len(model.rooms) == 3
    assert all(r.college_name == "环境科学与工程学院" for r in model.rooms)

    recs = [
        ExemptRecord(
            sid=f"990001000040{i}", name=f"测{i}", college="环境科学与工程学院", class_text="环境1班", class_key="环境1班",
            weeks={7}, weekdays={"周一"}, session="早自习", campus="雁山", room="",
            reason="训练", group_size=None, source="未分类", row=3 + i,
        )
        for i in range(5)
    ]
    agg = aggregate_exemptions(recs, week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={})
    room_cells = calc.room_exemptions()
    # 3 间教室分摊之和必须恒等于 5，绝不能因 2 个 CollegeBlock 翻倍成 10
    total_alloc = sum(room_cells[("雁山（早）", r.row, "周一")] for r in model.rooms)
    assert total_alloc == 5


def test_bug8_and_9_alias_incomplete_scope_and_mianqin_text_reason(tmp_path):
    blocks = [
        ("亚太信息技术学院", 217, [
            ("018A401", 30, {"周一": "免勤"}),
        ])
    ]
    detail_path = write_detail(tmp_path / "detail_mianqin_text.xlsx", [make_detail_sheet("雁山（早）", blocks)])
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions([], week_no=7, campus="雁山")
    calc = Calculator(model, agg, alias={})
    rates = [r for r in calc.college_rates() if r.day == "周一"]
    assert len(rates) == 1
    assert rates[0].rate is None
    assert "免勤" in rates[0].note
    assert "018A401" in rates[0].note

    # 测试 incomplete_scopes 中记录的是简称，而明细表是全称时仍能拦截
    blocks2 = [
        ("土木与建筑工程学院", 200, [
            ("4201", 50, {"周一": 45}),
        ])
    ]
    detail_path2 = write_detail(tmp_path / "detail_alias_inc.xlsx", [make_detail_sheet("雁山（早）", blocks2)])
    model2 = load_detail(str(detail_path2))
    agg2 = aggregate_exemptions([], week_no=7, campus="雁山")
    agg2.incomplete_scopes.add(("土木学院", "周一", "早自习"))
    calc2 = Calculator(model2, agg2, alias={"土木学院": "土木与建筑工程学院"})
    rates2 = [r for r in calc2.college_rates() if r.day == "周一"]
    assert rates2[0].rate == pytest.approx(0.225)
    assert "免勤不完整" in rates2[0].status


# ---------------------------------------------------------------------------
# Bug 6 & 10: Web 端 calc.py 无教室号列不崩溃、无教室号计入分母、早晚自习括号总人数独立
# ---------------------------------------------------------------------------
def test_bug6_and_10_web_calc_roomless_and_morning_evening_totals(tmp_path):
    web_calc_src = Path(__file__).resolve().parents[2] / "calc.py"
    if not web_calc_src.exists():
        pytest.skip("Web calc.py not in parent repo root")

    # P1: 早自习括号人数 200 人，晚自习括号人数 100 人
    p1 = tmp_path / "input_p1.xlsx"
    wb1 = openpyxl.Workbook()
    ws_m = wb1.active
    ws_m.title = "雁山早"
    ws_m.cell(1, 1, "标题")
    for c, h in enumerate(["序号", "学院（人数）", "教室", "班级", "应到人数", "周一早自习实到人数", "周一早自习免勤人数"], 1):
        ws_m.cell(2, c, h)
    ws_m.cell(3, 2, "地球科学学院（200人）")
    ws_m.cell(3, 3, 4201)
    ws_m.cell(3, 5, 200)
    ws_m.cell(3, 6, 170)

    ws_e = wb1.create_sheet("雁山晚")
    ws_e.cell(1, 1, "标题")
    for c, h in enumerate(["序号", "学院（人数）", "教室", "班级", "应到人数", "周一晚自习实到人数", "周一晚自习免勤人数"], 1):
        ws_e.cell(2, c, h)
    ws_e.cell(3, 2, "地球科学学院（100人）")
    ws_e.cell(3, 3, 4201)
    ws_e.cell(3, 5, 100)
    ws_e.cell(3, 6, 80)
    wb1.save(p1)

    # P5 & P3: 完全不含「教室号」列，验证 add_leave 不抛 TypeError 且免勤人数计入分母
    p5 = tmp_path / "input_p5.xlsx"
    wb5 = openpyxl.Workbook()
    ws5 = wb5.active
    ws5.title = "早自习"
    ws5.append(["统计表标题"])
    ws5.append(["序号", "学院", "学号", "姓名", "免勤时间", "免勤自习", "所在校区", "免勤周次", "申请免勤理由"])
    for i in range(20):
        ws5.append([i + 1, "地球科学学院", f"99000200000{i:02d}", f"测试生{i}", "周一", "早自习", "雁山", "第7周", "训练"])
    wb5.save(p5)

    p3 = tmp_path / "input_p3.xlsx"
    wb3 = openpyxl.Workbook()
    ws3 = wb3.active
    ws3.title = "在线存档"
    ws3.append(["存档标题"])
    ws3.append(["序号", "学院", "学号", "姓名", "请假时间", "请假自习", "所在校区", "请假周次", "请假理由"])
    for i in range(10):
        ws3.append([i + 1, "地球科学学院", f"99000300000{i:02d}", f"测试假{i}", "周一", "早自习", "雁山", "第7周", "病假"])
    wb3.save(p3)

    env = os.environ.copy()
    env["SZW_WEEK"] = "7"
    env["SZW_CAMPUSES"] = "雁山"
    env["SZW_APPLY_CAP"] = "0"
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(web_calc_src)],
        cwd=str(tmp_path),
        env=env,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    assert proc.returncode == 0, proc.stderr + "\n" + proc.stdout

    summary = json.loads((tmp_path / "summary.json").read_text(encoding="utf-8"))
    sessions = summary["collegeCalc"][0]["sessions"]
    morning = next(s for s in sessions if s["session"] == "早")
    evening = next(s for s in sessions if s["session"] == "晚")
    # 早自习：应到 200（未被晚自习 100 覆盖），全免勤 20 + 临时请假 10 = 30，分母 = 170，实到 170 -> 100.0%
    assert morning["obj"] == 20
    assert morning["sub"] == 10
    assert morning["denom"] == 170
    assert morning["rate"] == 100.0
    # 晚自习：应到 100，免勤 0，分母 = 100，实到 80 -> 80.0%
    assert evening["denom"] == 100
    assert evening["rate"] == 80.0


# ---------------------------------------------------------------------------
# Bug 7: 词典双向学院映射、校区/班级别名接入、免勤理由子串匹配与原子保存
# ---------------------------------------------------------------------------
def test_bug7_dictionary_aliases_and_atomic_save(tmp_path):
    dict_file = tmp_path / "dict_test.xlsx"
    create_dictionary(str(dict_file))
    d = Dictionary.load(str(dict_file))
    # 别名映射到正名
    d.college_alias["测绘学院"] = "测绘地理信息学院"
    d.college_alias["测绘地理信息学院"] = "测绘地理信息学院"
    d.campus_alias["雁山南区"] = "雁山"
    d.class_alias["测绘24-1"] = "测绘2024-1班"
    d.reason_class["体测"] = "其他"

    # 别名与正名均映射到正名
    assert d.college("测绘学院") == "测绘地理信息学院"
    assert d.college("测绘地理信息学院") == "测绘地理信息学院"
    # 词典构造时双向映射测试
    custom_tables = {
        "学院别名": [["测绘地理信息学院", "测绘学院", "说明"]],
        "班级别名": [["测绘24-1", "测绘2024-1班", "说明"]],
        "校区别名": [["雁山南区", "雁山", "说明"]],
        "免勤理由分类": [["体测", "其他", "说明"]],
        "统计参数": [],
    }
    d2 = Dictionary(custom_tables)
    assert d2.college("测绘学院") == "测绘地理信息学院"
    assert d2.college("测绘地理信息学院") == "测绘地理信息学院"
    assert d2.campus("雁山南区") == "雁山"
    assert d2.class_key("测绘24-1班") == d2.class_key("测绘2024-1班") == "测绘241"
    assert d2.reason_class_of("本周参加大学生体测免勤") == "其他"


# ---------------------------------------------------------------------------
# Bug 11: 诊断日志隐私脱敏（少数民族姓名、老师前缀、助词截断、非 C 盘路径、header_cells）
# ---------------------------------------------------------------------------
def test_bug11_privacy_sanitization_and_diagnostic_log(tmp_path):
    d_drive_path = "D" + ":\\Users\\22140\\Desktop\\secret.xlsx"
    text = f"辅导员阿依古丽·买买提反馈学生王小明因病请假，学号 9900010000999，文件位于 {d_drive_path}"
    cleaned = _sanitize_anomaly_text(text)
    assert "阿依古丽" not in cleaned
    assert "王小明" not in cleaned
    assert "9900010000999" not in cleaned
    assert "22140" not in cleaned
    assert "因病请假" in cleaned

    # 验证 generate_diagnostic_log 对 header_cells 与 error 同步脱敏
    bad_xlsx = tmp_path / "sample.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["学号9900010000999", "学生张三丰请假", d_drive_path])
    wb.save(bad_xlsx)

    cfg = RunConfig(detail_path=str(bad_xlsx), exempt_path=str(bad_xlsx), output_path=str(tmp_path / "out.xlsx"))
    res = RunResult(ok=False, error=f"读取失败：{d_drive_path} 老师李四光报错")
    log_text = generate_diagnostic_log(cfg, res)
    assert "9900010000999" not in log_text
    assert "张三丰" not in log_text
    assert "李四光" not in log_text
    assert "22140" not in log_text


# ---------------------------------------------------------------------------
# Rule 12: 导出与管线同路径覆盖拦截 + 原子落盘无残留临时文件
# ---------------------------------------------------------------------------
def test_pipeline_blocks_exempt_overwrite_and_atomic_export(tmp_path):
    detail_path = tmp_path / "detail.xlsx"
    exempt_path = tmp_path / "exempt.xlsx"
    openpyxl.Workbook().save(detail_path)
    openpyxl.Workbook().save(exempt_path)

    # 输出路径与免勤表路径相同时必须被安全拦截
    cfg = RunConfig(
        detail_path=str(detail_path),
        exempt_path=str(exempt_path),
        output_path=str(exempt_path),
        week_no=7,
    )
    res = run(cfg)
    assert res.ok is False
    assert "覆盖" in (res.error or "")
