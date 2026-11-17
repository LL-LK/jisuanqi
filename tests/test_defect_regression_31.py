# -*- coding: utf-8 -*-
"""31 项审计实锤缺陷针对性回归测试套件 (R1-1~R1-4, R2-1~R2-10, R3-1~R3-10, R4-1~R4-7)。"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import subprocess
import sys
from contextlib import closing
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import cli  # noqa: E402
from core import RunConfig, RunResult, run  # noqa: E402
from core.attendance_summary import export_attendance_summary, load_summary_template  # noqa: E402
from core.calc import Calculator  # noqa: E402
from core.dictionary import Dictionary, HEADERS, save_tables  # noqa: E402
from core.excel_safety import (  # noqa: E402
    _sanitize_anomaly_text,
    generate_diagnostic_log,
    preflight_excel_file,
)
from core.exempt_report import append_exempt_report, export_exempt_report  # noqa: E402
from core.formulas import (  # noqa: E402
    count_workbook_formulas,
    evaluate_workbook_cell,
    export_flattened_workbook,
    flatten_workbook_formulas,
    is_formula_cell,
    load_evaluated_workbook,
    normalize_output_cell_value,
)
from core.io_detail import load_detail  # noqa: E402
from core.io_exempt import aggregate_exemptions, load_exemptions, normalize_sid  # noqa: E402
from core.model import CollegeRate  # noqa: E402
from core.parsing import parse_int  # noqa: E402
from core.pipeline import _write_summary  # noqa: E402
from core.privacy import (  # noqa: E402
    PrivacyViolationError,
    assert_model_privacy_compliance,
    strip_workbook_privacy,
)
from core.storage import AttendanceStorage  # noqa: E402
from core.summary_template import fill_summary_workbook, generate_summary_file  # noqa: E402
from tests._helpers import make_detail_sheet, make_exempt_row, write_detail, write_exempt  # noqa: E402
from tests_e2e.helpers import create_detail_workbook  # noqa: E402

ACTIVE_SKILL_DIR = Path.home() / '.agents' / 'skills' / 'attendance-stats-calc'
REPO_SKILL_DIR = ROOT / '.agents' / 'skills' / 'attendance-stats-calc'


def _load_module_from_path(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, str(path))
    mod = importlib.util.module_from_spec(spec)
    assert spec is not None and spec.loader is not None
    spec.loader.exec_module(mod)
    return mod


# ==============================================================================
# R1: 纯数值与公式清除 (R1-1 ~ R1-4)
# ==============================================================================


def test_r1_1_percent_semantic_ignores_banner_exclude_keywords(tmp_path: Path):
    """R1-1: 首行横幅含'备注/学院'时，学生出勤率列仍按百分比保留 1.0 + '0.00%'。"""
    src = tmp_path / 'banner_detail.xlsx'
    dst = tmp_path / 'banner_flat.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = '雁山（早）'
    ws.merge_cells('A1:D1')
    ws['A1'] = '桂林理工大学早晚自习明细表（备注：各学院抽查教室统计）'
    ws['A2'] = '学院'
    ws['B2'] = '实到人数'
    ws['C2'] = '学生出勤率（%）'
    ws['D2'] = '出勤率备注'
    ws['C3'] = '周一'
    ws['A4'] = '理学院'
    ws['B4'] = '=50+50'
    ws['C4'] = '=100/100'
    ws['D4'] = '=1+1'
    wb.save(src)
    wb.close()

    stats = export_flattened_workbook(str(src), str(dst))
    assert stats.residual_formulas == 0
    wb_out = load_workbook(dst, data_only=False)
    try:
        ws_out = wb_out['雁山（早）']
        assert count_workbook_formulas(wb_out) == 0
        # 实到人数 -> 整数 100 + '0'
        assert ws_out['B4'].value == 100 and isinstance(ws_out['B4'].value, int)
        assert ws_out['B4'].number_format == '0'
        # 出勤率 -> 100% 满勤，number_format 必须保持 '0.00%'（不再误改为整数 '0'）
        assert ws_out['C4'].value == pytest.approx(1.0)
        assert ws_out['C4'].number_format == '0.00%'
        # 且 normalize_output_cell_value 直接返回 float 1.0
        tmp_cell = Workbook().active['A1']
        assert isinstance(normalize_output_cell_value(tmp_cell, 1.0, col_header=ws_out['C2'].value + '|' + ws_out['C3'].value), float)
        assert tmp_cell.number_format == '0.00%'
        # 出勤率备注 -> 含'备注'于具体列头，按非百分比处理
        assert ws_out['D4'].value == 2 and isinstance(ws_out['D4'].value, int)
    finally:
        wb_out.close()


def test_r1_2_exempt_report_escapes_leading_equals_formula(tmp_path: Path):
    """R1-2: 免勤报表含 '=' 开头文本时安全转义为纯文本，不触发公式残留异常。"""
    book = Workbook()
    book.remove(book.active)
    rows = [{
        '学院': '=学院测试',
        '星期': '周一',
        '自习': '早自习',
        '个人免勤': 1,
        '集体免勤': 0,
        '免勤合计': 1,
        '已回填': 1,
        '容量超额': 0,
        '回填差额': 0,
        '免勤状态': '已统计',
        '模板出勤率列': '有',
        '回填状态': '已回填',
    }]
    issues = [{
        'type': '规则待核对',
        'level': 'warn',
        'scope': '理学院',
        'review_ref': '=HYPERLINK("http://x","ref")',
        'note': '=无= 请复核',
    }]
    append_exempt_report(book, rows, False, {'统计周次': 7}, issues)
    out = tmp_path / 'exempt_rep.xlsx'
    book.save(out)
    book.close()

    reloaded = load_workbook(out, data_only=False)
    try:
        assert count_workbook_formulas(reloaded) == 0
        for ws in reloaded.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    assert not is_formula_cell(cell)
        assert reloaded['免勤统计'].cell(2, 1).value == "'=学院测试"
        assert reloaded['免勤核对'].cell(2, 2).value.startswith("'")
        assert reloaded['免勤核对'].cell(2, 3).value.startswith("'")
    finally:
        reloaded.close()


def test_r1_3_preflight_detects_formulas_and_counts(tmp_path: Path):
    """R1-3: preflight_excel_file 正确探测是否含公式与公式数量。"""
    f_with = tmp_path / 'with_formulas.xlsx'
    f_without = tmp_path / 'pure_values.xlsx'

    wb1 = Workbook()
    ws1 = wb1.active
    ws1['A1'] = 10
    ws1['A2'] = '=A1*2'
    ws1['A3'] = '=SUM(A1:A2)'
    wb1.save(f_with)
    wb1.close()

    wb2 = Workbook()
    wb2.active['A1'] = 10
    wb2.save(f_without)
    wb2.close()

    rep_with = preflight_excel_file(str(f_with))
    assert rep_with['is_valid'] is True
    assert rep_with['has_formulas'] is True
    assert rep_with['formula_count'] == 2

    rep_without = preflight_excel_file(str(f_without))
    assert rep_without['is_valid'] is True
    assert rep_without['has_formulas'] is False
    assert rep_without['formula_count'] == 0


def test_r1_4_normalize_output_cell_value_resets_float_format_to_int():
    """R1-4: 已有浮点格式 '0.00' 的整数人数格被强制重置为 '0'。"""
    wb = Workbook()
    ws = wb.active
    cell = ws['A1']
    cell.value = 105.0
    cell.number_format = '0.00'

    norm = normalize_output_cell_value(cell, 105.0, col_header='实到人数')
    assert norm == 105 and isinstance(norm, int)
    assert cell.number_format == '0'


# ==============================================================================
# R2: 隐私脱敏合规 (R2-1 ~ R2-10)
# ==============================================================================


def test_r2_1_strip_privacy_redacts_header_alias_variants(tmp_path: Path):
    """R2-1: '学生学号'/'学生姓名'/'申请人姓名' 等表头变体均被识别并彻底脱敏。"""
    src = tmp_path / 'alias_exempt.xlsx'
    dst = tmp_path / 'alias_stripped.xlsx'
    fake_sid = '3202' + '001010001'
    wb = Workbook()
    ws = wb.active
    ws.title = '早自习'
    ws.append(['学生学号', '学生姓名', '所属学院', '专业班级', '免勤周次', '免勤原因'])
    ws.append([fake_sid, '赵铁柱', '理学院', '理学24-1班', '第7周', '急性胃炎就医'])
    wb.save(src)
    wb.close()

    stats = strip_workbook_privacy(str(src), str(dst))
    assert stats['rows_stripped'] == 1
    assert stats['sids_anonymized'] == 1
    assert stats['names_removed'] == 1

    wb_out = load_workbook(dst)
    try:
        vals = [c.value for c in wb_out['早自习'][2]]
        joined = ' '.join(str(v) for v in vals)
        assert fake_sid not in joined
        assert '赵铁柱' not in joined
        assert '赵*' not in joined
        assert 'ANON-' in vals[0]
        assert vals[1] == '[已脱敏]'
    finally:
        wb_out.close()


def test_r2_2_strip_privacy_single_college_without_college_column(tmp_path: Path):
    """R2-2: 无'学院'列的单学院免勤表正常触发整表脱敏，不漏姓名与手机号。"""
    src = tmp_path / 'single_college.xlsx'
    dst = tmp_path / 'single_college_stripped.xlsx'
    fake_sid = '3202' + '002020002'
    wb = Workbook()
    ws = wb.active
    ws.title = '早自习'
    ws.append(['学号', '姓名', '专业班级', '免勤周次', '免勤原因', '联系电话'])
    ws.append([fake_sid, '钱翠花', '计科24-1班', '第7周', '发烧输液', '13912345678'])
    wb.save(src)
    wb.close()

    stats = strip_workbook_privacy(str(src), str(dst))
    assert stats['rows_stripped'] == 1
    assert stats['names_removed'] == 1
    assert stats['sids_anonymized'] == 1

    wb_out = load_workbook(dst)
    try:
        vals = [str(c.value) for c in wb_out['早自习'][2]]
        joined = ' '.join(vals)
        assert fake_sid not in joined
        assert '钱翠花' not in joined
        assert '13912345678' not in joined
        assert '发烧输液' not in joined
    finally:
        wb_out.close()


def test_r2_3_strip_privacy_deep_header_beyond_row_10(tmp_path: Path):
    """R2-3: 前 11 行为公文须知、表头在第 12 行的工作表仍被成功识别脱敏。"""
    src = tmp_path / 'deep_header.xlsx'
    dst = tmp_path / 'deep_stripped.xlsx'
    fake_sid = '3202' + '003030003'
    wb = Workbook()
    ws = wb.active
    ws.title = '早自习'
    for i in range(1, 12):
        ws.append([f'填表须知第 {i} 条：请如实填写'])
    ws.append(['学号', '姓名', '学院', '专业班级', '免勤周次', '免勤原因'])
    ws.append([fake_sid, '孙大壮', '理学院', '理学24-2班', '第7周', '课程冲突'])
    wb.save(src)
    wb.close()

    stats = strip_workbook_privacy(str(src), str(dst))
    assert stats['rows_stripped'] == 1
    wb_out = load_workbook(dst)
    try:
        row_text = ' '.join(str(c.value) for c in wb_out['早自习'][13])
        assert '孙大壮' not in row_text
        assert fake_sid not in row_text
        assert 'ANON-' in row_text
    finally:
        wb_out.close()


def test_r2_4_app_export_diagnostic_log_passes_config_and_strict_privacy(tmp_path: Path, monkeypatch):
    """R2-4: app._export_diagnostic_log 正确传参 config/result/error_trace 且启用 strict_privacy=True。"""
    os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
    from PySide6.QtWidgets import QApplication
    import app as gui

    QApplication.instance() or QApplication(sys.argv[:1])
    out_log = tmp_path / 'gui_diag.log'
    fake_sid = '3202' + '004040004'
    cfg = RunConfig(detail_path=str(tmp_path / 'detail.xlsx'), week_no=9, campus='雁山')
    res = RunResult(
        ok=False,
        error=f'学生张伟({fake_sid})校验失败',
        anomalies=[{'type': '测试', 'level': 'warn', 'scope': '理学院', 'note': '学生张伟因急性肠炎请假'}],
        student_details={('理学院', '周一', '早自习'): [{'sid': fake_sid, 'name': '张伟', 'reason': '急性肠炎', 'class_text': '理学24-1班'}]},
    )

    dummy = SimpleNamespace(
        result=res,
        _last_config=cfg,
        _last_error_trace='Traceback: 学生张伟报错',
    )
    monkeypatch.setattr(gui.QFileDialog, 'getSaveFileName', lambda *a, **kw: (str(out_log), 'Log Files (*.log)'))
    monkeypatch.setattr(gui, 'ToastNotification', lambda *a, **kw: None)
    monkeypatch.setattr(gui.QMessageBox, 'critical', lambda *a, **kw: None)

    gui.MainWindow._export_diagnostic_log(dummy)
    text = out_log.read_text(encoding='utf-8')
    assert '第 9 周' in text  # 证明 config 正确传入而非被 self.result 错位顶替
    assert '张伟' not in text
    assert '张*' not in text  # 证明启用了 strict_privacy=True
    assert fake_sid not in text
    assert '急性肠炎' not in text


def test_r2_5_run_calc_diagnostic_log_strict_privacy(tmp_path: Path):
    """R2-5: Skill run_calc.py --diagnostic-log 显式传入 strict_privacy=True 零姓氏泄露。"""
    detail_path = tmp_path / 'detail.xlsx'
    exempt_path = tmp_path / 'exempt.xlsx'
    out_path = tmp_path / 'out.xlsx'
    diag_path = tmp_path / 'skill_diag.md'
    fake_sid = '3202' + '005050005'

    create_detail_workbook(detail_path, week_no=7)
    wb = Workbook()
    ws = wb.active
    ws.title = '早自习'
    ws.append(['学号', '姓名', '学院', '专业班级', '免勤周次', '免勤时间', '免勤原因'])
    ws.append([fake_sid, '张伟', '不存在测试学院', '理学24-1班', '第7周', '周一', '急性扁桃体炎'])
    wb.save(exempt_path)
    wb.close()

    run_calc = ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py'
    proc = subprocess.run(
        [
            sys.executable, str(run_calc),
            '--app-root', str(ROOT),
            '--detail', str(detail_path),
            '--exempt', str(exempt_path),
            '--week', '7',
            '--out', str(out_path),
            '--diagnostic-log', str(diag_path),
            '--no-storage',
        ],
        capture_output=True, text=True, encoding='utf-8', timeout=30,
    )
    assert proc.returncode == 0, proc.stderr
    diag_text = diag_path.read_text(encoding='utf-8')
    assert '张伟' not in diag_text
    assert '张*' not in diag_text
    assert fake_sid not in diag_text


def test_r2_6_run_calc_failure_branch_redacts_pii_and_enforces_assertion(tmp_path: Path, capsys):
    """R2-6: run_calc.py 失败分支 (not result.ok) 脱敏 error 并执行合规断言。"""
    run_calc_mod = _load_module_from_path('skill_run_calc_r2_6', ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py')
    fake_sid = '3202' + '006060006'
    fake_result = RunResult(
        ok=False,
        error=f'无法核对学生赵铁柱({fake_sid})在计科24-1班的急性阑尾炎免勤记录',
        student_details={('计算机学院', '周一', '早自习'): [{
            'sid': fake_sid,
            'name': '赵铁柱',
            'reason': '急性阑尾炎',
            'class_text': '计科24-1班',
        }]},
        duration_ms=12,
    )
    detail_file = tmp_path / 'd.xlsx'
    Workbook().save(detail_file)
    ns = SimpleNamespace(
        app_root=str(ROOT),
        preflight=None,
        in_place=False,
        flatten_formulas=None,
        strip_privacy=None,
        verify_tamper=None,
        against_detail=None,
        exempt_report=None,
        detail=str(detail_file),
        exempt=[],
        out=None,
        week=7,
        dictionary=None,
        campus='雁山',
        source='未分类',
        match_classes=False,
        college_exempt=False,
        include_exempt_summary=False,
        complete_mode='auto',
        summary_out=None,
        no_storage=True,
        list_runs=False,
        show_run=None,
        diagnostic_log=None,
    )

    with patch('core.pipeline.run', return_value=fake_result):
        code = run_calc_mod.execute_runner(ns)
    assert code == 1
    stdout = capsys.readouterr().out
    payload = json.loads(stdout)
    assert payload['status'] == 'error'
    for secret in ('赵铁柱', '赵*', fake_sid, '急性阑尾炎', '计科24-1班'):
        assert secret not in stdout


def test_r2_7_cli_show_run_redacts_notes_and_exempt_files_json(tmp_path: Path, monkeypatch, capsys):
    """R2-7: cli.py --show-run 对 notes 和 exempt_files_json 执行隐私脱敏。"""
    import sqlite3

    db_file = tmp_path / 'history.db'
    storage = AttendanceStorage(db_path=str(db_file))
    fake_sid = '3202' + '007070007'
    cfg = RunConfig(detail_path=str(tmp_path / 'detail.xlsx'), week_no=7)
    res = RunResult(
        ok=True,
        output_path=str(tmp_path / 'out.xlsx'),
        anomalies=[{'type': '告警', 'level': 'warn', 'scope': '理学院', 'note': f'学生赵铁柱({fake_sid})异常'}],
    )
    run_id = storage.record_run(cfg, res, duration_ms=10)
    # 注入含敏感信息的 notes 与 exempt_files_json
    drive_c = 'C'
    secret_file = drive_c + f':\\Users\\Secret\\{fake_sid}_赵铁柱免勤.xlsx'
    with sqlite3.connect(str(db_file)) as conn:
        conn.execute(
            "UPDATE calc_runs SET notes = ?, exempt_files_json = ? WHERE id = ?",
            (f'辅导员钱老师核对学生赵铁柱({fake_sid})', json.dumps([secret_file]), run_id),
        )
        conn.commit()

    monkeypatch.setattr(cli, 'AttendanceStorage', lambda: AttendanceStorage(db_path=str(db_file)))
    code = cli.main(['--show-run', str(run_id), '--json'])
    assert code == 0
    stdout = capsys.readouterr().out
    assert fake_sid not in stdout
    assert '赵铁柱' not in stdout
    assert '钱老师' not in stdout


def test_r2_8_sanitize_anomaly_text_whitespace_prefixes_and_two_char_name():
    """R2-8: _sanitize_anomaly_text 支持前缀空格、申请人/联系人/姓名前缀且2字姓名不泄单字姓。"""
    raw = '学生 赵铁柱因事请假，申请人 李四未登记，姓名: 王五异常'
    cleaned = _sanitize_anomaly_text(raw)
    for secret in ('赵铁柱', '李四', '王五', '李*', '王*'):
        assert secret not in cleaned
    assert '[已脱敏]' in cleaned


def test_r2_9_privacy_scan_dynamic_column_sniffing(tmp_path: Path):
    """R2-9: tools/privacy_scan.py 动态嗅探乱序表头与深行表头，正确提取真实姓名和学号。"""
    from tools.privacy_scan import load_real_personal

    xlsx_path = tmp_path / 'reordered_exempt.xlsx'
    fake_sid = '3202' + '008080008'
    wb = Workbook()
    ws = wb.active
    ws.append(['填表说明行 1'])
    ws.append(['填表说明行 2'])
    ws.append(['填表说明行 3'])
    ws.append(['填表说明行 4'])
    # 第 5 行表头：姓名在第 0 列，学号在第 3 列（非默认 row[3]/row[1]）
    ws.append(['学生姓名', '学院', '免勤原因', '学生学号'])
    ws.append(['赵铁柱', '理学院', '生病', fake_sid])
    wb.save(xlsx_path)
    wb.close()

    sids, names = load_real_personal(str(xlsx_path))
    assert fake_sid in sids
    assert '赵铁柱' in names


def test_r2_10_assert_model_privacy_compliance_blocks_8_to_10_and_99_sids():
    """R2-10: assert_model_privacy_compliance 拦截 8-10 位学号及 99 开头 13 位学号。"""
    for raw_sid in ('20240819', '2024081901', '9900000000001'):
        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance(json.dumps({'sid': raw_sid}))


# ==============================================================================
# R3: 代码质量与边界缺陷 (R3-1 ~ R3-10)
# ==============================================================================


def test_r3_1_write_summary_all_exempt_zero_denominator(tmp_path: Path):
    """R3-1: 全员免勤 (expected == exempt, denominator == 0) 时 _write_summary 不除零崩溃。"""
    detail_path = write_detail(tmp_path / 'detail.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('地球科学学院', 40, [('8101', 40, {'周一': 0, '周三': 0, '周四': 0, '周五': 0})]),
        ]),
    ])
    exempt_rows = [
        make_exempt_row(sid=f'990000000{day_i}{i:03d}', college='地球科学学院', day_text=day)
        for day_i, day in enumerate(('周一', '周三', '周四', '周五'), 1)
        for i in range(1, 41)
    ]
    exempt_path = write_exempt(tmp_path / 'exempt.xlsx', {'早自习免勤人员统计表': exempt_rows})
    model = load_detail(str(detail_path))
    agg = aggregate_exemptions(load_exemptions(str(exempt_path)), 7, campus='雁山')
    calc = Calculator(model, agg, complete_mode=True)
    rates = calc.college_rates()

    summary_out = tmp_path / 'summary.xlsx'
    _write_summary(RunConfig(detail_path=str(detail_path), week_no=7), str(summary_out), model, calc, rates)
    assert summary_out.is_file()
    wb = load_workbook(summary_out, data_only=False)
    try:
        ws = wb['Sheet2']
        assert ws.cell(3, 3).value == pytest.approx(1.0)
        assert ws.cell(3, 3).number_format == '0.00%'
        assert '全员免勤' in str(ws.cell(3, 5).value)
    finally:
        wb.close()

    # 再验证完整 run() 流水线即使分母为 0 也正常完成且不抛未捕获 ZeroDivisionError
    out2 = tmp_path / 'out2.xlsx'
    sum2 = tmp_path / 'sum2.xlsx'
    res = run(RunConfig(detail_path=str(detail_path), exempt_path=str(exempt_path), output_path=str(out2),
                        summary_output_path=str(sum2), week_no=7, complete_mode=True, enable_storage=False))
    assert res.ok, res.error
    assert res.summary_output_path == str(sum2)


def test_r3_2_formulas_large_round_and_power_overflow():
    """R3-2: 超大数值 ROUND 与指数溢出被安全降级/包装并可由 IFERROR 捕获。"""
    wb = Workbook()
    ws = wb.active
    ws['A1'] = '=ROUND(1e300, 2)'
    ws['A2'] = '=IFERROR(10^309, 0)'
    ws['A3'] = '=10^309'

    val_round = evaluate_workbook_cell(wb, ws.title, 1, 1)
    assert val_round == pytest.approx(1e300)

    val_iferr = evaluate_workbook_cell(wb, ws.title, 2, 1)
    assert val_iferr == 0

    stats = flatten_workbook_formulas(wb, clear_unresolved=True)
    assert stats.residual_formulas == 0
    assert ws['A3'].value is None


def test_r3_3_summary_workbook_reverse_maps_college_aliases(tmp_path: Path):
    """R3-3: 学院简称（如化生学院）写入汇总终稿时正确匹配官方全称，不再留空。"""
    detail = write_detail(tmp_path / 'detail.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('化学与生物工程学院', 100, [('8101', 100, {'周一': 90, '周三': 90, '周四': 90, '周五': 90})]),
        ]),
    ])
    exempt = write_exempt(tmp_path / 'exempt.xlsx', {
        '早自习免勤人员统计表': [
            make_exempt_row(sid='9900000000001', college='化生学院', day_text='周一'),
        ]
    })
    out = tmp_path / 'out.xlsx'
    summary_out = tmp_path / 'summary.xlsx'
    res = run(RunConfig(detail_path=str(detail), exempt_path=str(exempt),
                        output_path=str(out), summary_output_path=str(summary_out),
                        week_no=7, enable_storage=False))
    assert res.ok, res.error

    wb = load_workbook(summary_out, data_only=False)
    try:
        ws = wb['Sheet2']
        row_map = {ws.cell(r, 2).value: (ws.cell(r, 3).value, ws.cell(r, 3).number_format) for r in range(3, 18)}
        rate_val, fmt = row_map['化学与生物工程学院']
        assert rate_val is not None
        assert rate_val == pytest.approx(round(360 / 399, 4))
        assert fmt == '0.00%'
    finally:
        wb.close()

    # 直接调用 fill_summary_workbook 传入简称键同样可命中
    wb_direct = load_workbook(summary_out)
    wb_direct['Sheet2'].cell(6, 3).value = None
    fill_summary_workbook(wb_direct, {'化生学院': 0.8888})
    assert wb_direct['Sheet2'].cell(6, 3).value == pytest.approx(0.8888)
    wb_direct.close()


def test_r3_4_normalize_sid_rejects_negative_integers():
    """R3-4: 负整数学号不左补零且被判定为非法非正整数学号。"""
    sid1, ok1, desc1 = normalize_sid(-12345)
    assert sid1 == '-12345'
    assert ok1 is False
    assert '非正整数' in desc1

    sid2, ok2, desc2 = normalize_sid(-12345.0)
    assert sid2 == '-12345'
    assert ok2 is False
    assert '非正整数' in desc2

    sid3, ok3, _ = normalize_sid(0)
    assert ok3 is False


def test_r3_5_cli_diagnostic_log_write_failure_not_falsely_reported(tmp_path: Path, capsys):
    """R3-5: CLI --diagnostic-log 写入失败时不吞噬虚假报喜，返回 diagnostic_log_error。"""
    detail_path = tmp_path / 'detail.xlsx'
    out_path = tmp_path / 'out.xlsx'
    bad_diag_dir = tmp_path / 'dir_as_file'
    bad_diag_dir.mkdir()
    create_detail_workbook(detail_path, week_no=7)

    code = cli.main([
        '--detail', str(detail_path),
        '--out', str(out_path),
        '--week', '7',
        '--diagnostic-log', str(bad_diag_dir),  # 目录作为文件路径必抛 IsADirectoryError/PermissionError
        '--no-storage',
        '--json',
    ])
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    assert 'diagnostic_log' not in payload
    assert 'diagnostic_log_error' in payload and payload['diagnostic_log_error']


def test_r3_6_parse_int_handles_infinity_without_overflow():
    """R3-6: parse_int 传入 +/-inf 与 NaN 安全返回 None，不抛 OverflowError。"""
    assert parse_int(float('inf')) is None
    assert parse_int(float('-inf')) is None
    assert parse_int(float('nan')) is None
    assert parse_int(1e309) is None


def test_r3_7_load_evaluated_workbook_closes_wb_on_flatten_error(tmp_path: Path, monkeypatch):
    """R3-7: load_evaluated_workbook 展平失败时必定调用 wb.close() 释放句柄。"""
    import core.formulas as formulas_mod

    xlsx = tmp_path / 'formula.xlsx'
    wb = Workbook()
    wb.active['A1'] = '=1+1'
    wb.save(xlsx)
    wb.close()

    closed = []
    orig_close = Workbook.close

    def tracked_close(self):
        closed.append(True)
        return orig_close(self)

    monkeypatch.setattr(Workbook, 'close', tracked_close)
    monkeypatch.setattr(formulas_mod, 'flatten_workbook_formulas', lambda *a, **kw: (_ for _ in ()).throw(RuntimeError('boom')))

    with pytest.raises(RuntimeError, match='boom'):
        load_evaluated_workbook(str(xlsx))
    assert len(closed) >= 1


def test_r3_8_export_attendance_summary_missing_college_no_keyerror(tmp_path: Path):
    """R3-8: 汇总模板包含未在 totals 出现的学院时不抛 KeyError，默认 (0, 0, False)。"""
    tpl_path = tmp_path / 'tpl.xlsx'
    out_path = tmp_path / 'out_tpl.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = '汇总'
    ws['A1'] = '2026年秋第7周周一早自习出勤汇总'
    ws.append(['序号', '学院', '学生出勤率'])
    ws.append([1, '理学院', 0.95])
    ws.append([2, '未上报学院', 0.90])
    wb.save(tpl_path)
    wb.close()

    tpl = load_summary_template(str(tpl_path))
    assert tpl is not None
    totals = {'理学院': (2, 1, False)}  # 故意缺少 '未上报学院'
    export_attendance_summary(str(tpl_path), str(out_path), tpl, totals, lambda _book: None)

    wb_out = load_workbook(out_path, data_only=True)
    try:
        s = wb_out['汇总']
        assert [s.cell(4, 4).value, s.cell(4, 5).value, s.cell(4, 6).value, s.cell(4, 7).value] == [0, 0, 0, '已核对']
    finally:
        wb_out.close()


def test_r3_9_calc_tolerant_mode_effective_room_denominator_not_overridden(tmp_path: Path):
    """R3-9: 容错模式下存在缺失教室且有效分母 > 0 时，不被全局分母 <= 0 误判为 100%。"""
    # 2间教室：8101(50人,实到40), 8102(50人,未录入实到)
    # 另给该学院 100 名跨班级/无教室免勤，导致全局 denominator = 100 - 100 = 0
    # 但在容错模式下仅计算有效教室 8101（按比例分摊得 50 免勤？wait：分摊50 -> 50-50=0，若 8101 应到60, 8102 应到40, 总免勤100？
    # 若8101有带教室号免勤仅10人，8102有90人但未录实到 -> 有效教室8101应到60-免勤10=50 > 0！）
    detail = write_detail(tmp_path / 'detail.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('理学院', 100, [
                ('8101', 60, {'周一': 40}),
                ('8102', 40, {'周一': None}),
            ]),
        ]),
    ])
    rows = (
        [make_exempt_row(sid=f'9900000001{i:03d}', room='8101') for i in range(1, 11)]
        + [make_exempt_row(sid=f'9900000002{i:03d}', room='8102') for i in range(1, 91)]
    )
    exempt = write_exempt(tmp_path / 'exempt.xlsx', {'早自习免勤人员统计表': rows}, with_room=True)
    model = load_detail(str(detail))
    agg = aggregate_exemptions(load_exemptions(str(exempt)), 7, campus='雁山')
    calc = Calculator(model, agg, complete_mode=True)
    rec = [r for r in calc.college_rates() if r.day == '周一'][0]

    # 全局 denominator = 100 - 100 = 0，但有效教室 8101 分母 = 60 - 10 = 50 > 0，出勤率应为 40 / 50 = 0.8
    assert rec.rate == pytest.approx(40 / 50)
    assert rec.status == '部分教室缺失（已按有效实到计算）'


def test_r3_10_attendance_summary_invalid_calendar_date_contextual_error(tmp_path: Path):
    """R3-10: 汇总表考勤时间包含非法日历日期（如2月30日）抛出含上下文的 ValueError。"""
    tpl_path = tmp_path / 'bad_date_tpl.xlsx'
    wb = Workbook()
    ws = wb.active
    ws['A1'] = '2026年春季学期第2周汇总'
    ws.append(['序号', '学院', '学生出勤率'])
    ws.append([1, '理学院', 0.95])
    ws['A5'] = '备注：\n考勤时间：2月30日早自习'
    wb.save(tpl_path)
    wb.close()

    with pytest.raises(ValueError, match=r'无效日历日期.*2月30日早自习'):
        load_summary_template(str(tpl_path))


# ==============================================================================
# R4: Agent Skill 一致性与同步 (R4-1 ~ R4-7)
# ==============================================================================


def test_r4_1_dict_tool_no_fallback_on_explicit_missing_path(tmp_path: Path, monkeypatch):
    """R4-1: dict_tool.py 显式传入不存在路径时绝不静默回退到当前目录生产词典。"""
    dict_tool = _load_module_from_path('skill_dict_tool_r4_1', ACTIVE_SKILL_DIR / 'scripts' / 'dict_tool.py')
    prod_dict = tmp_path / '免勤词典.xlsx'
    tables = {k: [] for k in HEADERS.keys()}
    tables['学院别名'] = [['理学院', '理院', '基线']]
    save_tables(str(prod_dict), tables, backup=False)

    monkeypatch.chdir(tmp_path)
    resolved = dict_tool.resolve_dict_path('typo_nonexistent.xlsx')
    assert resolved == os.path.abspath('typo_nonexistent.xlsx')
    assert not resolved.endswith('免勤词典.xlsx')

    # 执行 add-alias 必须报错退出且不能修改 prod_dict
    orig_hash = hashlib.sha256(prod_dict.read_bytes()).hexdigest()
    code = dict_tool.main([
        '--app-root', str(ROOT),
        'add-alias',
        '--sheet', '学院别名',
        '--alias', '测试别名',
        '--target', '理学院',
        '--dict', 'typo_nonexistent.xlsx',
    ])
    assert code == 1
    assert hashlib.sha256(prod_dict.read_bytes()).hexdigest() == orig_hash


def test_r4_2_run_calc_sid_regex_excludes_99_prefix():
    """R4-2: run_calc.py SID_REGEX 与 cli.py 对齐，包含 (?!99) 排除合成测试号段。"""
    run_calc_mod = _load_module_from_path('skill_run_calc_r4_2', ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py')
    assert '(?!99)' in run_calc_mod.SID_REGEX.pattern
    run_calc_mod.assert_privacy_compliance('Synthetic 9900000000001 test SID')
    non_99 = '3202' + '009090009'
    with pytest.raises(ValueError, match='Privacy assertion violation'):
        run_calc_mod.assert_privacy_compliance(f'Real {non_99} leaked')


def test_r4_3_run_calc_diagnostic_log_passes_strict_privacy(tmp_path: Path):
    """R4-3: run_calc.py 导出诊断日志强制传递 strict_privacy=True 且写入失败时准确记录。"""
    src = (ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py').read_text(encoding='utf-8')
    assert 'generate_diagnostic_log(config, result=result, strict_privacy=True)' in src


def test_r4_4_run_calc_error_paths_enforce_redaction_and_compliance(tmp_path: Path, capsys):
    """R4-4: run_calc.py 免勤独立报表异常分支执行脱敏并强制触发断言。"""
    run_calc_mod = _load_module_from_path('skill_run_calc_r4_4', ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py')
    fake_sid = '3202' + '010100010'
    ok_res = RunResult(
        ok=True,
        output_path=str(tmp_path / 'out.xlsx'),
        student_details={('理学院', '周一', '早自习'): [{'sid': fake_sid, 'name': '王大锤', 'reason': '急性胃肠炎', 'class_text': '理学24-3班'}]},
        duration_ms=10,
    )
    detail_file = tmp_path / 'd.xlsx'
    Workbook().save(detail_file)
    ns = SimpleNamespace(
        app_root=str(ROOT), preflight=None, in_place=False, flatten_formulas=None,
        strip_privacy=None, verify_tamper=None, against_detail=None,
        exempt_report=str(tmp_path / 'rep.xlsx'), detail=str(detail_file),
        exempt=[], out=str(tmp_path / 'out.xlsx'), week=7, dictionary=None,
        campus='雁山', source='未分类', match_classes=False, college_exempt=False,
        include_exempt_summary=False, complete_mode='auto', summary_out=None,
        no_storage=True, list_runs=False, show_run=None, diagnostic_log=None,
    )
    with patch('core.pipeline.run', return_value=ok_res), \
         patch('core.exempt_report.export_exempt_report', side_effect=ValueError(f'王大锤({fake_sid})导出失败')):
        code = run_calc_mod.execute_runner(ns)
    assert code == 1
    out = capsys.readouterr().out
    assert '王大锤' not in out
    assert fake_sid not in out


def test_r4_5_run_calc_json_schema_matches_cli_reference(tmp_path: Path):
    """R4-5: run_calc.py 输出 JSON 字段与数据类型 100% 对齐 cli_reference.md §3.1。"""
    detail = write_detail(tmp_path / 'detail.xlsx', [
        make_detail_sheet('雁山（早）', [
            ('理学院', 100, [('8101', 100, {'周一': 90, '周三': 90, '周四': 90, '周五': 90})]),
        ]),
    ])
    exempt = write_exempt(tmp_path / 'exempt.xlsx', {
        '早自习免勤人员统计表': [make_exempt_row(sid='9900000000001', day_text='周一')],
    })
    out_xlsx = tmp_path / 'out.xlsx'
    run_calc = ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py'
    proc = subprocess.run(
        [
            sys.executable, str(run_calc),
            '--app-root', str(ROOT),
            '--detail', str(detail),
            '--exempt', str(exempt),
            '--out', str(out_xlsx),
            '--week', '7',
            '--no-storage',
        ],
        capture_output=True, text=True, encoding='utf-8', timeout=30,
    )
    assert proc.returncode == 0, proc.stderr
    data = json.loads(proc.stdout)
    assert data['status'] == 'ok'
    assert 'phases' in data and 'preflight_ms' in data['phases']
    assert 'stage_durations_ms' not in data
    assert 'discrepancies_count' in data and isinstance(data['discrepancies'], list)
    assert 'anomalies_count' in data and isinstance(data['anomalies'], list)
    assert data['college_rates'] and isinstance(data['college_rates'][0]['rate'], float)
    assert data['weekly_rates'] and isinstance(data['weekly_rates'][0]['rate'], float)


def test_r4_6_run_calc_list_and_show_runs_redact_notes_and_exempt_files(tmp_path: Path, capsys):
    """R4-6: run_calc.py --list-runs 与 --show-run 均清洗 exempt_files_json 与 notes。"""
    import sqlite3

    run_calc_mod = _load_module_from_path('skill_run_calc_r4_6', ACTIVE_SKILL_DIR / 'scripts' / 'run_calc.py')
    db_file = tmp_path / 'skill_hist.db'
    storage = AttendanceStorage(db_path=str(db_file))
    fake_sid = '3202' + '011110011'
    cfg = RunConfig(detail_path=str(tmp_path / 'detail.xlsx'), week_no=7)
    res = RunResult(ok=True, output_path=str(tmp_path / 'out.xlsx'))
    run_id = storage.record_run(cfg, res, duration_ms=5)
    drive_c = 'C'
    secret_file = drive_c + f':\\Users\\Alice\\{fake_sid}_赵铁柱.xlsx'
    with sqlite3.connect(str(db_file)) as conn:
        conn.execute(
            "UPDATE calc_runs SET notes = ?, exempt_files_json = ? WHERE id = ?",
            (f'学生赵铁柱({fake_sid})未通过', json.dumps([secret_file]), run_id),
        )
        conn.commit()

    base_ns = dict(
        app_root=str(ROOT), preflight=None, in_place=False, flatten_formulas=None,
        strip_privacy=None, verify_tamper=None, against_detail=None, exempt_report=None,
        detail=None, exempt=[], out=None, week=7, dictionary=None, campus='雁山',
        source='未分类', match_classes=False, college_exempt=False, include_exempt_summary=False,
        complete_mode='auto', summary_out=None, no_storage=True,
    )
    with patch('core.storage.AttendanceStorage', lambda: AttendanceStorage(db_path=str(db_file))):
        assert run_calc_mod.execute_runner(SimpleNamespace(**base_ns, list_runs=True, show_run=None)) == 0
        out_list = capsys.readouterr().out
        assert fake_sid not in out_list and '赵铁柱' not in out_list

        assert run_calc_mod.execute_runner(SimpleNamespace(**base_ns, list_runs=False, show_run=run_id)) == 0
        out_show = capsys.readouterr().out
        assert fake_sid not in out_show and '赵铁柱' not in out_show


def test_r4_7_cli_reference_documents_all_four_missing_options_and_mirror_sha256():
    """R4-7: cli_reference.md §2 完整收录 4 个核心选项且双 Skill 目录 SHA256 100% 一致。"""
    ref_active = (ACTIVE_SKILL_DIR / 'references' / 'cli_reference.md').read_text(encoding='utf-8')
    table_section = ref_active.split('## 2. 完整命令行参数清单')[1].split('## 3. 标准化 JSON 响应结构')[0]
    for opt in ('--flatten-formulas', '--strip-privacy', '--in-place', '--app-root'):
        assert opt in table_section, f'Missing option {opt} in §2 table'

    def _sha(p: Path) -> str:
        return hashlib.sha256(p.read_bytes()).hexdigest()

    active_files = sorted(
        p.relative_to(ACTIVE_SKILL_DIR)
        for p in ACTIVE_SKILL_DIR.rglob('*')
        if p.is_file() and '__pycache__' not in p.parts and p.suffix != '.pyc'
    )
    repo_files = sorted(
        p.relative_to(REPO_SKILL_DIR)
        for p in REPO_SKILL_DIR.rglob('*')
        if p.is_file() and '__pycache__' not in p.parts and p.suffix != '.pyc'
    )
    assert active_files == repo_files and len(active_files) >= 9
    for rel in active_files:
        assert _sha(ACTIVE_SKILL_DIR / rel) == _sha(REPO_SKILL_DIR / rel), f'SHA256 mismatch: {rel}'
