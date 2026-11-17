# -*- coding: utf-8 -*-
"""公式求值、计算前清公式展平与零公式纯数值输出回归测试。"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import cli  # noqa: E402
from core import RunConfig, run  # noqa: E402
from core.export import compare_actuals, verify_workbook_integrity  # noqa: E402
from core.formulas import (  # noqa: E402
    count_workbook_formulas,
    evaluate_workbook_cell,
    export_flattened_workbook,
    flatten_workbook_formulas,
    is_formula_cell,
)


def _assert_zero_formulas(xlsx_path: Path) -> None:
    wb = load_workbook(xlsx_path, data_only=False)
    try:
        leaked = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if is_formula_cell(cell):
                        leaked.append(f'{ws.title}!{cell.coordinate}={cell.value!r}')
        assert not leaked, f'发现未清除的公式单元格: {leaked}'
    finally:
        wb.close()


class TestFormulaEvaluator:
    def test_evaluates_arithmetic_ranges_and_cross_sheet_refs(self):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = '雁山（早）'
        ws1['A1'] = 40
        ws1['A2'] = 65
        ws1['A3'] = '=SUM(A1:A2)'
        ws1['A4'] = '=IF(A1=0, 1/0, ROUND(A3/111, 4))'
        ws1['A5'] = '=IFERROR(1/0, 0)'
        ws1['B1'] = '甲学院'
        ws1['B2'] = '乙学院'
        ws1['B3'] = '=SUMIF(B1:B2, "甲学院", A1:A2)'
        ws1['B4'] = '=COUNTIF(A1:A2, ">=50")'

        ws2 = wb.create_sheet('屏风')
        ws2['C1'] = "='雁山（早）'!A3+5"

        assert evaluate_workbook_cell(wb, '雁山（早）', 3, 1) == 105
        assert evaluate_workbook_cell(wb, '雁山（早）', 4, 1) == pytest.approx(0.9459)
        assert evaluate_workbook_cell(wb, '雁山（早）', 5, 1) == 0
        assert evaluate_workbook_cell(wb, '雁山（早）', 3, 2) == 40
        assert evaluate_workbook_cell(wb, '雁山（早）', 4, 2) == 1
        assert evaluate_workbook_cell(wb, '屏风', 1, 3) == 110

    def test_flatten_preserves_integer_counts_and_two_decimal_percent_rates(self):
        wb = Workbook()
        ws = wb.active
        ws.title = '汇总'
        ws.append(['学院', '应到人数', '免勤人数', '实到人数', '学生出勤率', '未知公式'])
        ws.append(['测试学院', 150, '=20+19', '=50+55', '=D2/(B2-C2)', '=UNKNOWNFUNC(1,2)'])
        ws['E2'].number_format = '0.00%'

        stats = flatten_workbook_formulas(wb, clear_unresolved=True)
        assert stats.total_formula_cells == 4
        assert stats.evaluated_formulas == 3
        assert stats.cleared_unresolved == 1
        assert stats.residual_formulas == 0
        assert count_workbook_formulas(wb) == 0

        # 免勤与实到为纯整数数字 + '0'
        assert ws['C2'].value == 39 and isinstance(ws['C2'].value, int)
        assert ws['C2'].number_format == '0'
        assert ws['D2'].value == 105 and isinstance(ws['D2'].value, int)
        assert ws['D2'].number_format == '0'

        # 出勤率为浮点比例 + '0.00%'
        expected_rate = 105 / (150 - 39)
        assert ws['E2'].value == pytest.approx(expected_rate)
        assert isinstance(ws['E2'].value, float)
        assert ws['E2'].number_format == '0.00%'

        # 未识别公式被清空，不保留 '=...'
        assert ws['F2'].value is None
        assert ws['F2'].data_type != 'f'

    def test_cycle_detection_clears_safely(self):
        wb = Workbook()
        ws = wb.active
        ws['A1'] = '=B1+1'
        ws['B1'] = '=A1+1'
        stats = flatten_workbook_formulas(wb, clear_unresolved=True)
        assert stats.cleared_unresolved == 2
        assert stats.residual_formulas == 0
        assert count_workbook_formulas(wb) == 0


from openpyxl.utils import get_column_letter
from core.io_detail import _detect_data_cols, _find_header_row  # noqa: E402
from tests_e2e.helpers import create_detail_workbook  # noqa: E402


class TestEndToEndZeroFormulaPipeline:
    @staticmethod
    def _build_detail_with_uncached_formulas(path: Path) -> Dict[str, tuple[int, int]]:
        create_detail_workbook(
            path,
            early_blocks=[
                ('地球科学学院', 150, [
                    ('01101', 150, {'周一': 105, '周三': 140, '周四': 145, '周五': 148}),
                ]),
            ],
            late_blocks=[
                ('地球科学学院', 150, [
                    ('01101', 150, {'周一': 140, '周三': 142, '周四': 145}),
                ]),
            ],
            week_no=7,
        )
        wb = load_workbook(path)
        coords: Dict[str, tuple[int, int]] = {}
        ws_early = wb['雁山（早）']
        hdr = _find_header_row(ws_early)
        cols = _detect_data_cols(ws_early, hdr)
        data_row = hdr + 2  # row 2=表头, row 3=出勤率子表头, row 4=首行教室数据
        if ws_early.cell(data_row, cols['room']).value is None:
            data_row = hdr + 1

        exp_col = cols['expected']
        mon_act_col = cols['actuals']['周一']
        mon_ex_col = cols['exempts']['周一']
        ws_early.cell(data_row, exp_col, '=100+50')
        ws_early.cell(data_row, mon_act_col, '=50+55')
        if mon_ex_col:
            ws_early.cell(data_row, mon_ex_col, '=10+1')

        # 寻找出勤率列（周一）并放入待覆盖公式
        rate_col = None
        for c in range(1, ws_early.max_column + 1):
            if str(ws_early.cell(hdr + 1, c).value or '').strip() == '周一' and c not in cols['actuals'].values():
                rate_col = c
                ws_early.cell(
                    data_row,
                    c,
                    f'={get_column_letter(mon_act_col)}{data_row}/{get_column_letter(exp_col)}{data_row}',
                )
                ws_early.cell(data_row, c).number_format = '0.00%'
                break

        # 尾部新增受保护的人工巡查列，注入公式（必须同样固化为数字）
        patrol_col = ws_early.max_column + 1
        ws_early.cell(hdr, patrol_col, '人工巡查')
        ws_early.cell(data_row, patrol_col, '=2+3')

        ws_pf = wb['屏风']
        ws_pf['J2'] = '=10+20'
        ws_pf['K2'] = '=J2/100'
        ws_pf['K2'].number_format = '0.00%'

        wb.save(path)
        wb.close()
        coords['expected'] = (data_row, exp_col)
        coords['actual'] = (data_row, mon_act_col)
        coords['exempt'] = (data_row, mon_ex_col)
        coords['rate'] = (data_row, rate_col)
        coords['patrol'] = (data_row, patrol_col)
        return coords

    @staticmethod
    def _build_exempt_file(path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = '早自习'
        ws.append(['学号', '姓名', '学院', '所在班级', '所在教室', '免勤周次', '免勤时间', '免勤原因'])
        for idx in range(1, 40):
            ws.append([
                f'990000000{idx:04d}',
                f'测试学生{idx}',
                '地球科学学院',
                '地球科学学院26-1班',
                '01101',
                '第7周',
                '周一',
                '学院集体活动',
            ])
        wb.save(path)
        wb.close()

    def test_pipeline_outputs_zero_formulas_and_exact_number_formats(self, tmp_path: Path):
        detail_path = tmp_path / 'detail_formula.xlsx'
        exempt_path = tmp_path / 'exempt.xlsx'
        out_path = tmp_path / 'out_result.xlsx'
        summary_out = tmp_path / 'out_summary.xlsx'
        coords = self._build_detail_with_uncached_formulas(detail_path)
        self._build_exempt_file(exempt_path)

        res = run(RunConfig(
            detail_path=str(detail_path),
            exempt_path=str(exempt_path),
            output_path=str(out_path),
            summary_output_path=str(summary_out),
            week_no=7,
            enable_storage=False,
        ))
        assert res.ok, res.error

        _assert_zero_formulas(out_path)
        _assert_zero_formulas(summary_out)

        wb_out = load_workbook(out_path, data_only=False)
        try:
            ws_early = wb_out['雁山（早）']
            exp_cell = ws_early.cell(*coords['expected'])
            act_cell = ws_early.cell(*coords['actual'])
            ex_cell = ws_early.cell(*coords['exempt'])
            rate_cell = ws_early.cell(*coords['rate'])
            patrol_cell = ws_early.cell(*coords['patrol'])

            assert exp_cell.value == 150 and isinstance(exp_cell.value, int)
            assert act_cell.value == 105 and isinstance(act_cell.value, int)
            assert ex_cell.value == 39 and isinstance(ex_cell.value, int)
            assert ex_cell.number_format == '0'
            assert rate_cell.value == pytest.approx(105 / (150 - 39))
            assert isinstance(rate_cell.value, float)
            assert rate_cell.number_format == '0.00%'
            assert patrol_cell.value == 5 and isinstance(patrol_cell.value, int)

            ws_pf = wb_out['屏风']
            assert ws_pf['J2'].value == 30 and isinstance(ws_pf['J2'].value, int)
            assert ws_pf['K2'].value == pytest.approx(0.30)
            assert ws_pf['K2'].number_format == '0.00%'
        finally:
            wb_out.close()

        assert compare_actuals(str(detail_path), str(out_path)) == []
        ok_tamper, violations = verify_workbook_integrity(str(out_path))
        assert ok_tamper, violations

    def test_attendance_summary_mode_outputs_zero_formulas_and_integer_counts(self, tmp_path: Path):
        summary_in = tmp_path / 'summary_tpl.xlsx'
        exempt_path = tmp_path / 'exempt.xlsx'
        out_path = tmp_path / 'summary_result.xlsx'
        self._build_exempt_file(exempt_path)

        wb = Workbook()
        ws = wb.active
        ws.title = '汇总'
        ws['A1'] = '2026年秋第7周周一早自习出勤汇总'
        ws.append(['序号', '学院', '学生出勤率', '巡查备注'])
        ws.cell(3, 1, 1)
        ws.cell(3, 2, '地球科学学院')
        ws.cell(3, 3, '=95/100')
        ws.cell(3, 3).number_format = '0.00%'
        ws.cell(3, 4, '=1+2')
        wb.save(summary_in)
        wb.close()

        res = run(RunConfig(
            detail_path=str(summary_in),
            exempt_path=str(exempt_path),
            output_path=str(out_path),
            week_no=7,
            enable_storage=False,
        ))
        assert res.ok, res.error
        _assert_zero_formulas(out_path)

        wb_out = load_workbook(out_path, data_only=False)
        try:
            s = wb_out['汇总']
            assert s.cell(3, 3).value == pytest.approx(0.95)
            assert s.cell(3, 3).number_format == '0.00%'
            assert s.cell(3, 4).value == 3 and isinstance(s.cell(3, 4).value, int)
            # 早自习免勤人次 = 39，晚自习免勤人次 = 0，合计 = 39（均为 int + '0'）
            assert s.cell(3, 5).value == 39 and s.cell(3, 5).number_format == '0'
            assert s.cell(3, 6).value == 0 and s.cell(3, 6).number_format == '0'
            assert s.cell(3, 7).value == 39 and s.cell(3, 7).number_format == '0'
        finally:
            wb_out.close()

    def test_cli_flatten_formulas_before_calc_and_in_place(self, tmp_path: Path, capsys):
        src = tmp_path / 'raw_with_formulas.xlsx'
        out = tmp_path / 'pure_values.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = '数据'
        ws.append(['学院', '实到', '免勤', '出勤率'])
        ws.append(['测试院', '=50+55', '=30+9', '=B2/111'])
        ws['D2'].number_format = '0.00%'
        wb.save(src)
        wb.close()

        code = cli.main(['--flatten-formulas', str(src), '--out', str(out), '--json'])
        assert code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload['action'] == 'flatten_formulas'
        assert payload['status'] == 'ok'
        assert payload['total_formula_cells'] == 3
        assert payload['residual_formulas'] == 0
        _assert_zero_formulas(out)

        # 确认原始文件未被改动
        assert count_workbook_formulas(load_workbook(src, data_only=False)) == 3

        # 测试 --in-place 就地原子替换
        code2 = cli.main(['--flatten-formulas', str(src), '--in-place', '--json'])
        assert code2 == 0
        payload2 = json.loads(capsys.readouterr().out)
        assert payload2['residual_formulas'] == 0
        _assert_zero_formulas(src)
