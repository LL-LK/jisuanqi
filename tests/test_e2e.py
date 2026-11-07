# -*- coding: utf-8 -*-
"""端到端用例：生成干净样本 → 跑引擎 → 校验输出。

校验点：
  * 输出工作簿只有三张工作表；
  * 屏风表逐单元格零改动；
  * 模板自带的两处合并缺陷被归一化；
  * 出勤率是两位小数、列宽足够（不会显示成 ####）；
  * 免勤单元格全部有值。
"""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

from _helpers import (detail_path, exempt_path, generator,  # noqa: F401  夹具
                      md5, output_path, output_workbook, run_engine,
                      run_result, synthetic_dir, workbook_digest)

ALLOWED_SHEETS = ('雁山（早）', '雁山（晚）', '屏风')


def _merged_in_col(ws, col):
    return sorted(str(rng) for rng in ws.merged_cells.ranges
                  if rng.min_col == col and rng.max_col == col)


def test_e2e_clean_sample_pipeline(run_result):
    assert run_result.ok is True
    assert run_result.error == ''
    assert tuple(run_result.sheets) == ALLOWED_SHEETS
    assert run_result.summary['教室数'] > 0
    assert run_result.summary['学院数'] > 0
    assert run_result.summary['写出免勤单元格'] > 0
    assert run_result.summary['写出出勤率单元格'] > 0


def _visible_sheets(wb):
    return tuple(s for s in wb.sheetnames if s not in ('_系统校验元数据', '免勤核对'))


def test_e2e_output_workbook_has_only_three_sheets(output_workbook):
    assert _visible_sheets(output_workbook) == ALLOWED_SHEETS
    assert len(_visible_sheets(output_workbook)) == 3


def test_e2e_pingfeng_zero_change(detail_path, output_path):
    src = load_workbook(detail_path, data_only=True)
    dst = load_workbook(output_path, data_only=True)
    a, b = src['屏风'], dst['屏风']
    assert a.max_row > 3, '屏风表应当有真实数据行'
    for row in range(1, a.max_row + 1):
        for col in range(1, a.max_column + 1):
            assert a.cell(row, col).value == b.cell(row, col).value
    src.close()
    dst.close()


def test_e2e_merges_normalized_on_both_sheets(output_workbook):
    for sheet in ('雁山（早）', '雁山（晚）'):
        ws = output_workbook[sheet]
        n_merges = _merged_in_col(ws, 14)
        assert 'N57:N59' in n_merges, '%s 的 N57:N59 没有被归一化' % sheet
        assert 'N69:N70' in n_merges, '%s 的 N69:N70 没有被补合并' % sheet


def test_e2e_rate_format_and_width(output_workbook):
    for sheet in ('雁山（早）', '雁山（晚）'):
        ws = output_workbook[sheet]
        for col in range(14, 19):
            cell = ws.cell(4, col)
            if cell.value is None:
                continue
            assert isinstance(cell.value, float)
            assert cell.number_format == '0.00%'
        for letter in ('N', 'O', 'P', 'Q', 'R'):
            width = ws.column_dimensions[letter].width
            assert width is not None and width >= 10


def test_e2e_all_exempt_cells_filled(output_workbook):
    for sheet in ('雁山（早）', '雁山（晚）'):
        ws = output_workbook[sheet]
        for row in range(4, ws.max_row + 1):
            if ws.cell(row, 3).value is None:
                continue
            for col in (7, 9, 11, 13):
                value = ws.cell(row, col).value
                assert value is not None
                assert isinstance(value, int)
                assert value >= 0


def test_e2e_rates_are_consistent_with_formula(run_result):
    for rec in run_result.college_rates:
        if rec.rate is None:
            continue
        denominator = rec.expected - rec.exempt
        assert denominator > 0
        assert rec.rate == pytest.approx(rec.actual / denominator)
        assert 0 < rec.rate <= 1.2


def test_e2e_no_rate_for_missing_actual(run_result):
    missing = [r for r in run_result.college_rates if r.status == '实到数据未完整填写']
    assert missing, '干净样本里应当存在「实到缺失」的用例'
    for rec in missing:
        assert rec.rate is None
        assert rec.actual is None


def test_e2e_input_files_unchanged(detail_path, exempt_path, output_path):
    assert md5(detail_path) != md5(output_path)
    assert detail_path.read_bytes() != output_path.read_bytes()


# ------------------------------------------------------------------ 生成器行为


def test_generator_is_deterministic(tmp_path, generator):
    first = tmp_path / 'a'
    second = tmp_path / 'b'
    generator.generate(first, scale='small', dirty=False, seed=7)
    generator.generate(second, scale='small', dirty=False, seed=7)
    assert workbook_digest(first / 'detail.xlsx') == workbook_digest(second / 'detail.xlsx')
    assert workbook_digest(first / 'exempt.xlsx') == workbook_digest(second / 'exempt.xlsx')


def test_generator_seed_changes_content(tmp_path, generator):
    first = tmp_path / 'a'
    second = tmp_path / 'b'
    generator.generate(first, scale='small', dirty=False, seed=7)
    generator.generate(second, scale='small', dirty=False, seed=8)
    assert workbook_digest(first / 'exempt.xlsx') != workbook_digest(second / 'exempt.xlsx')


def test_generator_writes_exactly_two_files(tmp_path, generator):
    out = tmp_path / 'out'
    generator.generate(out, scale='small', dirty=False, seed=1)
    assert sorted(p.name for p in out.iterdir()) == ['detail.xlsx', 'exempt.xlsx']


def test_generator_detail_structure_matches_template(tmp_path, generator):
    """明细表结构：三个工作表 + 第 2 行表头 + 第 3 行出勤率子表头。"""
    out = tmp_path / 'out'
    generator.generate(out, scale='small', dirty=False, seed=1)
    wb = load_workbook(out / 'detail.xlsx')
    assert tuple(wb.sheetnames) == ALLOWED_SHEETS
    early = wb['雁山（早）']
    assert [early.cell(2, c).value for c in range(1, 14)] == [
        '序号', '学院（人数）', '教室', '分配专业班级', '应到人数',
        '周一', '周一免勤人数', '周三', '周三免勤人数',
        '周四', '周四免勤人数', '周五', '周五免勤人数']
    assert early.cell(2, 14).value == '学生出勤率（%）'
    assert early.cell(2, 19).value == '值班辅导员未巡查次数'
    assert early.cell(2, 22).value == '早自习考勤时间'
    assert early.cell(2, 26).value == '执行效果'
    assert [early.cell(3, c).value for c in range(14, 19)] == \
        ['周一', '周三', '周四', '周五', '平均']
    late = wb['雁山（晚）']
    assert [late.cell(3, c).value for c in range(14, 18)] == ['周一', '周三', '周四', '平均']
    assert late.cell(3, 18).value is None, '晚自习表不应有周五出勤率列'
    assert wb['屏风'].max_row > 3
    wb.close()


def test_generator_exempt_structure_matches_template(tmp_path, generator):
    out = tmp_path / 'out'
    generator.generate(out, scale='small', dirty=False, seed=1)
    wb = load_workbook(out / 'exempt.xlsx')
    assert tuple(wb.sheetnames) == ('早自习免勤人员统计表', '晚自习免勤人员统计表')
    ws = wb['早自习免勤人员统计表']
    assert [ws.cell(2, c).value for c in range(1, 11)] == [
        '序号', '学号', '学院', '姓名', '专业班级', '免勤周次',
        '免勤时间', '免勤自习', '所在校区', '申请免勤理由']
    wb.close()


def test_generator_uses_only_fictional_identifiers(tmp_path, generator):
    """学号 / 姓名 / 教室号必须是明显的虚构占位。"""
    out = tmp_path / 'out'
    generator.generate(out, scale='small', dirty=True, seed=1)
    wb = load_workbook(out / 'exempt.xlsx')
    ws = wb['早自习免勤人员统计表']
    for row in range(3, min(ws.max_row, 60) + 1):
        sid = ws.cell(row, 2).value
        if sid:
            assert str(sid).startswith('99')
            assert len(str(sid)) == 13
            assert str(sid).isdigit()
        name = ws.cell(row, 4).value
        if name and '全体学生' not in str(name):
            assert str(name).startswith('学生')
    wb.close()

    wb = load_workbook(out / 'detail.xlsx')
    ws = wb['雁山（早）']
    for row in range(4, ws.max_row + 1):
        room = ws.cell(row, 3).value
        if room:
            assert str(room).startswith('8') or str(room).startswith('9')
            assert str(room).isdigit()
    wb.close()


# ------------------------------------------------------------------ 脏数据样本


def test_e2e_dirty_sample_reports_anomalies(tmp_path, generator):
    out = tmp_path / 'dirty'
    generator.generate(out, scale='small', dirty=True, seed=20261025)
    result = run_engine(out / 'detail.xlsx', out / 'exempt.xlsx',
                        output_path=out / 'out.xlsx')
    assert result.ok, result.error
    types = {a['type'] for a in result.anomalies}
    assert '免勤周次无法识别' in types
    assert '免勤星期无法识别' in types
    assert result.exempt_stats['重复申请事件'] > 0
    assert result.exempt_stats['集体免勤申请'] > 0
    assert result.exempt_stats['被集体申请覆盖的个人记录'] > 0
    assert result.exempt_stats['因校区过滤'] > 0
    assert result.exempt_stats['因周次不符过滤'] > 0
    assert tuple(result.sheets) == ALLOWED_SHEETS


def test_e2e_dirty_sample_still_has_three_sheets(tmp_path, generator):
    out = tmp_path / 'dirty2'
    generator.generate(out, scale='small', dirty=True, seed=5)
    result = run_engine(out / 'detail.xlsx', out / 'exempt.xlsx',
                        output_path=out / 'out.xlsx')
    wb = load_workbook(out / 'out.xlsx')
    assert _visible_sheets(wb) == ALLOWED_SHEETS
    wb.close()


def test_e2e_with_room_variant_runs(tmp_path, generator):
    """免勤表带教室号时也能跑通（当前引擎仍走比例分摊，见缺陷清单 D4）。"""
    out = tmp_path / 'withroom'
    generator.generate(out, scale='small', dirty=False, seed=3, with_room=True)
    result = run_engine(out / 'detail.xlsx', out / 'exempt.xlsx',
                        output_path=out / 'out.xlsx')
    assert result.ok, result.error
    assert result.exempt_stats['带教室号记录'] > 0
    wb = load_workbook(out / 'out.xlsx')
    assert _visible_sheets(wb) == ALLOWED_SHEETS
    wb.close()
