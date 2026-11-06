"""按学院回填、多表去重与第八周模板结构的回归测试。"""

from contextlib import closing

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

from _helpers import make_exempt_row, md5, write_exempt
from core import RunConfig, run
from core.exempt_report import export_exempt_report
from core.io_detail import load_detail, week_from_label


def make_template(path):
    with closing(Workbook()) as wb:
        sheet = wb.active
        sheet.title = '雁山（早）'
        sheet.append(['第九周'])
        sheet.append(['序号', '学院（人数）', '教室', '分配专业班级', '应到人数',
                      '周一', '周一免勤人数', '周三', '周三免勤人数', '学生出勤率（%）',
                      None, None, '辅导员巡查', '考察时间', '巡查效果'])
        sheet.append([1, '土木工程学院（100人）', '8101', '', 30, 15, None, 25])
        sheet.append([None, None, '8102', '', 35, 19, None, 25])
        sheet.append([None, None, '8103', '', 35, 20, None, 20])
        sheet.append([2, '理学院（60人）', '8201', '', 60, 30])
        for area in ('J2:L2', 'B3:B5', 'G3:G5', 'I3:I5', 'J4:J5', 'K4:K5', 'L4:L5', 'M3:M5'):
            sheet.merge_cells(area)
        sheet['M3'] = '=1+2'
        sheet['N4'] = '人工记录'
        sheet['O4'] = '已巡查'
        sheet['O4'].font = Font(bold=True)
        wb.create_sheet('屏风')['A1'] = '原始内容'
        wb.save(path)


def test_college_template_and_multiple_exemption_files(tmp_path):
    detail = tmp_path / '第八周明细表.xlsx'
    make_template(detail)
    first = tmp_path / '一批免勤.xlsx'
    second = tmp_path / '二批免勤.xlsx'
    out = tmp_path / '结果.xlsx'
    early = '早自习免勤人员统计表'

    def entry(sid, college, **fields):
        return make_exempt_row(sid=sid, college=college, week_text='第8周', **fields)

    def group(college, size):
        return entry('', college, name=f'试验26-1班全体学生（共{size}人）',
                     class_text='试验26-1班')

    write_exempt(first, {early: [
        entry('9900000000001', '土木学院', day_text='周一、周三'),
        entry('9900000000002', '土木学院', class_text='试验26-1班'),
        group('土木学院', 38),
    ]})
    write_exempt(second, {early: [
        entry('9900000000001', '土木工程学院'),
        group('土木工程学院', 40),
    ]})
    inputs = {path: md5(path) for path in (detail, first, second)}
    result = run(RunConfig(str(detail), output_path=str(out), week_no=8,
                           exempt_paths=(str(first), str(second), str(first))))
    assert result.ok, result.error
    assert result.college_exempt and not result.exempt_overflow
    assert result.summary['导入免勤表数'] == 2
    assert result.exempt_stats['重复申请事件'] == 2
    assert any(item['type'] == '模板标题周次不一致' for item in result.anomalies)
    assert week_from_label(detail.name) == 8
    assert result.summary['写出免勤单元格'] == 4
    expected_week = (54 / 59 + 70 / 99) / 2
    weekly = {row.college: row for row in result.weekly_rates}
    assert weekly['土木工程学院'].rate == pytest.approx(expected_week)
    assert weekly['理学院'].rate is None
    assert weekly['理学院'].calculated_days == 1
    with closing(load_workbook(out)) as wb:
        assert wb.sheetnames == ['雁山（早）', '屏风']
        sheet = wb['雁山（早）']
        assert [sheet.cell(3, col).value for col in range(10, 13)] == ['周一', '周三', '周平均']
        assert sheet['G3'].value == 41
        assert sheet['I3'].value == 1
        assert sheet['G4'].value is None
        assert sheet['J4'].value == pytest.approx(54 / 59)
        assert sheet['L4'].value == pytest.approx(expected_week)
        assert sheet['L6'].value is None
        assert [sheet[k].value for k in ('M3', 'N4', 'O4')] == ['=1+2', '人工记录', '已巡查']
        assert sheet['O4'].font.bold
        assert 'M3:M5' in sheet.merged_cells
        assert wb['屏风']['A1'].value == '原始内容'
    assert load_detail(out).rate_label_rows == {'雁山（早）': 3}
    row = next(row for row in result.exempt_rows
               if row['学院'] == '土木工程学院' and row['星期'] == '周一')
    assert (row['个人免勤'], row['集体免勤'], row['已回填'], row['回填差额']) == (1, 40, 41, 0)

    report = tmp_path / '学院免勤统计.xlsx'
    export_exempt_report(report, result)
    with closing(load_workbook(report, data_only=True)) as wb:
        assert '学院免勤统计' in wb
        assert '容量超额' not in [cell.value for cell in wb['学院免勤统计'][1]]
    rejected = run(RunConfig(str(detail), output_path=str(second), week_no=8,
                             exempt_paths=(str(first), str(second))))
    assert not rejected.ok and '不能与任一输入免勤表相同' in rejected.error
    assert {path: md5(path) for path in inputs} == inputs
