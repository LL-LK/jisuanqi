from contextlib import closing

from openpyxl import Workbook, load_workbook

from _helpers import make_exempt_row, write_exempt
from core import RunConfig, run


def test_summary_uses_reported_sessions_and_checks_identity(tmp_path):
    summary = tmp_path / '2026年第八周终稿.xlsx'
    with closing(Workbook()) as book:
        sheet = book.active
        sheet.append(['2026年第八周终稿'])
        sheet.append(['序号', '学院', '学生出勤率', '辅导员巡查率', '执行效果'])
        sheet.append([1, '土木工程学院', 0.92, 1, '人工记录'])
        sheet.append([2, '商学院', 0.90])
        sheet['A5'] = '备注：\n考勤时间：10月26日早晚自习，10月28日早自习\n学生出勤率按原表口径'
        sheet['K3'] = '其他列'
        book.save(summary)

    first = tmp_path / '免勤一.xlsx'
    second = tmp_path / '免勤二.xlsx'

    def row(last_digits, name, college='土木工程学院', **values):
        sid = '99000000' + last_digits if last_digits else ''
        return make_exempt_row(sid=sid, name=name, college=college,
                               week_text='第8周', **values)

    write_exempt(first, {
        '早自习': [
            row('00001', '测试甲', '土木学院'),
            row('00002', '测试甲'),
            row('00003', '测试乙'),
            row('00004', '测试丙', day_text='周四'),
        ],
        '晚自习': [row('00001', '测试甲', session='晚')],
    })
    write_exempt(second, {'早自习': [
        row('00001', '测 试甲'),
        row('00003', '测试丁'),
        row('', '测试甲'),
        row('00005', '测试戊', day_text='周三'),
        row('00006', '测试己', '不存在学院'),
    ]})

    target = tmp_path / '结果.xlsx'
    result = run(RunConfig(str(summary), output_path=str(target), week_no=8,
                           exempt_paths=(str(first), str(second))))
    assert result.ok, result.error
    assert result.exempt_stats['同学号身份冲突'] == 1
    assert result.exempt_stats['身份冲突事件（已按1人计）'] == 1
    assert result.exempt_stats['身份冲突暂缓计入人次'] == 0
    assert result.exempt_stats['学院待核对记录'] == 1
    assert result.exempt_stats['不在汇总考勤时段的记录'] == 1
    assert result.exempt_stats['重复申请事件'] == 2
    assert not result.college_rates
    with closing(load_workbook(target, data_only=True)) as book:
        sheet = book.worksheets[0]
        assert [sheet[cell].value for cell in ('C3', 'D3', 'E3', 'K3')] == [0.92, 1, '人工记录', '其他列']
        assert [sheet[cell].value for cell in ('F3', 'G3', 'H3', 'I3')] == [4, 1, 5, '待核对']
        assert {'学院免勤统计', '汇总与口径', '免勤核对'} <= set(book.sheetnames)
        review = list(book['免勤核对'].values)
        assert any('免勤表2' in source for _, source, _ in review[1:])
        assert '测试' not in str(review)
