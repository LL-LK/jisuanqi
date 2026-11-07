"""免勤分项、回填差额及独立工作簿的端到端回归。"""

import pytest
from openpyxl import load_workbook

from _helpers import build_focused_sample, make_exempt_row, md5, run_engine
from core.exempt_report import REPORT_COLUMNS, export_exempt_report


def test_exempt_breakdown_reports_only_once_and_keeps_unwritten_scopes(tmp_path):
    early = [
        make_exempt_row(sid='', name='测试班26-1班全体学生（共45人）',
                        class_text='测试班26-1班', room='8101'),
        make_exempt_row(sid='9900000000001', class_text='测试班26-1班'),
        make_exempt_row(sid='9900000000002', room='08101'),
        make_exempt_row(sid='9900000000003', room='8101'),
        make_exempt_row(sid='9900000000004', day_text='周二'),
        make_exempt_row(sid='全班', name='测试班26-3班全体学生',
                        class_text='测试班26-3班', day_text='周二'),
    ]
    late = [make_exempt_row(sid='9900000000005', day_text='周五', session='晚')]
    detail, exempt = build_focused_sample(tmp_path, early, late, with_room=True)
    out = tmp_path / '回填结果.xlsx'
    result = run_engine(detail, exempt, out)
    assert result.ok, result.error
    rows = {(row['星期'], row['自习']): row for row in result.exempt_rows}

    monday = rows['周一', '早自习']
    assert [monday[key] for key in REPORT_COLUMNS[3:9]] == [2, 45, 47, 40, 7, 0]
    assert monday['回填状态'] == '有容量超额'
    tuesday = rows['周二', '早自习']
    assert [tuesday[key] for key in REPORT_COLUMNS[3:9]] == [1, 0, 1, 0, 0, 1]
    assert tuesday['免勤状态'] == '待补人数'
    assert tuesday['模板出勤率列'] == '无'
    friday = rows['周五', '晚自习']
    assert friday['已回填'] == 1
    assert friday['模板出勤率列'] == '无'

    report = tmp_path / '免勤统计.xlsx'
    export_exempt_report(report, result, [monday, tuesday])
    book = load_workbook(report, data_only=True)
    try:
        assert tuple(book['免勤统计'].values) == (
            REPORT_COLUMNS, tuple(monday[key] for key in REPORT_COLUMNS),
            tuple(tuesday[key] for key in REPORT_COLUMNS))
        summary = dict(book['汇总与口径'].values)
        assert summary['免勤合计（人次）'] == 48
        assert summary['未落表（人次）'] == 1
    finally:
        book.close()
    book = load_workbook(out, data_only=True)
    try:
        assert [s for s in book.sheetnames if s != '_系统校验元数据'] == ['雁山（早）', '雁山（晚）', '屏风']
        assert book['雁山（早）']['G4'].value == 40
    finally:
        book.close()

    for protected in (detail, exempt, out):
        original = md5(protected)
        with pytest.raises(ValueError, match='受保护'):
            export_exempt_report(protected, result)
        assert md5(protected) == original
