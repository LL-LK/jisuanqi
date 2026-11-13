from contextlib import closing
from datetime import datetime

import pytest
from openpyxl import load_workbook

from _helpers import build_focused_sample, make_exempt_row, write_exempt
from core import RunConfig, run
from core.dictionary import Dictionary, default_tables, save_tables
from core.exempt_report import export_exempt_report
from core.io_exempt import aggregate_exemptions, load_exemptions


def test_old_dictionary_inherits_defaults_and_keeps_custom_rules(tmp_path):
    path = tmp_path / '词典.xlsx'
    tables = default_tables()
    tables['学院别名'] = [['测试学院', '专用别名', '']]
    save_tables(path, tables)

    loaded = Dictionary.load(path)
    assert loaded.exact_college('体现与控制工程学院') == '机械与控制工程学院'
    assert loaded.exact_college('专用别名（80人）') == '测试学院'
    assert loaded.exact_college('测试学院与机械学院') == '测试学院与机械学院'

    original = path.read_bytes()
    tables['班级别名'] = [['测试2026-1班', '甲26-1班', ''],
                       ['测试26－1班', '乙26-1班', '']]
    with pytest.raises(ValueError, match='班级别名.*冲突'):
        save_tables(path, tables)
    assert path.read_bytes() == original

    with pytest.raises(ValueError, match='无法读取免勤词典'):
        Dictionary.load(tmp_path / '不存在.xlsx')


def test_date_and_session_pairing_does_not_multiply_exemptions(tmp_path):
    path = tmp_path / '免勤.xlsx'
    write_exempt(path, {'早自习免勤人员统计表': [
        make_exempt_row(day_text='周一早自习，周三晚自习', session='早晚自习'),
        make_exempt_row(sid='9900000000002', day_text=datetime(2026, 11, 5), session='晚'),
        make_exempt_row(sid='9900000000003', day_text='周一早\n周三晚', session=''),
        make_exempt_row(sid='9900000000004', day_text='周一早，周三，周四晚', session='早晚'),
    ]})
    rows = load_exemptions(path, week_no=7, source_number=1)
    result = aggregate_exemptions(rows, 7, identity_check=True)
    assert result.exempt_count('理学院', '周一', '早自习') == 2
    assert result.exempt_count('理学院', '周一', '晚自习') == 0
    assert result.exempt_count('理学院', '周三', '晚自习') == 2
    assert result.exempt_count('理学院', '周三', '早自习') == 0
    assert result.exempt_count('理学院', '周四', '晚自习') == 1
    assert any(issue['type'] == '自习时段无法识别' for issue in result.anomalies)


def test_multi_name_rows_use_verified_identities_and_class_aliases(tmp_path):
    tables = default_tables()
    tables['班级别名'] = [['试验班26-1班', '应用统计学26-1班', '']]
    tables['免勤理由分类'] = [['党课', '校内活动', '']]
    dictionary = Dictionary(tables)
    path = tmp_path / '免勤.xlsx'
    write_exempt(path, {'早自习免勤人员统计表': [
        make_exempt_row(sid='', name='测试甲、测试乙', class_text='试验班26-1班', reason='党课'),
        make_exempt_row(sid='9900000000001', name='测试甲', class_text='应用统计学2026-1班'),
        make_exempt_row(sid='', name='测试丙', class_text=''),
    ]})
    result = aggregate_exemptions(load_exemptions(path, dictionary=dictionary), 7, identity_check=True)
    assert result.exempt_count('理学院', '周一', '早自习') == 2
    assert result.name_matched_records == 1
    assert result.invalid_records == 1
    assert any(person['reason_category'] == '校内活动'
               for person in result.get_students('理学院', '周一', '早自习'))


def test_group_and_campus_ambiguity_is_reviewable_in_standalone_report(tmp_path):
    rows = [
        make_exempt_row(sid='', name='全体学生（共40人）', class_text=''),
        make_exempt_row(sid='9900000000001', class_text=''),
        make_exempt_row(sid='9900000000002', day_text='周三', campus=''),
        make_exempt_row(sid='9900000000003', day_text='周三', campus='雁山/屏风'),
        make_exempt_row(sid='', name='测试26-1班全体学生（共5人）',
                        class_text='测试26-1班', day_text='周四'),
        make_exempt_row(sid='', name='测试26-1班全体学生（共7人）',
                        class_text='测试26-1班', day_text='周四'),
    ]
    detail, exempt = build_focused_sample(tmp_path, early_rows=rows)
    tables = default_tables()
    next(row for row in tables['统计参数'] if row[0] == '空白校区默认值')[1] = '待确认'
    dictionary = tmp_path / '词典.xlsx'
    save_tables(dictionary, tables)

    result = run(RunConfig(str(detail), str(exempt), str(tmp_path / '结果.xlsx'),
                           dictionary_path=str(dictionary), enable_storage=False))
    assert result.ok, result.error
    early = {row['星期']: row for row in result.exempt_rows
             if row['学院'] == '理学院' and row['自习'] == '早自习'}
    assert [early[day]['免勤合计'] for day in ('周一', '周三', '周四')] == [1, 0, 7]
    assert all(early[day]['免勤状态'] == '规则待核对' for day in ('周一', '周三', '周四'))
    report = tmp_path / '报告.xlsx'
    export_exempt_report(report, result)
    with closing(load_workbook(report, data_only=True)) as book:
        issues = list(book['免勤核对'].values)[1:]
    assert {kind for kind, _, _ in issues} >= {'集体申请缺班级', '校区待核对', '集体免勤重复申报'}
    assert all('免勤表1' in ref for _, ref, _ in issues)

    rejected = run(RunConfig(str(detail), str(exempt), str(tmp_path / '错误范围.xlsx'),
                             campus='屏风', enable_storage=False))
    assert not rejected.ok and '只计算雁山校区' in rejected.error


def test_missing_college_requires_unique_explicit_class_evidence(tmp_path):
    path = tmp_path / '免勤.xlsx'
    write_exempt(path, {'早自习免勤人员统计表': [
        make_exempt_row(sid='9900000000001', college='理学院', class_text='测试2026-1班'),
        make_exempt_row(sid='9900000000002', college='', class_text='测试26-1班'),
        make_exempt_row(sid='9900000000003', college='', class_text='产品设计26-1班'),
        make_exempt_row(sid='9900000000004', college='理学院', class_text='测试26-2班'),
        make_exempt_row(sid='9900000000005', college='艺术学院', class_text='测试26-2班'),
        make_exempt_row(sid='9900000000006', college='', class_text='测试26-2班'),
    ]})
    result = aggregate_exemptions(load_exemptions(path, source_number=1, keep_unassigned=True),
                                  7, identity_check=True, allowed_colleges={'理学院', '艺术学院'})
    assert result.exempt_count('理学院', '周一', '早自习') == 3
    assert result.exempt_count('艺术学院', '周一', '早自习') == 1
    assert result.college_matched_by_class == 1
    assert result.unknown_college_records == 2
