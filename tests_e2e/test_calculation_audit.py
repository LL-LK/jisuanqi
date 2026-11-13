from contextlib import closing
from dataclasses import replace
from pathlib import Path

import pytest
from openpyxl import load_workbook

from core import RunConfig, run
from core.attendance_summary import load_summary_template
from core.export import TAMPER_META_SHEET, compare_actuals, verify_workbook_integrity
from core.io_detail import infer_workbook_week
from tests_e2e.helpers import create_detail_workbook, create_exempt_workbook


def test_summary_uses_complete_totals_and_keeps_patrol_blank(tmp_path):
    source = tmp_path / '（11.6早）桂林理工大学2026秋季学期第八周明细表.xlsx'
    college = '土木工程学院'
    missing = '地球科学学院'
    early = {'周一': 0, '周三': 50, '周四': 25, '周五': '59'}
    late = {'周一': 20, '周三': 30, '周四': 40}
    create_detail_workbook(
        source,
        early_blocks=[(college, 100, [('8101', 100, early)]),
                      (missing, 100, [('8102', 100, {**early, '周一': '请假2人'})])],
        late_blocks=[(college, 50, [('8103', 50, late)]),
                     (missing, 50, [('8104', 50, late)])],
        week_no=9,
    )
    with closing(load_workbook(source)) as book:
        book['雁山（晚）']['A1'] = '第八周'
        book.save(source)
    assert infer_workbook_week(source) == 9
    exempt = tmp_path / '免勤.xlsx'
    create_exempt_workbook(exempt, early_rows=[
        {'college': college, 'week_text': '第9周', 'room': '8101', 'day_text': '周三'},
    ])
    output = tmp_path / '明细结果.xlsx'
    config = RunConfig(str(source), exempt_paths=(str(exempt),), output_path=str(output),
                       week_no=9, enable_storage=False)
    before = source.read_bytes()
    result = run(config)
    assert result.ok, result.error
    assert source.read_bytes() == before
    assert compare_actuals(source, output) == []
    assert verify_workbook_integrity(output) == (True, [])
    assert verify_workbook_integrity(result.summary_output_path) == (True, [])

    template = load_summary_template(result.summary_output_path)
    assert len(template.scopes) == 7
    assert '11月6日早自习' in template.period
    with closing(load_workbook(result.summary_output_path, data_only=True)) as book:
        sheet = book['Sheet2']
        rates = {row[1].value: row for row in sheet.iter_rows(min_row=3, max_row=17)}
        assert rates[college][2].value == round(224 / (4 * 100 + 3 * 50 - 1), 4)
        assert rates[missing][2].value is None
        assert '待复核' in rates[missing][4].value
        assert all(row[3].value is None for row in rates.values())

    rejected = run(replace(config, output_path=str(tmp_path / 'blocked.xlsx'),
                           summary_output_path=str(source)))
    assert not rejected.ok and '受保护' in rejected.error
    assert not (tmp_path / 'blocked.xlsx').exists()
    assert source.read_bytes() == before


def test_missing_expected_count_does_not_become_zero(tmp_path, standard_detail_path,
                                                     standard_exempt_path):
    with closing(load_workbook(standard_detail_path)) as book:
        sheet = book['雁山（早）']
        sheet['B4'] = '计算机学院'
        sheet['E5'] = None
        book.save(standard_detail_path)
    result = run(RunConfig(str(standard_detail_path), exempt_path=str(standard_exempt_path),
                           output_path=str(tmp_path / 'output.xlsx'), enable_storage=False))
    assert result.ok, result.error
    record = next(rate for rate in result.college_rates
                  if rate.sheet == '雁山（早）' and rate.college == '计算机学院' and rate.day == '周一')
    assert record.rate is None
    assert record.status == '应到人数缺失，无法计算'


def test_result_audit_detects_edits_and_missing_signatures(tmp_path, standard_detail_path,
                                                           standard_exempt_path):
    output = tmp_path / '正常结果.xlsx'
    config = RunConfig(str(standard_detail_path), exempt_path=str(standard_exempt_path),
                       output_path=str(output), enable_storage=False)
    result = run(config)
    assert result.ok, result.error
    assert compare_actuals(standard_detail_path, output) == []

    changed = tmp_path / '改动结果.xlsx'
    with closing(load_workbook(output)) as book:
        for cell in ('F4', 'G4', 'N4'):
            book['雁山（早）'][cell] = 999
        book[TAMPER_META_SHEET].delete_rows(book[TAMPER_META_SHEET].max_row)
        book.save(changed)
    valid, issues = verify_workbook_integrity(changed)
    assert not valid
    assert all(any(f'雁山（早）!{cell}' in issue for issue in issues) for cell in ('F4', 'G4', 'N4'))
    assert any('清单与签名不符' in issue for issue in issues)
    assert compare_actuals(standard_detail_path, changed) == ['雁山（早）!F4']

    rejected = run(replace(config, detail_path=str(changed), output_path=str(tmp_path / 'blocked.xlsx')))
    assert not rejected.ok and '输入考勤表完整性校验失败' in rejected.error


def test_desktop_button_runs_the_source_pipeline(qapp, tmp_path, monkeypatch,
                                                 standard_detail_path, standard_exempt_path):
    from PySide6.QtCore import QEventLoop, QSettings, QTimer

    import app

    monkeypatch.setattr(app.MainWindow, '_settings', lambda self: QSettings(
        str(tmp_path / 'test.ini'), QSettings.Format.IniFormat))
    real_run = app.run
    monkeypatch.setattr(app, 'run', lambda config: real_run(replace(config, enable_storage=False)))

    window = app.MainWindow()
    window.detail_slot.set_path(str(standard_detail_path))
    window.exempt_slot.set_paths([str(standard_exempt_path)])
    window.out_edit.setText(str(tmp_path / 'desktop.xlsx'))
    loop = QEventLoop()
    window.run_btn.click()
    assert window.thread is not None
    window.thread.finished.connect(loop.quit)
    QTimer.singleShot(10000, loop.quit)
    loop.exec()
    assert window.result is not None and window.result.ok
    assert Path(window.result.output_path).is_file()
    assert Path(window.result.summary_output_path).is_file()
    window.deleteLater()


def test_switch_dictionary_loads_selected_rules_before_save(qapp, tmp_path, monkeypatch):
    from PySide6.QtCore import QSettings
    from PySide6.QtWidgets import QFileDialog

    import app
    from core.dictionary import Dictionary, default_tables, save_tables

    first, second = tmp_path / '词典甲.xlsx', tmp_path / '词典乙.xlsx'
    for target, alias in ((first, '甲别名'), (second, '乙别名')):
        tables = default_tables()
        tables['学院别名'].append(['土木工程学院', alias, '测试'])
        save_tables(target, tables)
    before = first.read_bytes()
    monkeypatch.setattr(app, 'default_dictionary_path', lambda: str(first))
    monkeypatch.setattr(app.MainWindow, '_settings', lambda self: QSettings(
        str(tmp_path / 'test.ini'), QSettings.Format.IniFormat))
    monkeypatch.setattr(QFileDialog, 'getOpenFileName', lambda *_: (str(second), 'Excel 文件'))

    window = app.MainWindow()
    assert any(row[1] == '甲别名' for row in window.dict_tables['学院别名'].rows())
    window._pick_dict()
    assert any(row[1] == '乙别名' for row in window.dict_tables['学院别名'].rows())
    window._save_dictionary()
    assert Dictionary.load(second).exact_college('乙别名') == '土木工程学院'
    assert Dictionary.load(second).exact_college('甲别名') == '甲别名'
    assert first.read_bytes() == before
    window.deleteLater()
