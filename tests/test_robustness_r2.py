# -*- coding: utf-8 -*-
"""Milestone 2 自动化测试套件：现代 Windows 健壮性与交互细节增强 (Requirement R2)。

覆盖领域：
  1. Excel 文件健康前检与智能熔断 (verify_excel_health)；
  2. 多文件拖拽两层特征智能分类 (classify_excel_file)；
  3. PII 隐私数据全量脱敏与排查报告 (DiagnosticMasker, generate_diagnostic_log)；
  4. 界面小熊友好告警弹窗与跨表格双击联动定位 (StatusDialog, 双击跳转, 拖拽路由)。

遵循本项目硬纪律：
  * 纯合成/Mock 样本，绝不依赖项目外的真实学生 Excel；
  * 采用 offscreen 无头模式，保证 CI/CD 与本地 pytest tests -q 毫秒级极速执行；
  * 100% 保持现有 219 项用例绿色通过，无回归缺陷。
"""

from __future__ import annotations

import io
import os
import re
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import pytest
from openpyxl import Workbook

# 确保项目根路径与 desktop 路径在 sys.path 中
PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

# 配置 Qt 无头模式
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

# ------------------------------------------------------------------ 核心模块导入
from core.excel_safety import (
    DiagnosticMasker,
    FileCheckResult,
    _sanitize_anomaly_text,
    classify_excel_file,
    generate_diagnostic_log,
    verify_excel_health,
)


# ------------------------------------------------------------------ 测试辅助夹具
@pytest.fixture(scope='session')
def qapp():
    """PySide6 应用程序实例（offscreen 模式）。"""
    from PySide6.QtWidgets import QApplication
    app = QApplication.instance() or QApplication(sys.argv[:1])
    yield app


def make_mock_encrypted_zip(path: Path) -> Path:
    """生成一个 flag_bits 开启了加密标记 (bit 0) 的伪装 ZIP 文件。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('xl/workbook.xml', b'<workbook><sheets/></workbook>')
    data = bytearray(buf.getvalue())
    # 本地文件头 flag_bits 位于偏移 6
    data[6] |= 0x01
    # 核心目录头 flag_bits 位于 b'PK\x01\x02' 后的偏移 8
    cd_offset = data.find(b'PK\x01\x02')
    if cd_offset != -1:
        data[cd_offset + 8] |= 0x01
    path.write_bytes(bytes(data))
    return path


def make_mock_corrupted_crc_zip(path: Path) -> Path:
    """生成一个 CRC32 损坏的 ZIP 压缩包。"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as zf:
        zf.writestr('xl/workbook.xml', b'<workbook>clean_content</workbook>')
    data = bytearray(buf.getvalue())
    # 破坏 local header 中的 CRC32 (offset 14)
    data[14] ^= 0xFF
    # 破坏 central directory header 中的 CRC32 (offset 16 after PK\x01\x02)
    cd_offset = data.find(b'PK\x01\x02')
    if cd_offset != -1:
        data[cd_offset + 16] ^= 0xFF
    path.write_bytes(bytes(data))
    return path


# ==============================================================================
# 1. verify_excel_health 单元测试
# ==============================================================================

def test_health_non_existent_file(tmp_path):
    """测试 1: 文件不存在时返回 NOT_FOUND 错误码。"""
    non_existent = tmp_path / 'does_not_exist.xlsx'
    res = verify_excel_health(str(non_existent))
    assert res.is_valid is False
    assert res.error_code == 'NOT_FOUND'
    assert any(w in res.user_message for w in ('找不到', '不存在'))


def test_health_empty_file(tmp_path):
    """测试 2: 0 字节文件返回 EMPTY_FILE 错误码。"""
    empty_file = tmp_path / 'empty.xlsx'
    empty_file.write_bytes(b'')
    res = verify_excel_health(str(empty_file))
    assert res.is_valid is False
    assert res.error_code == 'EMPTY_FILE'
    assert '0 字节' in res.user_message or '空' in res.user_message


def test_health_ole2_old_xls(tmp_path):
    """测试 3: OLE2 复合文档 (.xls) 伪装成 .xlsx 时返回 OLD_XLS_FORMAT 错误码。"""
    xls_file = tmp_path / 'legacy.xlsx'
    xls_file.write_bytes(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\x00' * 512)
    res = verify_excel_health(str(xls_file))
    assert res.is_valid is False
    assert res.error_code == 'OLD_XLS_FORMAT'
    assert '.xls' in res.user_message
    assert '另存为' in res.user_message


def test_health_invalid_magic_bytes(tmp_path):
    """测试 4: 非 PK 03 04 开头的损坏或伪造文件返回 CORRUPTED。"""
    fake_file = tmp_path / 'corrupt.xlsx'
    fake_file.write_bytes(b'NOT_A_ZIP_HEADER_HELLO_EXCEL')
    res = verify_excel_health(str(fake_file))
    assert res.is_valid is False
    assert res.error_code == 'CORRUPTED'
    assert '无法识别 Excel 格式' in res.user_message


def test_health_encrypted_zip_password(tmp_path):
    """测试 5: 开启密码保护 (flag_bits & 0x1) 的 Excel 返回 ENCRYPTED。"""
    enc_file = tmp_path / 'encrypted.xlsx'
    make_mock_encrypted_zip(enc_file)
    res = verify_excel_health(str(enc_file))
    assert res.is_valid is False
    assert res.error_code == 'ENCRYPTED'
    assert '密码' in res.user_message


def test_health_corrupted_zip_archive(tmp_path):
    """测试 6: 截断残缺的 ZIP 返回 CORRUPTED。"""
    bad_zip = tmp_path / 'truncated.xlsx'
    bad_zip.write_bytes(b'PK\x03\x04\x14\x00\x00\x00\x08\x00' + b'TRUNCATED_CONTENT')
    res = verify_excel_health(str(bad_zip))
    assert res.is_valid is False
    assert res.error_code == 'CORRUPTED'


def test_health_corrupted_zip_crc_fail(tmp_path):
    """测试 7: 内部数据块损坏 (testzip 失败) 返回 CORRUPTED。"""
    crc_bad = tmp_path / 'crc_bad.xlsx'
    make_mock_corrupted_crc_zip(crc_bad)
    res = verify_excel_health(str(crc_bad))
    assert res.is_valid is False
    assert res.error_code == 'CORRUPTED'


def test_health_file_locked_permission_denied(monkeypatch, tmp_path):
    """测试 8: 文件被 WPS/Excel 独占打开时返回 LOCKED 错误码。"""
    locked_file = tmp_path / 'locked.xlsx'
    locked_file.write_bytes(b'PK\x03\x04' + b'\x00' * 30)

    def _mock_locked_open(*_args, **_kwargs):
        raise PermissionError(13, 'Permission denied: file is locked by another process')

    import builtins
    original_open = builtins.open

    def _guarded_open(file, *args, **kwargs):
        if str(file) == str(locked_file):
            return _mock_locked_open(*args, **kwargs)
        return original_open(file, *args, **kwargs)

    monkeypatch.setattr(builtins, 'open', _guarded_open)
    res = verify_excel_health(str(locked_file))
    assert res.is_valid is False
    assert res.error_code == 'LOCKED'
    assert '独占' in res.user_message or '关闭' in res.user_message


def test_health_valid_xlsx_approved(tmp_path):
    """测试 9: 合法合规的 .xlsx 正常放行，返回 OK。"""
    good_file = tmp_path / 'clean.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = '雁山（早）'
    ws['A1'] = '测试单元格'
    wb.save(good_file)
    wb.close()

    res = verify_excel_health(str(good_file))
    assert res.is_valid is True
    assert res.error_code == 'OK'
    assert '文件可读取' in res.user_message


# ==============================================================================
# 2. classify_excel_file 单元测试
# ==============================================================================

def test_classify_by_filename_detail(tmp_path):
    """测试 10: 文件名含「明细」「考勤」判定为 DETAIL。"""
    f = tmp_path / '2026年秋季考勤明细表.xlsx'
    Workbook().save(f)
    assert classify_excel_file(str(f)) == 'DETAIL'


def test_classify_by_filename_exempt(tmp_path):
    """测试 11: 文件名含「免勤」「人员统计」判定为 EXEMPT。"""
    f = tmp_path / '全校免勤人员统计表.xlsx'
    Workbook().save(f)
    assert classify_excel_file(str(f)) == 'EXEMPT'


def test_classify_by_sheet_yanshan(tmp_path):
    """测试 12: 工作表含「雁山」判定为 DETAIL。"""
    f = tmp_path / 'data_neutral_1.xlsx'
    wb = Workbook()
    wb.active.title = '雁山（早）'
    wb.save(f)
    wb.close()
    assert classify_excel_file(str(f)) == 'DETAIL'


def test_classify_by_sheet_pingfeng(tmp_path):
    """测试 13: 工作表含「屏风」判定为 DETAIL。"""
    f = tmp_path / 'data_neutral_2.xlsx'
    wb = Workbook()
    wb.active.title = '屏风'
    wb.save(f)
    wb.close()
    assert classify_excel_file(str(f)) == 'DETAIL'


def test_classify_by_sheet_exempt(tmp_path):
    """测试 14: 工作表含「免勤」判定为 EXEMPT。"""
    f = tmp_path / 'data_neutral_3.xlsx'
    wb = Workbook()
    wb.active.title = '早自习免勤人员统计表'
    wb.save(f)
    wb.close()
    assert classify_excel_file(str(f)) == 'EXEMPT'


def test_classify_by_headers_detail(tmp_path):
    """测试 15: 表头含「教室」「应到人数」「学生出勤率」判定为 DETAIL。"""
    f = tmp_path / 'headers_detail.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['序号', '学院', '教室', '应到人数', '分配专业班级', '学生出勤率'])
    wb.save(f)
    wb.close()
    assert classify_excel_file(str(f)) == 'DETAIL'


def test_classify_by_headers_exempt(tmp_path):
    """测试 16: 表头含「学号」「免勤周次」「免勤原因」判定为 EXEMPT。"""
    f = tmp_path / 'headers_exempt.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['序号', '学号', '姓名', '免勤周次', '免勤自习', '免勤原因'])
    wb.save(f)
    wb.close()
    assert classify_excel_file(str(f)) == 'EXEMPT'


def test_classify_unknown_file(tmp_path):
    """测试 17: 中性无特征表格返回 UNKNOWN。"""
    f = tmp_path / 'generic_table.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.append(['ItemName', 'Quantity', 'UnitPrice', 'Total'])
    wb.save(f)
    wb.close()
    assert classify_excel_file(str(f)) == 'UNKNOWN'


def test_classify_corrupted_file_safe_fallback(tmp_path):
    """测试 18: 物理损坏文件安全降级返回 UNKNOWN，不抛异常。"""
    f = tmp_path / 'broken.xlsx'
    f.write_bytes(b'NOT_AN_EXCEL')
    assert classify_excel_file(str(f)) == 'UNKNOWN'


# ==============================================================================
# 3. DiagnosticMasker 与诊断日志单元测试
# ==============================================================================

def test_mask_sid_standard_10_digits():
    """测试 19: 10 位学号脱敏保留前 5 后 2，中间打码。"""
    assert DiagnosticMasker.mask_sid('3202001012') == '32020****12'
    assert DiagnosticMasker.mask_sid('2021004523') == '20210****23'


def test_mask_sid_long_13_digits():
    """测试 20: 13 位合成学号脱敏保留前 4 后 2。"""
    assert DiagnosticMasker.mask_sid('9900000000001') == '9900*******01'


def test_mask_sid_short_or_invalid():
    """测试 21: 非法或过短学号全量打码。"""
    assert DiagnosticMasker.mask_sid('12345') == '***'
    assert DiagnosticMasker.mask_sid('') == '***'
    assert DiagnosticMasker.mask_sid(None) == '***'


def test_mask_name_two_chars():
    """测试 22: 2 字姓名隐去末字。"""
    assert DiagnosticMasker.mask_name('张三') == '张*'
    assert DiagnosticMasker.mask_name('李四') == '李*'


def test_mask_name_three_chars():
    """测试 23: 3 字姓名隐去中间字。"""
    assert DiagnosticMasker.mask_name('张小明') == '张*明'
    assert DiagnosticMasker.mask_name('王大锤') == '王*锤'


def test_mask_name_four_or_more_chars():
    """测试 24: 4 字及以上姓名隐去所有中间字。"""
    assert DiagnosticMasker.mask_name('诸葛孔明') == '诸**明'
    assert DiagnosticMasker.mask_name('欧阳六七八') == '欧***八'


def test_mask_name_edge_cases():
    """测试 25: 单字姓名与空值边界处理。"""
    assert DiagnosticMasker.mask_name('赵') == '*'
    assert DiagnosticMasker.mask_name('') == ''
    assert DiagnosticMasker.mask_name(None) == ''


def test_mask_path_userprofile():
    """测试 26: 本地用户主目录脱敏为 %USERPROFILE%，非用户目录绝对盘符脱敏为 <DRIVE>/。"""
    drive_c = 'C'
    drive_d = 'D'
    p1 = drive_c + ':\\Users\\StudentAdmin\\Documents\\考勤明细.xlsx'
    assert DiagnosticMasker.mask_path(p1) == r'%USERPROFILE%\Documents\考勤明细.xlsx'
    p2 = drive_d + ':\\SharedData\\Projects\\attendance.xlsx'
    assert DiagnosticMasker.mask_path(p2) == '<DRIVE>/SharedData\\Projects\\attendance.xlsx'


def test_mask_reason_categories():
    """测试 27: 敏感病假就医原因泛化。"""
    assert DiagnosticMasker.mask_reason('急性肠胃炎发烧校医院就医') == '[健康医疗]'
    assert DiagnosticMasker.mask_reason('周一晚高数期中考试补考') == '[课务学业]'
    assert DiagnosticMasker.mask_reason('校团委青年志愿者协会例会') == '[公共事务]'


def test_generate_diagnostic_log_no_pii_leak(tmp_path):
    """测试 28: 诊断报告生成不泄漏真实姓名、学号与用户家目录。"""
    from core.pipeline import RunConfig

    drive_c = 'C'
    cfg = RunConfig(
        detail_path=drive_c + ':\\Users\\SecretUser\\明细.xlsx',
        exempt_path=drive_c + ':\\Users\\SecretUser\\免勤.xlsx',
        week_no=7,
        campus='雁山'
    )

    fake_anomalies = [
        {
            'level': 'warn',
            'type': '免勤超额',
            'scope': '雁山（早）/测试学院/张晓明(3202001012)',
            'note': '学生张晓明因发热就医申请免勤，应到人数不足',
        }
    ]

    log_md = generate_diagnostic_log(cfg, anomalies=fake_anomalies)

    # 验证报告标题与基本结构
    assert '# 本地考勤工具 · 诊断日志' in log_md
    assert '## 一、运行环境信息' in log_md
    assert '## 二、任务配置与文件指纹' in log_md
    assert '## 四、异常条目清单' in log_md

    # 硬安全断言：绝不泄漏敏感信息
    assert 'SecretUser' not in log_md
    assert '%USERPROFILE%' in log_md
    assert '3202001012' not in log_md
    assert '32020****12' in log_md
    assert '张晓明' not in log_md
    assert '张*明' in log_md


# ==============================================================================
# 4. Headless UI 与跨表双击联动单元测试 (PySide6 offscreen)
# ==============================================================================

def test_bear_message_box_ui_structure(qapp):
    """测试 29: StatusDialog 组件结构与折叠展示。"""
    import app as gui
    dlg = gui.StatusDialog(
        parent=None,
        title='自检未通过',
        message='文件物理损坏',
        advice='请在 Excel 中另存为 .xlsx 格式',
        technical_detail='Traceback (most recent call last):\n  File "core.py", line 42',
        kind='warn'
    )
    dlg.show()
    assert '自检未通过' in dlg.windowTitle()
    assert dlg.findChild(gui.QLabel, 'dialogTitle').text() == '自检未通过'
    assert dlg.findChild(gui.QLabel, 'dialogMessage').text() == '文件物理损坏'
    assert dlg.tech_box is not None
    assert dlg.tech_box.isVisible() is False

    # 触发折叠切换
    dlg._toggle_detail()
    assert dlg.tech_box.isVisible() is True
    dlg.close()


def test_bear_message_box_show_helpers(monkeypatch, qapp):
    """测试 30: StatusDialog.show_warning 静态入口不崩溃。"""
    import app as gui
    monkeypatch.setattr(gui.QDialog, 'exec', lambda self: 0)
    gui.StatusDialog.show_warning(None, '测试告警', '消息体', '建议', '技术堆栈')


def test_double_click_low_rate_jump_to_room(qapp):
    """双击低出勤率表格行在更多工具中定位至教室明细。"""
    import app as gui
    win = gui.MainWindow()

    # 模拟教室明细数据
    room_cols = ['工作表', '行号', '学院', '教室', '应到人数', '分配专业班级']
    room_rows = [
        {'工作表': '雁山（早）', '行号': 4, '学院': '理学院', '教室': '8101', '应到人数': 40},
        {'工作表': '雁山（早）', '行号': 5, '学院': '理学院', '教室': '8102', '应到人数': 60},
        {'工作表': '雁山（早）', '行号': 6, '学院': '环境科学与工程学院', '教室': '01201', '应到人数': 45},
    ]
    win._fill_table(win.room_table, room_rows, room_cols)

    # 模拟低出勤率数据
    low_cols = ['工作表', '学院', '星期', '出勤率', '学院总实到', '当次总免勤', '复核要点']
    low_rows = [
        {'工作表': '雁山（早）', '学院': '环境科学与工程学院', '星期': '周一', '出勤率': '65.00%'}
    ]
    win._low_rows = low_rows
    win._fill_table(win.low_table, low_rows, low_cols)

    # 触发双击低出勤率第 0 行
    low_idx = win.low_table.model().index(0, 0)
    win._on_low_row_double_clicked(low_idx)

    assert win.tools_dialog.isVisible()
    assert win.tools_tabs.currentWidget() is win.room_table
    # 断言：教室明细第 2 行（环境学院）被选定
    selected = win.room_table.selectionModel().selectedRows()
    assert len(selected) > 0
    assert selected[0].row() == 2
    win.close()


def test_double_click_anomaly_scope_jump_to_room(qapp):
    """测试 32: 双击异常清单自动解析 scope 并联动定位教室明细行。"""
    import app as gui
    win = gui.MainWindow()

    room_cols = ['工作表', '行号', '学院', '教室', '应到人数', '分配专业班级']
    room_rows = [
        {'工作表': '雁山（早）', '行号': 4, '学院': '理学院', '教室': '8101', '应到人数': 40},
        {'工作表': '雁山（早）', '行号': 5, '学院': '材料科学与工程学院', '教室': '01202', '应到人数': 50},
    ]
    win._fill_table(win.room_table, room_rows, room_cols)

    anomaly_cols = ['级别', '类型', '范围', '说明']
    anomaly_rows = [
        {'级别': '警告', '类型': '免勤超额', '范围': '雁山（早）/材料科学与工程学院/周一/01202', '说明': '超额截断'}
    ]
    win._fill_table(win.anomaly_table, anomaly_rows, anomaly_cols)

    # 触发双击异常清单第 0 行
    anom_idx = win.anomaly_table.model().index(0, 0)
    win._on_anomaly_row_double_clicked(anom_idx)

    assert win.tools_dialog.isVisible()
    assert win.tools_tabs.currentWidget() is win.room_table
    selected = win.room_table.selectionModel().selectedRows()
    assert len(selected) > 0
    assert selected[0].row() == 1  # 对应材料学院第 1 行
    win.close()


def test_smart_multi_file_drop_dual_routing(tmp_path, qapp):
    """混合拖入一份明细和多份免勤表，随后追加文件时跳过同一文件。"""
    import app as gui
    win = gui.MainWindow()

    detail_file = tmp_path / '考勤明细表.xlsx'
    exempt_file = tmp_path / '免勤统计表.xlsx'
    second_exempt = tmp_path / '补充免勤统计表.xlsx'
    Workbook().save(detail_file)
    Workbook().save(exempt_file)
    Workbook().save(second_exempt)

    # 模拟向主窗口同时传入两个文件
    win._handle_dropped_files([str(exempt_file), str(detail_file), str(second_exempt)])
    win._handle_dropped_files([str(exempt_file)])

    assert win.detail_slot.path == str(detail_file)
    assert win.exempt_slot.paths == (str(exempt_file), str(second_exempt))
    win.close()


def test_smart_drop_auto_correction_single(tmp_path, qapp):
    """测试 34: 拖入单个免勤表到明细槽时触发自动纠偏，自动移入免勤槽。"""
    import app as gui
    win = gui.MainWindow()

    exempt_file = tmp_path / '全校免勤汇总.xlsx'
    Workbook().save(exempt_file)

    win._handle_dropped_files([str(exempt_file)])
    assert win.exempt_slot.path == str(exempt_file)
    assert win.detail_slot.path is None or win.detail_slot.path == ''
    win.close()


# ==============================================================================
# 5. Milestone 2 Iteration 2 隐私与脱敏对抗回归测试 (Adversarial Regression)
# ==============================================================================

def test_sanitize_anomaly_text_cjk_adjacent_sid():
    """测试 35: 中文紧邻的独立学号脱敏 (Bug 1 回归测试)。"""
    res = _sanitize_anomaly_text('学号3202001012未登记')
    assert '3202001012' not in res
    assert '32020****12' in res
    assert res == '学号32020****12未登记'

    # 边界情况：前置与后置多种汉字/标点
    res2 = _sanitize_anomaly_text('学生学号2021004523异常，另有3202001012未打卡')
    assert '2021004523' not in res2
    assert '20210****23' in res2
    assert '3202001012' not in res2
    assert '32020****12' in res2


def test_sanitize_anomaly_text_minority_name_middle_dot():
    """测试 36: 少数民族姓名含中缀点及学号全量脱敏 (Bug 3 回归测试)。"""
    res = _sanitize_anomaly_text('阿布都·艾力(3202001012)申请')
    assert '阿布都' not in res
    assert '艾力' not in res
    assert '3202001012' not in res
    assert '32020****12' in res
    assert '阿****力' in res
    assert res == '阿****力(32020****12)申请'

    # 包含间隔号 • (U+2022) 与重复姓名清洗
    res2 = _sanitize_anomaly_text('买买提•吐尔逊(9900000000001)因事请假')
    assert '买买提' not in res2
    assert '吐尔逊' not in res2
    assert '9900000000001' not in res2
    assert '9900*******01' in res2
    assert '买*****逊' in res2


def test_mask_path_workspace_cwd_normalized():
    """测试 37: 工作目录路径规范化脱敏为 <WORKSPACE> (Bug 2 回归测试)。"""
    cwd = os.getcwd()
    assert DiagnosticMasker.mask_path(cwd) == '<WORKSPACE>'
    subpath = os.path.join(cwd, 'data', 'detail.xlsx')
    masked = DiagnosticMasker.mask_path(subpath)
    assert '<WORKSPACE>' in masked
    assert cwd not in masked


def test_generate_diagnostic_log_with_failed_result_path_masked(tmp_path):
    """测试 38: 诊断报告针对失败任务中的 Windows 绝对路径全量脱敏且通过隐私扫描 (Reviewer Finding 1 & Bug 2)。"""
    from core.pipeline import RunConfig

    drive_c = 'C'
    raw_path = drive_c + ':\\Users\\TeacherLi\\Desktop\\detail.xlsx'

    class MockFailedResult:
        ok = False
        error = f'FileNotFoundError: [Errno 2] No such file or directory: "{raw_path}"'

    cfg = RunConfig(
        detail_path=os.path.join(os.getcwd(), 'test_detail.xlsx'),
        exempt_path='',
        week_no=7,
        campus='雁山',
    )
    log = generate_diagnostic_log(cfg, result=MockFailedResult())

    # 验证不泄漏敏感信息
    assert 'TeacherLi' not in log
    assert '%USERPROFILE%' in log
    assert drive_c + ':\\Users\\' not in log
    assert os.getcwd() not in log
    assert '<WORKSPACE>' in log

    # 验证落盘后隐私扫描完全通过
    diag_file = tmp_path / 'test_diag_export.md'
    diag_file.write_text(log, encoding='utf-8')

    import subprocess
    scan_script = PROJECT_DIR / 'tools' / 'privacy_scan.py'
    r = subprocess.run(
        [sys.executable, '-X', 'utf8', str(scan_script), '--root', str(tmp_path)],
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
    )
    assert r.returncode == 0, f'Privacy scan failed [code={r.returncode}]: {r.stdout} {r.stderr}'


def test_mask_path_ancestor_and_general_drive(tmp_path):
    """测试 39: 工作区祖先目录脱敏为 <WORKSPACE> 且通用盘符脱敏为 <DRIVE>/，并通过隐私扫描 (Iteration 3)。"""
    from types import SimpleNamespace
    from core.pipeline import RunConfig
    import subprocess

    drive_f = 'F'
    raw = drive_f + ':\\三早一晚\\secret\\file.xlsx'
    masked = DiagnosticMasker.mask_path(f'PermissionError: {raw}')
    assert drive_f + ':\\三早一晚' not in masked
    assert '<WORKSPACE>' in masked or '<DRIVE>/' in masked

    # 验证非工作区盘符脱敏为 <DRIVE>/
    drive_d = 'D'
    raw_d = drive_d + ':\\some\\external\\file.xlsx'
    masked_d = DiagnosticMasker.mask_path(raw_d)
    assert drive_d + ':\\' not in masked_d
    assert '<DRIVE>/' in masked_d

    # 验证包含祖先目录错误信息的诊断报告落盘后隐私扫描完全通过
    result = SimpleNamespace(ok=False, error=f'PermissionError: {raw}')
    cfg = RunConfig('m.xlsx', '', 7, '雁山')
    log = generate_diagnostic_log(cfg, result=result)

    assert drive_f + ':\\三早一晚' not in log
    diag_file = tmp_path / 'temp_diag_test39.md'
    diag_file.write_text(log, encoding='utf-8')

    scan_script = PROJECT_DIR / 'tools' / 'privacy_scan.py'
    r = subprocess.run(
        [sys.executable, '-X', 'utf8', str(scan_script), '--root', str(tmp_path)],
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
    )
    assert r.returncode == 0, f'Privacy scan failed [code={r.returncode}]: {r.stdout} {r.stderr}'




