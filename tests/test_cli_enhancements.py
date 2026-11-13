# -*- coding: utf-8 -*-
"""自动化测试套件：覆盖 CLI 增强特性、安全自检、独立免勤报表、防篡改校验与全局双语 Skill 脚本。

测试范围：
1. CLI --json 结构化输出（状态码、阶段耗时、异常信封、摘要）；
2. CLI --preflight 表型自动识别（DETAIL / EXEMPT / DICTIONARY / SUMMARY / UNKNOWN）与健康自检（损坏、加密、独占占用）；
3. CLI 独立免勤分项报表导出（--exempt-report 且无 --detail）与回填差额核对；
4. CLI --verify-tamper 与 --against-detail（正常验签 0 与篡改报警 2）；
5. 核心层 dictionary.backup_dictionary 与 save_tables(backup=True) 时间戳安全备份；
6. 核心层 export.diff_actuals_detailed 结构化实到比对；
7. 全局 Skill 脚本（run_calc.py 零学号/姓名泄露断言，dict_tool.py 词典检查、备份与白名单拦截）。

注意约束：
- 零硬编码绝对路径（由 Path(__file__).resolve() 动态推导）；
- 虚构合成学号统一使用 99 开头约定号段；
- 绝不使用被否口径旋钮，确保隐私扫描 0 泄露。
"""

from __future__ import annotations

import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pytest
from openpyxl import Workbook, load_workbook

from cli import main as cli_main
from core.dictionary import (
    HEADERS,
    SHEET_NAMES,
    Dictionary,
    backup_dictionary,
    load_tables,
    save_tables,
)
from core.excel_safety import (
    DiagnosticMasker,
    classify_excel_file,
    preflight_excel_file,
    verify_excel_health,
)
from core.exempt_report import (
    export_standalone_exempt_report,
    reconcile_backfill_discrepancies,
)
from core.export import diff_actuals_detailed, verify_workbook_integrity
from tests._helpers import (
    focused_sheets,
    make_exempt_row,
    write_detail,
    write_exempt,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SID_REGEX = re.compile(r'(?<![0-9A-Fa-f])\d{13}(?![0-9A-Fa-f])')


def run_cli_in_process(args: List[str]) -> Tuple[int, str]:
    """在进程内运行 cli.main，捕获标准输出与退出码。"""
    buf = io.StringIO()
    with redirect_stdout(buf):
        try:
            code = cli_main(args)
        except SystemExit as exc:
            code = exc.code if isinstance(exc.code, int) else 1
    return code if code is not None else 0, buf.getvalue()


def run_cli_subprocess(args: List[str]) -> subprocess.CompletedProcess:
    """在子进程中运行 cli.py，测试物理进程隔离与标准输出。"""
    cli_path = REPO_ROOT / 'cli.py'
    return subprocess.run(
        [sys.executable, str(cli_path)] + args,
        capture_output=True,
        text=True,
        encoding='utf-8',
        cwd=str(REPO_ROOT),
    )


def find_skill_scripts_dir() -> Optional[Path]:
    """动态查找全局或本地 attendance-stats-calc 技能脚本目录，避免硬编码绝对路径。"""
    # 候选 1: 用户主目录下的 .agents/skills/attendance-stats-calc/scripts
    home_candidate = Path.home() / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts'
    if home_candidate.is_dir() and (home_candidate / 'run_calc.py').is_file():
        return home_candidate

    # 候选 2: 工作区同级发布目录下的 .agents/skills/attendance-stats-calc/scripts
    for parent in [REPO_ROOT.parent, REPO_ROOT.parent.parent]:
        for subname in ['三早一晚免勤自动统计', 'release', 'deploy']:
            cand = parent / subname / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts'
            if cand.is_dir() and (cand / 'run_calc.py').is_file():
                return cand

    # 候选 3: REPO_ROOT 向上追溯
    for parent in REPO_ROOT.parents:
        cand = parent / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts'
        if cand.is_dir() and (cand / 'run_calc.py').is_file():
            return cand

    return None


@pytest.fixture
def synthetic_sample(tmp_path: Path):
    """构建干净的微型合成数据集（明细表、免勤表、标准词典）。"""
    detail_path = tmp_path / 'detail_sample.xlsx'
    exempt_path = tmp_path / 'exempt_sample.xlsx'
    dict_path = tmp_path / 'dict_sample.xlsx'

    write_detail(detail_path, focused_sheets())
    write_exempt(
        exempt_path,
        {
            '早自习免勤人员统计表': [
                make_exempt_row(
                    sid='9900000000001',
                    college='理学院',
                    name='测试生甲',
                    class_text='应用统计学26-1班',
                    week_text='第7周',
                    day_text='周一',
                    session='早',
                    campus='雁山校区',
                    reason='生病',
                    room='8101',
                ),
            ],
            '晚自习免勤人员统计表': [],
        },
        with_room=True,
    )

    dummy_tables = {k: [] for k in HEADERS.keys()}
    dummy_tables['免勤理由分类'] = [
        ['生病', '健康医疗', '通用健康情形词，不含具体诊断或症状'],
        ['有选修课', '课务学业', '通用课程/事务类别，不含个人指向'],
    ]
    dummy_tables['学院别名'] = [['理学院', '理院', '标准测试别名']]
    save_tables(str(dict_path), dummy_tables)

    return {
        'detail': detail_path,
        'exempt': exempt_path,
        'dict': dict_path,
        'out': tmp_path / 'result_sample.xlsx',
    }


class TestCliJsonOutput:
    """CLI --json 结构化输出测试。"""

    def test_cli_json_full_pipeline_success(self, synthetic_sample):
        """测试完整计算流水线在 --json 模式下的返回字段、阶段耗时与状态码。"""
        out_file = synthetic_sample['out']
        args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
            '--json',
        ]
        code, stdout = run_cli_in_process(args)
        assert code == 0, f'CLI exited with code {code}, stdout: {stdout}'

        data = json.loads(stdout)
        assert data['status'] == 'ok'
        assert data['action'] == 'run'
        assert data['code'] == 0
        assert data['duration_ms'] >= 0

        phases = data['phases']
        assert 'preflight_ms' in phases
        assert 'parse_ms' in phases
        assert 'calc_ms' in phases
        assert 'export_ms' in phases

        assert os.path.isfile(data['output_file'])
        assert 'summary' in data
        assert isinstance(data['college_rates'], list)
        assert len(data['college_rates']) > 0
        assert isinstance(data['weekly_rates'], list)
        assert data['anomalies_count'] >= 0

    def test_cli_json_argument_error_envelope(self):
        """测试缺失必需参数时 --json 返回合规的错误信封。"""
        code, stdout = run_cli_in_process(['--json'])
        assert code == 1

        data = json.loads(stdout)
        assert data['action'] == 'cli_error'
        assert data['status'] == 'error'
        assert data['code'] == 1
        assert 'error' in data
        assert '--detail' in data['error']
        assert 'duration_ms' in data

    def test_cli_json_engine_error_envelope(self, tmp_path: Path):
        """测试引擎处理不存在的文件时 --json 返回规范错误结构。"""
        non_existent = tmp_path / 'ghost_file.xlsx'
        args = ['--detail', str(non_existent), '--json']
        code, stdout = run_cli_in_process(args)
        assert code == 1

        data = json.loads(stdout)
        assert data['action'] == 'run'
        assert data['status'] == 'error'
        assert data['code'] == 1
        assert 'error' in data

    def test_cli_json_subprocess_execution(self, synthetic_sample):
        """通过子进程验证 CLI --json 标准输出纯净性（无夹带非 JSON 文本）。"""
        out_file = synthetic_sample['detail'].parent / 'subproc_out.xlsx'
        args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
            '--json',
        ]
        res = run_cli_subprocess(args)
        assert res.returncode == 0
        # 严格校验标准输出能被纯净解析为 JSON
        parsed = json.loads(res.stdout)
        assert parsed['status'] == 'ok'
        assert parsed['code'] == 0


class TestCliPreflight:
    """CLI --preflight 前置体检与 4 类表型自动识别测试。"""

    def test_preflight_table_type_classification(self, tmp_path: Path, synthetic_sample):
        """测试 DETAIL、EXEMPT、DICTIONARY、SUMMARY 与 UNKNOWN 五类表型识别。"""
        # 1. DETAIL 表
        detail_type = classify_excel_file(str(synthetic_sample['detail']))
        assert detail_type == 'DETAIL'

        # 2. EXEMPT 表
        exempt_type = classify_excel_file(str(synthetic_sample['exempt']))
        assert exempt_type == 'EXEMPT'

        # 3. DICTIONARY 表
        dict_type = classify_excel_file(str(synthetic_sample['dict']))
        assert dict_type == 'DICTIONARY'

        # 4. SUMMARY 表（含出勤率与巡查率，且不含教室）
        summary_path = tmp_path / 'summary_test.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = 'Sheet2'
        ws.append(['序号', '学院', '学生出勤率', '值班辅导员巡查率', '执行效果'])
        ws.append([1, '理学院', '96.50%', '100.00%', '良好'])
        wb.save(summary_path)
        wb.close()
        summary_type = classify_excel_file(str(summary_path))
        assert summary_type == 'SUMMARY'

        # 5. UNKNOWN 表
        unknown_path = tmp_path / 'random_table.xlsx'
        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = 'Other'
        ws2.append(['Project', 'Score', 'Status'])
        ws2.append(['Alpha', 90, 'Pending'])
        wb2.save(unknown_path)
        wb2.close()
        unknown_type = classify_excel_file(str(unknown_path))
        assert unknown_type == 'UNKNOWN'

    def test_preflight_corrupted_zip_detection(self, tmp_path: Path):
        """测试损坏的截断 ZIP 文件被准确拦截并返回 CORRUPTED。"""
        corrupted_path = tmp_path / 'corrupted.xlsx'
        with open(corrupted_path, 'wb') as f:
            f.write(b'PK\x03\x04\x14\x00truncated_corrupted_stream_data')

        health = verify_excel_health(str(corrupted_path))
        assert not health.is_valid
        assert health.error_code == 'CORRUPTED'

        rep = preflight_excel_file(str(corrupted_path))
        assert not rep['is_valid']
        assert rep['error_code'] == 'CORRUPTED'

    def test_preflight_empty_file(self, tmp_path: Path):
        """测试 0 字节空文件拦截并返回 EMPTY_FILE。"""
        empty_path = tmp_path / 'empty.xlsx'
        empty_path.write_bytes(b'')

        health = verify_excel_health(str(empty_path))
        assert not health.is_valid
        assert health.error_code == 'EMPTY_FILE'

    def test_preflight_not_found_and_directory(self, tmp_path: Path):
        """测试不存在路径与目录路径拦截。"""
        ghost = tmp_path / 'ghost.xlsx'
        assert not verify_excel_health(str(ghost)).is_valid
        assert verify_excel_health(str(ghost)).error_code == 'NOT_FOUND'

        some_dir = tmp_path / 'folder.xlsx'
        some_dir.mkdir()
        assert not verify_excel_health(str(some_dir)).is_valid
        assert verify_excel_health(str(some_dir)).error_code == 'IS_DIRECTORY'

    def test_preflight_legacy_xls_format(self, tmp_path: Path):
        """测试 OLE2 旧版 .xls 格式拦截并返回 OLD_XLS_FORMAT。"""
        legacy_path = tmp_path / 'legacy.xlsx'
        with open(legacy_path, 'wb') as f:
            f.write(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\x00' * 128)

        health = verify_excel_health(str(legacy_path))
        assert not health.is_valid
        assert health.error_code == 'OLD_XLS_FORMAT'

    def test_preflight_encrypted_zip_detection(self, tmp_path: Path):
        """测试设置了加密标志位（flag_bits Bit 0）的 Excel 被拦截为 ENCRYPTED。"""
        enc_path = tmp_path / 'encrypted.xlsx'
        with zipfile.ZipFile(enc_path, 'w') as zf:
            zf.writestr('xl/workbook.xml', b'<xml>test</xml>')

        raw_bytes = bytearray(enc_path.read_bytes())
        # 在 Local Header (offset 6) 与 Central Directory 处打上加密位 0x01
        raw_bytes[6] |= 0x01
        cd_idx = raw_bytes.find(b'PK\x01\x02')
        if cd_idx != -1:
            raw_bytes[cd_idx + 8] |= 0x01
        enc_path.write_bytes(raw_bytes)

        health = verify_excel_health(str(enc_path))
        assert not health.is_valid
        assert health.error_code == 'ENCRYPTED'

    def test_preflight_locked_file_on_windows(self, tmp_path: Path):
        """测试独占打开的锁定文件在 Windows 下被正确识别为 LOCKED。"""
        if sys.platform != 'win32':
            pytest.skip('Exclusive file locking test applies to Windows')

        import ctypes
        locked_path = tmp_path / 'locked.xlsx'
        locked_path.write_bytes(b'PK\x03\x04' + b'\x00' * 60)

        GENERIC_READ = 0x80000000
        GENERIC_WRITE = 0x40000000
        FILE_SHARE_NONE = 0
        OPEN_EXISTING = 3

        handle = ctypes.windll.kernel32.CreateFileW(
            str(locked_path),
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_NONE,
            None,
            OPEN_EXISTING,
            0,
            None,
        )
        assert handle != -1
        try:
            health = verify_excel_health(str(locked_path))
            assert not health.is_valid
            assert health.error_code == 'LOCKED'
        finally:
            ctypes.windll.kernel32.CloseHandle(handle)

    def test_cli_preflight_command_json_and_plain(self, synthetic_sample):
        """测试 CLI --preflight 在普通文本回显与 --json 模式下的执行。"""
        # 1. 正常多文件体检 (--json)
        args_json = [
            '--preflight',
            str(synthetic_sample['detail']),
            str(synthetic_sample['exempt']),
            str(synthetic_sample['dict']),
            '--json',
        ]
        code, stdout = run_cli_in_process(args_json)
        assert code == 0
        data = json.loads(stdout)
        assert data['action'] == 'preflight'
        assert data['status'] == 'ok'
        assert len(data['reports']) == 3

        # 2. 普通文本回显模式
        args_plain = [
            '--preflight',
            str(synthetic_sample['detail']),
            str(synthetic_sample['dict']),
        ]
        code2, stdout2 = run_cli_in_process(args_plain)
        assert code2 == 0
        assert '前置 Excel 健康诊断与表型自检' in stdout2
        assert 'DETAIL' in stdout2
        assert 'DICTIONARY' in stdout2


class TestCliExemptReportStandalone:
    """独立免勤分项报表导出与差额对账测试。"""

    def test_standalone_exempt_report_cli_json(self, synthetic_sample, tmp_path: Path):
        """测试 CLI standalone --exempt-report（无 --detail）在 --json 模式下导出。"""
        rep_out = tmp_path / 'standalone_report.xlsx'
        args = [
            '--exempt-report', str(rep_out),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--week', '7',
            '--json',
        ]
        code, stdout = run_cli_in_process(args)
        assert code == 0, f'Failed with code {code}: {stdout}'

        data = json.loads(stdout)
        assert data['action'] == 'exempt_report_standalone'
        assert data['status'] == 'ok'
        assert data['code'] == 0
        assert os.path.isfile(data['output_path'])

        # 验证生成的 Excel 包含免勤统计工作表
        wb = load_workbook(rep_out, data_only=True)
        assert '免勤统计' in wb.sheetnames
        ws = wb['免勤统计']
        assert ws.cell(1, 1).value == '学院'
        wb.close()

    def test_standalone_exempt_report_missing_exempt_flag_error(self, tmp_path: Path):
        """测试独立运行模式缺少 --exempt 时给出明确错误。"""
        rep_out = tmp_path / 'standalone_report.xlsx'
        args = ['--exempt-report', str(rep_out), '--json']
        code, stdout = run_cli_in_process(args)
        assert code == 1
        data = json.loads(stdout)
        assert data['action'] == 'cli_error'
        assert '必须指定至少一份 --exempt' in data['error']

    def test_core_export_standalone_exempt_report_api(self, synthetic_sample, tmp_path: Path):
        """测试核心模块 export_standalone_exempt_report API 调用。"""
        rep_out = tmp_path / 'api_standalone_report.xlsx'
        metrics = export_standalone_exempt_report(
            output_path=str(rep_out),
            exempt_paths=[str(synthetic_sample['exempt'])],
            week_no=7,
            dictionary_path=str(synthetic_sample['dict']),
        )
        assert os.path.isfile(rep_out)
        assert metrics['rows_count'] >= 1
        assert metrics['total_exempt'] >= 1
        assert 'anomalies_count' in metrics

    def test_core_reconcile_backfill_discrepancies(self):
        """测试 reconcile_backfill_discrepancies 对账差额提炼逻辑。"""
        # 1. 正常行：回填差额为 0，状态正常
        normal_rows = [
            {'学院': '理学院', '星期': '周一', '自习': '早', '免勤合计': 5, '已回填': 5, '容量超额': 0, '回填差额': 0, '回填状态': '已回填'},
        ]
        diffs1 = reconcile_backfill_discrepancies(normal_rows)
        assert len(diffs1) == 0

        # 2. 存在差额或需复核状态行
        discrepant_rows = [
            {'学院': '地科院', '星期': '周一', '自习': '早', '免勤合计': 10, '已回填': 6, '容量超额': 2, '回填差额': 2, '回填状态': '容量超额，需复核'},
            {'学院': '化生院', '星期': '周二', '自习': '晚', '免勤合计': 4, '已回填': 4, '容量超额': 0, '回填差额': 0, '回填状态': '已回填'},
        ]
        diffs2 = reconcile_backfill_discrepancies(discrepant_rows)
        assert len(diffs2) == 1
        assert diffs2[0]['college'] == '地科院'
        assert diffs2[0]['difference'] == 2
        assert diffs2[0]['overflow'] == 2


class TestCliVerifyTamperAndAgainstDetail:
    """CLI --verify-tamper 与 --against-detail 防篡改校验测试。"""

    def test_cli_verify_tamper_happy_path(self, synthetic_sample):
        """测试合法计算产物的防篡改签名与原表核对（返回码 0）。"""
        out_file = synthetic_sample['out']
        # 先运行生成带防篡改签名的产物
        run_args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
            '--json',
        ]
        run_code, _ = run_cli_in_process(run_args)
        assert run_code == 0

        # 执行防篡改校验与对拍
        verify_args = [
            '--verify-tamper', str(out_file),
            '--against-detail', str(synthetic_sample['detail']),
            '--json',
        ]
        v_code, v_stdout = run_cli_in_process(verify_args)
        assert v_code == 0, f'Expected 0, got {v_code}: {v_stdout}'

        v_data = json.loads(v_stdout)
        assert v_data['action'] == 'verify_tamper'
        assert v_data['status'] == 'ok'
        assert v_data['code'] == 0
        assert v_data['verified'] is True
        assert v_data['violations_count'] == 0
        assert len(v_data['diffs']) == 0

    def test_cli_verify_tamper_detects_signature_corruption(self, synthetic_sample):
        """测试篡改防篡改工作表签名后，验签返回状态码 2。"""
        out_file = synthetic_sample['out']
        run_args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
        ]
        run_code, _ = run_cli_in_process(run_args)
        assert run_code == 0

        # 手工破坏元数据工作表签名 (第 2 行第 6 列 manifest_hmac_sha256)
        wb = load_workbook(out_file)
        if '_系统校验元数据' in wb.sheetnames:
            ws = wb['_系统校验元数据']
            ws.cell(2, 6).value = 'corrupted_manifest_hash_1234567890abcdef'
        wb.save(out_file)
        wb.close()

        verify_args = ['--verify-tamper', str(out_file), '--json']
        v_code, v_stdout = run_cli_in_process(verify_args)
        assert v_code == 2

        v_data = json.loads(v_stdout)
        assert v_data['status'] == 'tamper_detected'
        assert v_data['code'] == 2
        assert v_data['verified'] is False
        assert v_data['violations_count'] > 0

    def test_cli_verify_tamper_against_detail_detects_headcount_tamper(self, synthetic_sample):
        """测试篡改实到人数单元格后，对拍返回状态码 2 并输出结构化差异清单。"""
        out_file = synthetic_sample['out']
        run_args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
        ]
        run_code, _ = run_cli_in_process(run_args)
        assert run_code == 0

        # 修改结果表中的某处实到人数
        wb = load_workbook(out_file)
        ws = wb['雁山（早）']
        # 寻找实到列修改
        ws.cell(4, 6).value = 999
        wb.save(out_file)
        wb.close()

        verify_args = [
            '--verify-tamper', str(out_file),
            '--against-detail', str(synthetic_sample['detail']),
            '--json',
        ]
        v_code, v_stdout = run_cli_in_process(verify_args)
        assert v_code == 2

        v_data = json.loads(v_stdout)
        assert v_data['status'] == 'tamper_detected'
        assert v_data['code'] == 2
        assert len(v_data['diffs']) > 0
        diff_item = v_data['diffs'][0]
        assert 'sheet' in diff_item
        assert 'cell' in diff_item
        assert 'source_value' in diff_item
        assert 'result_value' in diff_item
        assert diff_item['result_value'] == 999

    def test_cli_against_detail_without_verify_tamper_flag_error(self, synthetic_sample):
        """测试单独使用 --against-detail 而缺少 --verify-tamper 时报错。"""
        args = ['--against-detail', str(synthetic_sample['detail']), '--json']
        code, stdout = run_cli_in_process(args)
        assert code == 1
        data = json.loads(stdout)
        assert data['action'] == 'cli_error'
        assert '--against-detail 必须与 --verify-tamper 同时使用' in data['error']


class TestCoreDictionaryBackupAndDiff:
    """核心层词典备份与实到比对测试。"""

    def test_backup_dictionary_creates_timestamped_copy(self, tmp_path: Path):
        """测试 backup_dictionary 创建标准命名的时间戳备份文件。"""
        dict_path = tmp_path / '免勤词典.xlsx'
        dummy_tables = {k: [] for k in HEADERS.keys()}
        dummy_tables['学院别名'] = [['理学院', '理院', '测试']]
        save_tables(str(dict_path), dummy_tables)

        backup_path = backup_dictionary(str(dict_path))
        assert backup_path is not None
        assert os.path.isfile(backup_path)
        assert '_backup_' in os.path.basename(backup_path)
        assert os.path.getsize(backup_path) == os.path.getsize(dict_path)

    def test_backup_dictionary_invalid_path_returns_none(self, tmp_path: Path):
        """测试对不存在文件调用 backup_dictionary 返回 None。"""
        ghost = tmp_path / 'nonexistent_dict.xlsx'
        assert backup_dictionary(str(ghost)) is None
        assert backup_dictionary('') is None

    def test_save_tables_with_backup_parameter(self, tmp_path: Path):
        """测试 save_tables(backup=True) 自动执行前置备份并落地修改。"""
        dict_path = tmp_path / '免勤词典.xlsx'
        tables_v1 = {k: [] for k in HEADERS.keys()}
        tables_v1['学院别名'] = [['理学院', '理院', 'v1']]
        save_tables(str(dict_path), tables_v1, backup=False)

        # 第二次保存时传入 backup=True
        tables_v2 = {k: [] for k in HEADERS.keys()}
        tables_v2['学院别名'] = [['理学院', '理院', 'v2']]
        backup_created = save_tables(str(dict_path), tables_v2, backup=True)
        assert backup_created is not None
        assert os.path.isfile(backup_created)

        # 检查最新词典已更新
        loaded = load_tables(str(dict_path))
        assert loaded['学院别名'][0][2] == 'v2'

    def test_core_diff_actuals_detailed(self, synthetic_sample):
        """测试 core.export.diff_actuals_detailed 结构化差异比对。"""
        src = synthetic_sample['detail']
        res = synthetic_sample['out']

        # 生成正确产物
        run_args = [
            '--detail', str(src),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(res),
            '--week', '7',
        ]
        assert run_cli_in_process(run_args)[0] == 0

        # 未被修改时，diffs 为空
        diffs_clean = diff_actuals_detailed(str(src), str(res))
        assert len(diffs_clean) == 0

        # 修改目标单元格
        wb = load_workbook(res)
        wb['雁山（早）'].cell(4, 6).value = 777
        wb.save(res)
        wb.close()

        diffs_dirty = diff_actuals_detailed(str(src), str(res))
        assert len(diffs_dirty) == 1
        d = diffs_dirty[0]
        assert d['sheet'] == '雁山（早）'
        assert d['row'] == 4
        assert d['col'] == 6
        assert d['result_value'] == 777


class TestSkillScriptsIntegration:
    """全局双语 Skill 包装脚本（run_calc.py 与 dict_tool.py）集成测试。"""

    @pytest.fixture(autouse=True)
    def check_skill_scripts_available(self):
        scripts_dir = find_skill_scripts_dir()
        if not scripts_dir:
            pytest.skip('Skill scripts directory not located in test environment')
        self.scripts_dir = scripts_dir
        self.run_calc = scripts_dir / 'run_calc.py'
        self.dict_tool = scripts_dir / 'dict_tool.py'

    def test_skill_run_calc_preflight(self, synthetic_sample):
        """测试 run_calc.py --preflight 输出干净合规 JSON。"""
        cmd = [
            sys.executable,
            str(self.run_calc),
            '--preflight',
            str(synthetic_sample['detail']),
            str(synthetic_sample['dict']),
            '--app-root',
            str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 0
        data = json.loads(res.stdout)
        assert data['action'] == 'preflight'
        assert data['status'] == 'ok'
        assert len(data['reports']) == 2

    def test_skill_run_calc_full_calc_zero_pii(self, synthetic_sample, tmp_path: Path):
        """测试 run_calc.py 执行完整计算，严格断言标准输出中 0 学号与 0 姓名泄露。"""
        out_file = tmp_path / 'skill_calc_out.xlsx'
        cmd = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
            '--app-root', str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 0, f'run_calc failed: {res.stderr}'

        stdout = res.stdout
        # 1. 严格正则隐私断言：stdout 中绝不能出现任何 13 位数字串
        sid_match = SID_REGEX.search(stdout)
        assert sid_match is None, f'Privacy violation: student ID found in stdout: {sid_match.group(0)}'

        # 2. 姓名隐私断言：测试数据中的姓名不可在 stdout 中出现
        assert '测试生甲' not in stdout

        # 3. 校验聚合结果
        data = json.loads(stdout)
        assert data['action'] == 'run'
        assert data['status'] == 'success'
        assert data['code'] == 0
        assert 'college_rates' in data
        assert 'weekly_rates' in data
        assert 'anomaly_summary' in data
        assert os.path.isfile(data['output_file'])

    def test_skill_run_calc_verify_tamper(self, synthetic_sample):
        """测试 run_calc.py --verify-tamper 防篡改检验。"""
        out_file = synthetic_sample['out']
        # 先生成
        cmd_run = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--out', str(out_file),
            '--week', '7',
            '--app-root', str(REPO_ROOT),
        ]
        assert subprocess.run(cmd_run, capture_output=True, text=True, encoding='utf-8').returncode == 0

        # 验签合法
        cmd_verify = [
            sys.executable,
            str(self.run_calc),
            '--verify-tamper', str(out_file),
            '--against-detail', str(synthetic_sample['detail']),
            '--app-root', str(REPO_ROOT),
        ]
        res_v = subprocess.run(cmd_verify, capture_output=True, text=True, encoding='utf-8')
        assert res_v.returncode == 0
        data_v = json.loads(res_v.stdout)
        assert data_v['status'] == 'ok'
        assert data_v['verified'] is True

    def test_skill_dict_tool_check(self, synthetic_sample):
        """测试 dict_tool.py check 校验通过。"""
        cmd = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'check',
            '--dict', str(synthetic_sample['dict']),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 0
        assert '[OK] Dictionary validation passed' in res.stdout

    def test_skill_dict_tool_backup(self, synthetic_sample, tmp_path: Path):
        """测试 dict_tool.py backup 生成时间戳备份。"""
        cmd = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'backup',
            '--dict', str(synthetic_sample['dict']),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 0
        assert '[BACKUP]' in res.stdout

    def test_skill_dict_tool_add_alias_whitelist_enforcement(self, synthetic_sample, tmp_path: Path):
        """测试 dict_tool.py add-alias 的严格白名单拦截、差异预览与自动备份。"""
        dict_copy = tmp_path / 'dict_for_tool.xlsx'
        shutil.copy2(synthetic_sample['dict'], dict_copy)

        # 1. 尝试添加非法免勤理由分类 -> 必须拦截并退出码 1
        cmd_bad = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'add-alias',
            '--sheet', '免勤理由分类',
            '--alias', '急性肠胃炎',
            '--target', '非法未批准分类',
            '--dict', str(dict_copy),
        ]
        res_bad = subprocess.run(cmd_bad, capture_output=True, text=True, encoding='utf-8')
        assert res_bad.returncode == 1
        assert 'Whitelist violation' in res_bad.stderr or 'Whitelist violation' in res_bad.stdout

        # 2. 添加合法的免勤理由分类（在白名单内：健康医疗）-> 必须成功并前置备份
        cmd_good = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'add-alias',
            '--sheet', '免勤理由分类',
            '--alias', '发热就诊',
            '--target', '健康医疗',
            '--note', '测试合法新增',
            '--dict', str(dict_copy),
        ]
        res_good = subprocess.run(cmd_good, capture_output=True, text=True, encoding='utf-8')
        assert res_good.returncode == 0
        assert '[DIFF]' in res_good.stdout
        assert '[BACKUP]' in res_good.stdout
        assert '[SUCCESS]' in res_good.stdout

        # 3. 添加学院别名 -> 必须成功
        cmd_alias = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'add-alias',
            '--sheet', '学院别名',
            '--alias', '数理学院',
            '--target', '理学院',
            '--dict', str(dict_copy),
        ]
        res_alias = subprocess.run(cmd_alias, capture_output=True, text=True, encoding='utf-8')
        assert res_alias.returncode == 0
        assert '[SUCCESS]' in res_alias.stdout


class TestCliRemediationAndHardening:
    """全面整改增强测试：脱敏、参数暴露、运行记录查询与词典校验。"""

    def test_preflight_json_eliminates_raw_path(self, synthetic_sample):
        """测试 --preflight --json 输出中完全剔除 raw_path 字段。"""
        args = [
            '--preflight',
            str(synthetic_sample['detail']),
            str(synthetic_sample['dict']),
            '--json',
        ]
        code, stdout = run_cli_in_process(args)
        assert code == 0
        data = json.loads(stdout)
        assert data['action'] == 'preflight'
        assert data['status'] == 'ok'
        assert len(data['reports']) == 2
        for rep in data['reports']:
            assert 'raw_path' not in rep, "raw_path must be eliminated from preflight json reports"
            assert 'path' in rep

    def test_sid_pattern_assertion_and_anomaly_masking(self):
        """测试 13 位学号（SID）隐私断言与异常条目脱敏。"""
        from cli import SID_PATTERN, _print_json
        import io
        import contextlib

        # 1. 验证 SID_PATTERN 拦截未脱敏 13 位真实学号（非 99 开头）
        unmasked_test_sid = '2024' + '012345678'
        leak_payload = {"student_id": unmasked_test_sid, "action": "test"}
        with pytest.raises(AssertionError, match="Privacy violation"):
            _print_json(leak_payload)

        # 2. 验证脱敏后的学号（带星号）与 99 开头的测试占位学号不被拦截
        safe_payload = {
            "masked_sid": "2024*******78",
            "test_sid": "9900000000001",
            "action": "test",
        }
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            _print_json(safe_payload)
        assert "2024*******78" in buf.getvalue()

    def test_cli_complete_mode_and_summary_out(self, synthetic_sample, tmp_path: Path):
        """测试 --complete-mode 与 --summary-out 核心参数。"""
        out_detail = tmp_path / 'out_detail.xlsx'
        out_summary = tmp_path / 'custom_summary_final.xlsx'

        args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--complete-mode', 'tolerant',
            '--summary-out', str(out_summary),
            '--out', str(out_detail),
            '--week', '7',
            '--json',
        ]
        code, stdout = run_cli_in_process(args)
        assert code == 0
        data = json.loads(stdout)
        assert data['action'] == 'run'
        assert data['status'] == 'ok'
        assert out_detail.is_file()
        assert out_summary.is_file()
        assert data['summary_output_file'] == str(out_summary)

    def test_cli_storage_list_and_show_runs(self):
        """测试 --list-runs 与 --show-run 历史批次查询。"""
        # 1. 测试 --list-runs --json
        code, stdout = run_cli_in_process(['--list-runs', '--json'])
        assert code == 0
        data = json.loads(stdout)
        assert data['action'] == 'list_runs'
        assert data['status'] == 'ok'
        assert isinstance(data['runs'], list)

        # 2. 测试 --list-runs 文本模式
        code_text, stdout_text = run_cli_in_process(['--list-runs'])
        assert code_text == 0
        assert '历史计算记录' in stdout_text

        # 3. 测试 --show-run 不存在 ID
        code_missing, stdout_missing = run_cli_in_process(['--show-run', '99999999', '--json'])
        assert code_missing == 1
        data_missing = json.loads(stdout_missing)
        assert data_missing['status'] == 'not_found'

    def test_cli_diagnostic_log_export(self, synthetic_sample, tmp_path: Path):
        """测试 --diagnostic-log 参数一键导出脱敏 Markdown 诊断报告。"""
        out_detail = tmp_path / 'out_detail_diag.xlsx'
        diag_log = tmp_path / 'diagnostic_report.md'

        args = [
            '--detail', str(synthetic_sample['detail']),
            '--exempt', str(synthetic_sample['exempt']),
            '--dictionary', str(synthetic_sample['dict']),
            '--diagnostic-log', str(diag_log),
            '--out', str(out_detail),
            '--week', '7',
            '--json',
        ]
        code, stdout = run_cli_in_process(args)
        assert code == 0
        assert diag_log.is_file()
        content = diag_log.read_text(encoding='utf-8')
        assert '诊断日志' in content
        assert '运行环境信息' in content
        assert '文件指纹' in content

    def test_dict_tool_semantic_prevalidation_aborts_without_backup(self, synthetic_sample, tmp_path: Path):
        """测试 dict_tool.py add-alias 在非法分类时直接熔断退出且不产生孤儿备份文件。"""
        import shutil
        dict_copy = tmp_path / 'dict_test_preval.xlsx'
        shutil.copy2(synthetic_sample['dict'], dict_copy)

        dict_tool = REPO_ROOT.parent.parent / '三早一晚免勤自动统计' / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts' / 'dict_tool.py'
        if not dict_tool.is_file():
            dict_tool = Path(os.environ.get('USERPROFILE', '')) / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts' / 'dict_tool.py'

        # 记录初始目录内文件清单
        initial_files = set(tmp_path.iterdir())

        # 执行非法分类添加
        cmd_bad = [
            sys.executable,
            str(dict_tool),
            '--app-root', str(REPO_ROOT),
            'add-alias',
            '--sheet', '免勤理由分类',
            '--alias', '非法理由测试',
            '--target', '未知非法分类名称',
            '--dict', str(dict_copy),
        ]
        res = subprocess.run(cmd_bad, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 1
        assert 'Whitelist violation' in res.stderr

        # 验证没有生成任何新的 .bak 备份文件（无孤儿备份）
        after_files = set(tmp_path.iterdir())
        new_files = after_files - initial_files
        bak_files = [f for f in new_files if '.bak' in f.name]
        assert len(bak_files) == 0, f"Orphan backup files created on validation failure: {bak_files}"

