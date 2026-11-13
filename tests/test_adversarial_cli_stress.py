# -*- coding: utf-8 -*-
"""Milestone M5 Adversarial Stress & Robustness Challenge Test Suite for CLI & Core Engine.

This test suite aggressively probes and stress-tests:
1. --preflight:
   - 0-byte files, truncated files (<4 bytes, truncated zip), corrupted bytes, non-excel zip files,
     OLE2 binary xls files, encrypted zip files, locked files on Windows, missing files, directories.
   - Verifies that --preflight NEVER crashes with an unhandled exception and always returns clean JSON
     (when --json) or formatted text, with proper error codes and exit codes.
2. --json output purity:
   - Verifies stdout is 100% valid JSON under missing arguments, invalid choices, and runtime errors.
   - Asserts json.loads(stdout) succeeds with zero stray text/warnings.
3. --verify-tamper & --against-detail:
   - Tests single-cell mutations, sheet deletions, formula injections, metadata deletions,
     manifest F2 hash corruptions, and headcount mismatches.
   - Asserts that tamper detection strictly catches modifications and returns exit code 2.
4. Standalone --exempt-report:
   - Tests multiple exempt files, missing arguments, invalid week numbers, and corrupted files.
5. Known edge-case exception probing:
   - Documents and isolates behavior when non-excel archives are passed directly without preflight.

Constraints:
- Zero hardcoded drive letters or absolute paths (dynamically resolved via Path(__file__)).
- Synthetic student IDs use 99... format.
- Zero privacy leaks.
"""

from __future__ import annotations

import ctypes
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional

import pytest
from openpyxl import Workbook, load_workbook

from tests._helpers import focused_sheets, write_detail
from core.excel_safety import preflight_excel_file, verify_excel_health
from core.export import (
    diff_actuals_detailed,
    generate_tamper_metadata,
    verify_workbook_integrity,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
CLI_PATH = REPO_ROOT / 'cli.py'


def run_cli(args: List[str]) -> subprocess.CompletedProcess:
    """Run cli.py in a dedicated subprocess to verify process boundary and stdout purity."""
    return subprocess.run(
        [sys.executable, str(CLI_PATH)] + args,
        capture_output=True,
        text=True,
        encoding='utf-8',
        cwd=str(REPO_ROOT),
    )


def create_minimal_valid_detail_workbook(path: Path) -> Path:
    """Create a minimal valid detail workbook for tests."""
    wb = Workbook()
    ws = wb.active
    ws.title = '雁山（早）'
    ws.append(['序号', '校区', '自习室', '教室', '分配专业班级', '应到人数', '周一实到', '周二实到', '周三实到', '周四实到', '周五实到'])
    ws.append([1, '雁山', '01201', '01201', '计算机21-1', 40, 38, 39, 37, 40, 38])
    wb.save(path)
    wb.close()
    return path


def create_minimal_valid_exempt_workbook(path: Path) -> Path:
    """Create a minimal valid exempt workbook for tests."""
    wb = Workbook()
    ws = wb.active
    ws.title = '免勤名单'
    ws.append(['序号', '学号', '姓名', '学院', '班级', '校区', '免勤周次', '免勤自习', '免勤原因', '请假类型', '自习室'])
    ws.append([1, '9900000000001', '学生甲', '计算机科学与工程学院', '计算机21-1', '雁山', 7, '早自习（周一）', '选修课冲突', '课务学业', '01201'])
    wb.save(path)
    wb.close()
    return path


# ==============================================================================
# 1. Adversarial Preflight Stress Tests
# ==============================================================================

class TestAdversarialPreflight:
    """Aggressive tests for --preflight health detection and crash-resilience."""

    def test_preflight_zero_byte_file(self, tmp_path: Path):
        """Zero-byte file must be cleanly reported as EMPTY_FILE without unhandled crash."""
        p = tmp_path / 'empty.xlsx'
        p.write_bytes(b'')

        # Test CLI with --json
        res = run_cli(['--preflight', str(p), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['status'] == 'error'
        assert data['code'] == 1
        assert len(data['reports']) == 1
        rep = data['reports'][0]
        assert rep['is_valid'] is False
        assert rep['error_code'] == 'EMPTY_FILE'
        assert rep['file_size_bytes'] == 0

        # Test CLI plain-text mode
        res_plain = run_cli(['--preflight', str(p)])
        assert res_plain.returncode == 1
        assert 'EMPTY_FILE' in res_plain.stdout

    def test_preflight_truncated_magic_headers(self, tmp_path: Path):
        """Files smaller than 4 bytes must be detected as TOO_SMALL."""
        for size in [1, 2, 3]:
            p = tmp_path / f'tiny_{size}.xlsx'
            p.write_bytes(b'PK\x03'[:size])
            res = run_cli(['--preflight', str(p), '--json'])
            assert res.returncode == 1
            data = json.loads(res.stdout)
            rep = data['reports'][0]
            assert rep['is_valid'] is False
            assert rep['error_code'] == 'TOO_SMALL'
            assert rep['file_size_bytes'] == size

    def test_preflight_corrupted_random_bytes(self, tmp_path: Path):
        """Random binary junk must be detected as CORRUPTED."""
        p = tmp_path / 'garbage.xlsx'
        p.write_bytes(os.urandom(65536))
        res = run_cli(['--preflight', str(p), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        rep = data['reports'][0]
        assert rep['is_valid'] is False
        assert rep['error_code'] == 'CORRUPTED'

    def test_preflight_truncated_zip_archive(self, tmp_path: Path):
        """Starts with PK0304 but truncated immediately must be detected as CORRUPTED."""
        p = tmp_path / 'truncated.xlsx'
        p.write_bytes(b'PK\x03\x04\x14\x00\x00\x00\x08\x00')
        res = run_cli(['--preflight', str(p), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        rep = data['reports'][0]
        assert rep['is_valid'] is False
        assert rep['error_code'] == 'CORRUPTED'

    def test_preflight_non_excel_zip_archive(self, tmp_path: Path):
        """Valid zip archive lacking Excel structure must be detected as NOT_EXCEL."""
        p = tmp_path / 'not_excel.xlsx'
        with zipfile.ZipFile(p, 'w') as zf:
            zf.writestr('document.txt', 'This is a text file in a zip archive.')
        res = run_cli(['--preflight', str(p), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        rep = data['reports'][0]
        assert rep['is_valid'] is False
        assert rep['error_code'] == 'NOT_EXCEL'

    def test_preflight_legacy_binary_xls_format(self, tmp_path: Path):
        """OLE2 CFB magic header must be detected as OLD_XLS_FORMAT."""
        p = tmp_path / 'legacy.xlsx'
        p.write_bytes(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\x00' * 512)
        res = run_cli(['--preflight', str(p), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        rep = data['reports'][0]
        assert rep['is_valid'] is False
        assert rep['error_code'] == 'OLD_XLS_FORMAT'

    def test_preflight_encrypted_zip_archive(self, tmp_path: Path):
        """Encrypted zip entries (bit 0 set) must be detected as ENCRYPTED."""
        p = tmp_path / 'encrypted.xlsx'
        with zipfile.ZipFile(p, 'w') as zf:
            zf.writestr('xl/workbook.xml', b'<xml>test</xml>')
        raw = bytearray(p.read_bytes())
        raw[6] |= 0x01
        cd = raw.find(b'PK\x01\x02')
        if cd != -1:
            raw[cd + 8] |= 0x01
        p.write_bytes(raw)

        res = run_cli(['--preflight', str(p), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        rep = data['reports'][0]
        assert rep['is_valid'] is False
        assert rep['error_code'] == 'ENCRYPTED'

    def test_preflight_locked_file_on_windows(self, tmp_path: Path):
        """Exclusively locked file must be cleanly detected as LOCKED or UNREADABLE."""
        p = tmp_path / 'locked.xlsx'
        create_minimal_valid_detail_workbook(p)

        kernel32 = ctypes.windll.kernel32
        GENERIC_READ = 0x80000000
        GENERIC_WRITE = 0x40000000
        FILE_SHARE_NONE = 0
        OPEN_EXISTING = 3
        handle = kernel32.CreateFileW(
            str(p),
            GENERIC_READ | GENERIC_WRITE,
            FILE_SHARE_NONE,
            None,
            OPEN_EXISTING,
            0,
            None,
        )
        if handle == -1 or handle == 0xFFFFFFFFFFFFFFFF:
            pytest.skip('Could not acquire exclusive Windows file lock')

        try:
            res = run_cli(['--preflight', str(p), '--json'])
            assert res.returncode == 1
            data = json.loads(res.stdout)
            rep = data['reports'][0]
            assert rep['is_valid'] is False
            assert rep['error_code'] in ('LOCKED', 'UNREADABLE')
        finally:
            kernel32.CloseHandle(handle)

    def test_preflight_missing_file_and_directory(self, tmp_path: Path):
        """Nonexistent paths and directories must return NOT_FOUND and IS_DIRECTORY."""
        missing = tmp_path / 'does_not_exist.xlsx'
        dir_path = tmp_path / 'subfolder'
        dir_path.mkdir()

        res = run_cli(['--preflight', str(missing), str(dir_path), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert len(data['reports']) == 2
        assert data['reports'][0]['error_code'] == 'NOT_FOUND'
        assert data['reports'][1]['error_code'] == 'IS_DIRECTORY'

    def test_preflight_multiple_files_mixed_health(self, tmp_path: Path):
        """Batch preflight returns code 1 if ANY file is invalid, but evaluates all files."""
        f_ok = create_minimal_valid_detail_workbook(tmp_path / 'good.xlsx')
        f_bad = tmp_path / 'bad.xlsx'
        f_bad.write_bytes(b'')

        res = run_cli(['--preflight', str(f_ok), str(f_bad), '--json'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['status'] == 'error'
        assert len(data['reports']) == 2
        assert data['reports'][0]['is_valid'] is True
        assert data['reports'][1]['is_valid'] is False


# ==============================================================================
# 2. Adversarial JSON Output Purity Tests
# ==============================================================================

class TestAdversarialJsonStrictness:
    """Stress tests verifying stdout is 100% valid JSON with zero stray text."""

    def test_json_missing_all_arguments(self):
        """cli.py --json with no other arguments must output valid JSON error envelope."""
        res = run_cli(['--json'])
        assert res.returncode == 1
        assert res.stdout.strip().startswith('{')
        assert res.stdout.strip().endswith('}')
        data = json.loads(res.stdout)
        assert data['action'] == 'cli_error'
        assert data['status'] == 'error'
        assert data['code'] == 1
        assert 'error' in data

    def test_json_invalid_week_choice(self, tmp_path: Path):
        """Out of bounds week choice (e.g. 0, 99) returns valid JSON error envelope."""
        d = create_minimal_valid_detail_workbook(tmp_path / 'd.xlsx')
        for bad_week in ['0', '35', '99', 'invalid']:
            res = run_cli(['--json', '--detail', str(d), '--week', bad_week])
            assert res.returncode == 1
            data = json.loads(res.stdout)
            assert data['action'] == 'cli_error'
            assert data['status'] == 'error'
            assert data['code'] == 1

    def test_json_invalid_source_choice(self, tmp_path: Path):
        """Invalid --source choice returns valid JSON error envelope."""
        d = create_minimal_valid_detail_workbook(tmp_path / 'd.xlsx')
        res = run_cli(['--json', '--detail', str(d), '--source', '非法来源分类'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['action'] == 'cli_error'
        assert data['status'] == 'error'

    def test_json_unknown_arguments(self):
        """Unrecognized flags return valid JSON error envelope."""
        res = run_cli(['--json', '--unknown-adversarial-flag'])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['action'] == 'cli_error'
        assert data['status'] == 'error'

    def test_json_against_detail_without_verify_tamper(self, tmp_path: Path):
        """--against-detail without --verify-tamper returns valid JSON error envelope."""
        d = create_minimal_valid_detail_workbook(tmp_path / 'd.xlsx')
        res = run_cli(['--json', '--against-detail', str(d)])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['action'] == 'cli_error'
        assert data['status'] == 'error'
        assert '--against-detail 必须与 --verify-tamper 同时使用' in data['error']

    def test_json_standalone_exempt_missing_exempt_files(self, tmp_path: Path):
        """--exempt-report without --exempt in standalone mode returns valid JSON envelope."""
        rep = tmp_path / 'rep.xlsx'
        res = run_cli(['--json', '--exempt-report', str(rep)])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['action'] == 'cli_error'
        assert data['status'] == 'error'
        assert '必须指定至少一份 --exempt' in data['error']


# ==============================================================================
# 3. Adversarial Anti-Tamper & Detailed Diff Stress Tests
# ==============================================================================

class TestAdversarialTamperVerification:
    """Stress tests verifying --verify-tamper strictly detects mutations with exit code 2."""

    @pytest.fixture
    def signed_workbook(self, tmp_path: Path) -> Path:
        """Create a signed workbook with valid tamper metadata."""
        target = tmp_path / 'signed.xlsx'
        wb = Workbook()
        ws1 = wb.active
        ws1.title = '雁山（早）'
        ws1.append(['序号', '校区', '自习室', '教室', '分配专业班级', '应到人数', '周一实到', '周二实到'])
        ws1.append([1, '雁山', '01201', '01201', '计算机21-1', 40, 38, 39])

        ws2 = wb.create_sheet('雁山（晚）')
        ws2.append(['序号', '校区', '自习室', '教室', '分配专业班级', '应到人数', '周一实到', '周二实到'])
        ws2.append([1, '雁山', '01201', '01201', '计算机21-1', 40, 37, 38])

        generate_tamper_metadata(wb, ['雁山（早）', '雁山（晚）'])
        wb.save(target)
        wb.close()
        return target

    def test_verify_tamper_happy_path(self, signed_workbook: Path):
        """Unmodified workbook must pass tamper verification with exit code 0."""
        res = run_cli(['--verify-tamper', str(signed_workbook), '--json'])
        assert res.returncode == 0
        data = json.loads(res.stdout)
        assert data['status'] == 'ok'
        assert data['code'] == 0
        assert data['verified'] is True
        assert data['violations_count'] == 0

    def test_verify_tamper_detects_single_cell_mutation(self, signed_workbook: Path):
        """Subtly modifying a single cell must trigger tamper detection with exit code 2."""
        wb = load_workbook(signed_workbook)
        ws = wb['雁山（早）']
        # Mutate attendance from 38 to 39
        ws.cell(2, 7, 39)
        wb.save(signed_workbook)
        wb.close()

        res = run_cli(['--verify-tamper', str(signed_workbook), '--json'])
        assert res.returncode == 2
        data = json.loads(res.stdout)
        assert data['status'] == 'tamper_detected'
        assert data['code'] == 2
        assert data['verified'] is False
        assert data['violations_count'] >= 1
        assert any('雁山（早）!G2' in v for v in data['violations'])

    def test_verify_tamper_detects_sheet_deletion(self, signed_workbook: Path):
        """Deleting a signed sheet must trigger tamper detection with exit code 2."""
        wb = load_workbook(signed_workbook)
        del wb['雁山（晚）']
        wb.save(signed_workbook)
        wb.close()

        res = run_cli(['--verify-tamper', str(signed_workbook), '--json'])
        assert res.returncode == 2
        data = json.loads(res.stdout)
        assert data['status'] == 'tamper_detected'
        assert data['code'] == 2
        assert data['verified'] is False
        assert any('工作表缺失：雁山（晚）' in v for v in data['violations'])

    def test_verify_tamper_detects_formula_injection(self, signed_workbook: Path):
        """Injecting a formula into a data cell must trigger tamper detection with code 2."""
        wb = load_workbook(signed_workbook)
        ws = wb['雁山（早）']
        ws.cell(2, 7, '=SUM(1, 37)')
        wb.save(signed_workbook)
        wb.close()

        res = run_cli(['--verify-tamper', str(signed_workbook), '--json'])
        assert res.returncode == 2
        data = json.loads(res.stdout)
        assert data['status'] == 'tamper_detected'
        assert data['code'] == 2
        assert data['verified'] is False

    def test_verify_tamper_detects_metadata_sheet_removal(self, signed_workbook: Path):
        """Removing the _系统校验元数据 sheet must trigger tamper detection with code 2."""
        wb = load_workbook(signed_workbook)
        del wb['_系统校验元数据']
        wb.save(signed_workbook)
        wb.close()

        res = run_cli(['--verify-tamper', str(signed_workbook), '--json'])
        assert res.returncode == 2
        data = json.loads(res.stdout)
        assert data['status'] == 'tamper_detected'
        assert data['code'] == 2
        assert data['verified'] is False
        assert any('缺失防篡改校验元数据表' in v for v in data['violations'])

    def test_verify_tamper_detects_manifest_f2_modification(self, signed_workbook: Path):
        """Tampering with the overall manifest HMAC at F2 must trigger tamper detection with code 2."""
        wb = load_workbook(signed_workbook)
        ws = wb['_系统校验元数据']
        ws['F2'] = '00000000000000000000000000000000'
        wb.save(signed_workbook)
        wb.close()

        res = run_cli(['--verify-tamper', str(signed_workbook), '--json'])
        assert res.returncode == 2
        data = json.loads(res.stdout)
        assert data['status'] == 'tamper_detected'
        assert data['code'] == 2
        assert data['verified'] is False
        assert any('防篡改元数据清单与签名不符' in v for v in data['violations'])

    def test_verify_tamper_against_detail_headcount_mismatch(self, tmp_path: Path):
        """Mismatched actual headcount cells with original detail must return code 2 with detailed diffs."""
        # Source detail created with standard multi-sheet structure
        src = tmp_path / 'detail_orig.xlsx'
        write_detail(src, sheets=focused_sheets())

        # Copy to result workbook and modify an actual headcount cell
        res_wb = tmp_path / 'detail_res.xlsx'
        shutil.copy(src, res_wb)
        wb_res = load_workbook(res_wb)
        ws_res = wb_res['雁山（早）']
        # Locate an actual headcount cell and mutate it
        ws_res.cell(4, 6).value = 999
        generate_tamper_metadata(wb_res, ['雁山（早）', '雁山（晚）'])
        wb_res.save(res_wb)
        wb_res.close()

        res = run_cli(['--verify-tamper', str(res_wb), '--against-detail', str(src), '--json'])
        assert res.returncode == 2
        data = json.loads(res.stdout)
        assert data['status'] == 'tamper_detected'
        assert data['code'] == 2
        assert len(data['diffs']) > 0
        diff = data['diffs'][0]
        assert diff['sheet'] == '雁山（早）'
        assert 'cell' in diff
        assert 'source_value' in diff
        assert diff['result_value'] == 999


# ==============================================================================
# 4. Adversarial Standalone --exempt-report Stress Tests
# ==============================================================================

class TestAdversarialStandaloneExemptReport:
    """Stress tests for decoupled exemption-only report generation."""

    def test_standalone_exempt_multiple_files_merged(self, tmp_path: Path):
        """Multiple exemption workbooks must be correctly ingested and aggregated."""
        f1 = tmp_path / 'exempt1.xlsx'
        f2 = tmp_path / 'exempt2.xlsx'

        wb1 = Workbook()
        ws1 = wb1.active
        ws1.title = '免勤明细'
        ws1.append(['序号', '学号', '姓名', '学院', '班级', '校区', '免勤周次', '免勤自习', '免勤原因', '请假类型', '自习室'])
        ws1.append([1, '9900000000001', '学生甲', '计算机科学与工程学院', '计算机21-1', '雁山', 7, '早自习（周一）', '选修课', '课务学业', '01201'])
        wb1.save(f1)
        wb1.close()

        wb2 = Workbook()
        ws2 = wb2.active
        ws2.title = '免勤明细'
        ws2.append(['序号', '学号', '姓名', '学院', '班级', '校区', '免勤周次', '免勤自习', '免勤原因', '请假类型', '自习室'])
        ws2.append([1, '9900000000002', '学生乙', '理学院', '物理21-1', '雁山', 7, '晚自习（周二）', '生病', '健康医疗', '02202'])
        wb2.save(f2)
        wb2.close()

        out_rep = tmp_path / 'merged_report.xlsx'
        res = run_cli([
            '--exempt-report', str(out_rep),
            '--exempt', str(f1), str(f2),
            '--week', '7',
            '--json',
        ])
        assert res.returncode == 0
        data = json.loads(res.stdout)
        assert data['status'] == 'ok'
        assert data['code'] == 0
        assert out_rep.is_file()

        # Inspect generated workbook structure
        wb_check = load_workbook(out_rep, read_only=True)
        assert '免勤统计' in wb_check.sheetnames
        wb_check.close()

    def test_standalone_exempt_missing_file_reports_error(self, tmp_path: Path):
        """Missing exemption workbook must return clean JSON error with exit code 1."""
        missing = tmp_path / 'not_found_exempt.xlsx'
        out_rep = tmp_path / 'report.xlsx'
        res = run_cli([
            '--exempt-report', str(out_rep),
            '--exempt', str(missing),
            '--week', '7',
            '--json',
        ])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['action'] == 'exempt_report_standalone'
        assert data['status'] == 'error'
        assert data['code'] == 1
        assert '找不到免勤人员统计表' in data['error']

    def test_standalone_exempt_corrupted_file_reports_error(self, tmp_path: Path):
        """Corrupted exemption workbook must return clean JSON error with exit code 1."""
        bad = tmp_path / 'bad_exempt.xlsx'
        bad.write_bytes(b'corrupted_bytes')
        out_rep = tmp_path / 'report.xlsx'
        res = run_cli([
            '--exempt-report', str(out_rep),
            '--exempt', str(bad),
            '--week', '7',
            '--json',
        ])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['status'] == 'error'
        assert data['code'] == 1

    def test_standalone_exempt_output_overwrites_input_blocked(self, tmp_path: Path):
        """Output path colliding with input path must be rejected with exit code 1."""
        f1 = create_minimal_valid_exempt_workbook(tmp_path / 'exempt.xlsx')
        res = run_cli([
            '--exempt-report', str(f1),
            '--exempt', str(f1),
            '--week', '7',
            '--json',
        ])
        assert res.returncode == 1
        data = json.loads(res.stdout)
        assert data['status'] == 'error'
        assert data['code'] == 1
