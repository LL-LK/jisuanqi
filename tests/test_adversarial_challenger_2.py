# -*- coding: utf-8 -*-
"""Adversarial stress testing suite for Skill scripts and dictionary security.

Conducted by Challenger 2 (Adversarial Stress Tester).
Scope:
1. scripts/run_calc.py Zero-PII Defense:
   - Calculations with inputs containing diverse student names and 13-digit IDs.
   - Output verification under error conditions (invalid args, missing sheets, div/0, corrupted cells).
   - Validation that internal regex assertion triggers and blocks execution on PII leaks.
2. scripts/dict_tool.py Security:
   - Strict rejection of unauthorized categories in 免勤理由分类.
   - Formula injection protection (=1+1, =HYPERLINK, =cmd|...).
   - Auto-backup exact snapshot verification and complete recovery.
3. NTFS Junction Link Integrity:
   - Bidirectional modification reflection between release dir and global skill dir.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple

import pytest
from openpyxl import Workbook, load_workbook

from core.dictionary import (
    DEFAULT_CAMPUS_ALIASES,
    DEFAULT_COLLEGE_ALIASES,
    DEFAULT_REASON_CATEGORIES,
    HEADERS,
    SHEET_NAMES,
    Dictionary,
    backup_dictionary,
    load_tables,
    save_tables,
)
from tests._helpers import (
    focused_sheets,
    make_exempt_row,
    write_detail,
    write_exempt,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SID_REGEX = re.compile(r'(?<![0-9A-Fa-f])\d{13}(?![0-9A-Fa-f])')


def find_skill_scripts() -> Tuple[Path, Path]:
    """Dynamically locate run_calc.py and dict_tool.py."""
    candidates = [
        Path.home() / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts',
        REPO_ROOT.parents[1] / '三早一晚免勤自动统计' / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts',
    ]
    for c in candidates:
        rc = c / 'run_calc.py'
        dt = c / 'dict_tool.py'
        if rc.is_file() and dt.is_file():
            return rc, dt
    raise FileNotFoundError("Could not find attendance-stats-calc scripts directory")


def create_adversarial_workbooks(tmp_dir: Path) -> Dict[str, Path]:
    """Create test workbooks containing synthetic realistic student names and 13-digit SIDs."""
    # Diverse synthetic students with realistic naming patterns
    students = [
        ('9900000000001', '张伟测试', '理学院', '应用统计学26-1班', '生病就医住院', '雁山校区', '周一', '早', '8101'),
        ('9900000000002', '李强测试', '理学院', '应用统计学26-1班', '急性胃肠炎急诊', '雁山校区', '周一', '早', '8101'),
        ('9900000000003', '欧阳夏丹测试', '理学院', '应用统计学26-1班', '骨折治疗复查', '雁山校区', '周一', '早', '8101'),
        ('9900000000004', '买买提·艾力测试', '理学院', '应用统计学26-1班', '大学物理实验重修冲突', '雁山校区', '周一', '早', '8101'),
    ]

    # 1. Compliant Detail Workbook
    detail_path = tmp_dir / 'adversarial_detail.xlsx'
    write_detail(detail_path, focused_sheets())

    # 2. Exempt Workbook with diverse students
    exempt_rows = [
        make_exempt_row(
            sid=sid,
            college=college,
            name=name,
            class_text=class_text,
            week_text='第7周',
            day_text=day,
            session=session,
            campus=campus,
            reason=reason,
            room=room,
        )
        for sid, name, college, class_text, reason, campus, day, session, room in students
    ]
    exempt_path = tmp_dir / 'adversarial_exempt.xlsx'
    write_exempt(
        exempt_path,
        {
            '早自习免勤人员统计表': exempt_rows,
            '晚自习免勤人员统计表': [],
        },
        with_room=True,
    )

    # 3. Standard Dictionary
    dict_path = tmp_dir / 'adversarial_dict.xlsx'
    tables = {
        '学院别名': [['理学院', '理院', '标准测试别名']],
        '班级别名': [],
        '校区别名': [list(r) for r in DEFAULT_CAMPUS_ALIASES],
        '免勤理由分类': [
            ['生病就医住院', '健康医疗', '住院'],
            ['急性胃肠炎急诊', '健康医疗', '急诊'],
            ['骨折治疗复查', '健康医疗', '骨折治疗'],
            ['大学物理实验重修冲突', '课务学业', '实验冲突'],
            ['生病', '健康医疗', '通用健康情形词'],
            ['有选修课', '课务学业', '通用课程'],
        ],
        '统计参数': [
            ['统计周次', '7', '周次'],
            ['计入校区', '雁山', '校区'],
            ['空白校区默认值', '雁山', '默认'],
            ['出勤率完整计算模式', '0', '模式'],
            ['估算标记固定口径', 'FFFFF3CD', '颜色'],
        ],
    }
    save_tables(str(dict_path), tables, backup=False)

    return {
        'detail': detail_path,
        'exempt': exempt_path,
        'dict': dict_path,
        'students': students,
    }


class TestRunCalcZeroPIIDefense:
    """Adversarially probe run_calc.py Zero-PII defenses."""

    @pytest.fixture(autouse=True)
    def setup(self):
        self.run_calc, self.dict_tool = find_skill_scripts()

    def test_run_calc_realworld_pii_no_leak(self, tmp_path: Path):
        """Test calculation with diverse real-world style names and 13-digit SIDs."""
        data = create_adversarial_workbooks(tmp_path)
        out_file = tmp_path / 'calc_out.xlsx'
        cmd = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(data['detail']),
            '--exempt', str(data['exempt']),
            '--dictionary', str(data['dict']),
            '--out', str(out_file),
            '--week', '7',
            '--app-root', str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 0, f"run_calc execution failed: {res.stderr}\nstdout: {res.stdout}"

        # 1. Adversarially verify stdout: 0 student IDs
        sid_match = SID_REGEX.search(res.stdout)
        assert sid_match is None, f"Leak: 13-digit SID found in stdout: {sid_match.group(0)}"

        # 2. Adversarially verify stdout: 0 student names
        for item in data['students']:
            sid, name = item[0], item[1]
            assert sid not in res.stdout, f"Leak: SID {sid} leaked in stdout"
            assert name not in res.stdout, f"Leak: Student name {name} leaked in stdout"

        # 3. Adversarially verify stderr
        sid_err_match = SID_REGEX.search(res.stderr)
        assert sid_err_match is None, f"Leak: 13-digit SID found in stderr: {sid_err_match.group(0)}"

        # 4. JSON contract verification
        payload = json.loads(res.stdout)
        assert payload['action'] == 'run'
        assert payload['status'] == 'success'
        assert payload['code'] == 0
        assert 'college_rates' in payload
        assert 'weekly_rates' in payload
        assert 'anomalies_count' in payload
        assert 'anomaly_summary' in payload
        # Ensure raw anomalies list is NOT in output
        assert 'anomalies' not in payload

    def test_run_calc_error_missing_sheets_no_pii_leak(self, tmp_path: Path):
        """Test error when detail file has missing required sheets."""
        data = create_adversarial_workbooks(tmp_path)
        # Create invalid detail file without '雁山（早）'
        bad_detail = tmp_path / 'missing_sheets_detail.xlsx'
        wb = Workbook()
        ws = wb.active
        ws.title = 'OtherSheet'
        ws.append(['一些测试数据'])
        wb.save(bad_detail)
        wb.close()

        cmd = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(bad_detail),
            '--exempt', str(data['exempt']),
            '--dictionary', str(data['dict']),
            '--week', '7',
            '--app-root', str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 1

        # Check no PII leaked in stdout or stderr
        assert SID_REGEX.search(res.stdout) is None
        assert SID_REGEX.search(res.stderr) is None
        for item in data['students']:
            assert item[0] not in res.stdout
            assert item[1] not in res.stdout
            assert item[0] not in res.stderr
            assert item[1] not in res.stderr

    def test_run_calc_error_corrupted_cells_no_pii_leak(self, tmp_path: Path):
        """Test calculation with corrupted cells (#DIV/0!, formulas, non-numeric headcounts)."""
        data = create_adversarial_workbooks(tmp_path)
        # Inject corrupted cells into detail workbook
        wb = load_workbook(data['detail'])
        ws = wb['雁山（早）']
        # Set expected headcount to invalid text
        ws.cell(4, 5, value='#VALUE!')
        # Set actual headcount to formula error
        ws.cell(4, 6, value='#DIV/0!')
        corrupt_detail = tmp_path / 'corrupt_cells_detail.xlsx'
        wb.save(corrupt_detail)
        wb.close()

        cmd = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(corrupt_detail),
            '--exempt', str(data['exempt']),
            '--dictionary', str(data['dict']),
            '--out', str(tmp_path / 'corrupt_out.xlsx'),
            '--week', '7',
            '--app-root', str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')

        # Check no PII leaked in stdout or stderr
        assert SID_REGEX.search(res.stdout) is None
        assert SID_REGEX.search(res.stderr) is None
        for item in data['students']:
            assert item[0] not in res.stdout
            assert item[1] not in res.stdout
            assert item[0] not in res.stderr
            assert item[1] not in res.stderr

    def test_run_calc_error_division_by_zero_handling(self, tmp_path: Path):
        """Test calculation when expected headcount is 0 or all exempted (denominator 0)."""
        data = create_adversarial_workbooks(tmp_path)
        wb = load_workbook(data['detail'])
        ws = wb['雁山（早）']
        ws.cell(4, 5, value=0)  # expected headcount = 0
        ws.cell(4, 6, value=0)  # actual headcount = 0
        div0_detail = tmp_path / 'div0_detail.xlsx'
        wb.save(div0_detail)
        wb.close()

        cmd = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(div0_detail),
            '--exempt', str(data['exempt']),
            '--dictionary', str(data['dict']),
            '--out', str(tmp_path / 'div0_out.xlsx'),
            '--week', '7',
            '--app-root', str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        # Regardless of returncode 0 or 1, PII must not leak
        assert SID_REGEX.search(res.stdout) is None
        assert SID_REGEX.search(res.stderr) is None
        for item in data['students']:
            assert item[0] not in res.stdout
            assert item[1] not in res.stdout

    def test_run_calc_error_invalid_arguments_no_pii_leak(self, tmp_path: Path):
        """Test running with completely invalid/missing arguments."""
        cmd = [
            sys.executable,
            str(self.run_calc),
            '--detail', str(tmp_path / 'non_existent_file.xlsx'),
            '--app-root', str(REPO_ROOT),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 1
        assert SID_REGEX.search(res.stdout) is None
        assert SID_REGEX.search(res.stderr) is None

    def test_run_calc_assert_privacy_compliance_regex_triggers(self):
        """Directly verify that assert_privacy_compliance triggers on 13-digit IDs and permits hashes."""
        sys.path.insert(0, str(self.run_calc.parent))
        from run_calc import assert_privacy_compliance

        # 1. 13-digit standalone ID -> MUST RAISE
        with pytest.raises(ValueError, match="Privacy assertion violation"):
            assert_privacy_compliance("Student ID 9923101010001 detected")

        # 2. 13-digit in JSON -> MUST RAISE
        with pytest.raises(ValueError, match="Privacy assertion violation"):
            assert_privacy_compliance(json.dumps({"student": "9923101010001"}))

        # 3. 13-digit surrounded by Chinese characters -> MUST RAISE
        with pytest.raises(ValueError, match="Privacy assertion violation"):
            assert_privacy_compliance("学号9923101010001存在异常")

        # 4. Safe strings -> MUST PASS
        assert_privacy_compliance("Action completed successfully. Code: 0, Total: 15")

        # 5. Hexadecimal SHA-256 containing 13 consecutive digits -> MUST NOT FALSE-POSITIVE
        hex_with_13_digits = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        assert_privacy_compliance(f"sha256: {hex_with_13_digits}")


class TestDictToolSecurity:
    """Adversarially probe dict_tool.py security mechanisms."""

    @pytest.fixture(autouse=True)
    def setup(self):
        self.run_calc, self.dict_tool = find_skill_scripts()

    def test_dict_tool_rejects_unauthorized_categories(self, tmp_path: Path):
        """Test strict rejection (exit code 1) of unauthorized reason categories."""
        data = create_adversarial_workbooks(tmp_path)
        dict_copy = tmp_path / 'dict_test_categories.xlsx'
        shutil.copy2(data['dict'], dict_copy)

        unauthorized_categories = [
            ('请假私事', '私事请假'),
            ('看望亲人', '探亲'),
            ('校外兼职', '打工'),
            ('创业实践', '创业'),
            ('外出旅游', '个人旅游'),
        ]

        for alias, bad_category in unauthorized_categories:
            cmd = [
                sys.executable,
                str(self.dict_tool),
                '--app-root', str(REPO_ROOT),
                'add-alias',
                '--sheet', '免勤理由分类',
                '--alias', alias,
                '--target', bad_category,
                '--dict', str(dict_copy),
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            assert res.returncode == 1, f"Failed to reject unauthorized category '{bad_category}'"
            assert "Whitelist violation" in res.stderr or "Whitelist violation" in res.stdout

    def test_dict_tool_check_catches_tampered_unauthorized_categories(self, tmp_path: Path):
        """Test dict_tool check flags unauthorized categories injected into file."""
        data = create_adversarial_workbooks(tmp_path)
        dict_copy = tmp_path / 'dict_tampered_cat.xlsx'
        shutil.copy2(data['dict'], dict_copy)

        # Manually inject unauthorized category into workbook
        wb = load_workbook(dict_copy)
        ws = wb['免勤理由分类']
        ws.append(['兼职送外卖', '打工', '非白名单类别'])
        wb.save(dict_copy)
        wb.close()

        # Run check
        cmd = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'check',
            '--dict', str(dict_copy),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 1
        assert "Whitelist violation" in res.stderr or "Whitelist violation" in res.stdout
        assert "打工" in res.stderr or "打工" in res.stdout

    def test_dict_tool_formula_injection_defense(self, tmp_path: Path):
        """Test formula injection defense (=1+1, =HYPERLINK, =cmd|...)."""
        data = create_adversarial_workbooks(tmp_path)
        dict_copy = tmp_path / 'dict_formula_injection.xlsx'
        shutil.copy2(data['dict'], dict_copy)

        formula_payloads = [
            ("=1+1", "理学院"),
            ("=HYPERLINK(\"http://malicious.com\",\"Click\")", "理学院"),
            ("=cmd|' /C calc'!A0", "理学院"),
        ]

        for payload, target in formula_payloads:
            cmd = [
                sys.executable,
                str(self.dict_tool),
                '--app-root', str(REPO_ROOT),
                'add-alias',
                '--sheet', '学院别名',
                '--alias', payload,
                '--target', target,
                '--dict', str(dict_copy),
            ]
            res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            assert res.returncode == 0, f"add-alias command failed: {res.stderr}"

            # Inspect the saved workbook
            wb = load_workbook(dict_copy, data_only=False)
            ws = wb['学院别名']
            found = False
            for row in ws.iter_rows(min_row=2):
                cell_val = row[1].value
                if cell_val == payload:
                    found = True
                    # Verify that cell data_type is strictly string 's', NOT formula 'f'
                    assert row[1].data_type == 's', f"Formula injection vulnerability: cell {row[1].coordinate} has data_type '{row[1].data_type}'"
            wb.close()
            assert found, f"Injected row not found for payload {payload}"

            # Verify that load_tables(strict=True) does not crash or execute formulas
            tables = load_tables(str(dict_copy), strict=True)
            assert '学院别名' in tables

    def test_dict_tool_auto_backup_exact_snapshot_and_recovery(self, tmp_path: Path):
        """Test auto-backup exact SHA256 snapshot prior to writing and full recovery."""
        data = create_adversarial_workbooks(tmp_path)
        dict_target = tmp_path / 'dict_for_backup_test.xlsx'
        shutil.copy2(data['dict'], dict_target)

        # 1. Compute initial SHA256
        with open(dict_target, 'rb') as f:
            original_sha256 = hashlib.sha256(f.read()).hexdigest()

        # 2. Perform safe modification via dict_tool add-alias
        cmd = [
            sys.executable,
            str(self.dict_tool),
            '--app-root', str(REPO_ROOT),
            'add-alias',
            '--sheet', '学院别名',
            '--alias', '理学新别名',
            '--target', '理学院',
            '--dict', str(dict_target),
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
        assert res.returncode == 0
        assert '[BACKUP]' in res.stdout

        # 3. Locate backup file
        prefix = '[BACKUP] Created safety backup:'
        assert prefix in res.stdout, "Backup path not printed in stdout"
        backup_line = [line for line in res.stdout.splitlines() if prefix in line][0]
        backup_file = Path(backup_line.split(prefix, 1)[1].strip())
        assert backup_file.is_file(), f"Backup file {backup_file} does not exist"

        # 4. Verify backup is an exact byte-for-byte snapshot of original file
        with open(backup_file, 'rb') as f:
            backup_sha256 = hashlib.sha256(f.read()).hexdigest()
        assert backup_sha256 == original_sha256, "Backup snapshot does not match pre-modification file!"

        # 5. Verify recovery: corrupt/damage the active dictionary
        with open(dict_target, 'wb') as f:
            f.write(b'CORRUPTED_GARBAGE_DATA')

        # Restore from backup
        shutil.copy2(backup_file, dict_target)
        with open(dict_target, 'rb') as f:
            restored_sha256 = hashlib.sha256(f.read()).hexdigest()
        assert restored_sha256 == original_sha256, "Restoration from backup failed!"

        # Verify restored workbook loads cleanly
        dict_obj = Dictionary(load_tables(str(dict_target), strict=True))
        assert dict_obj.exact_college('理学院') == '理学院'


class TestNTFSJunctionLinkIntegrity:
    """Verify bidirectional integrity of the NTFS Junction link."""

    def test_junction_bidirectional_sync_and_integrity(self):
        """Test write from Release -> Global, modify from Global -> Release, and verify single-source sync."""
        release_skill_dir = REPO_ROOT.parents[1] / '三早一晚免勤自动统计' / '.agents' / 'skills' / 'attendance-stats-calc'
        global_skill_dir = Path.home() / '.agents' / 'skills' / 'attendance-stats-calc'

        if not release_skill_dir.is_dir() or not global_skill_dir.is_dir():
            pytest.skip("Release or Global skill directory not accessible in this environment")

        probe_release = release_skill_dir / '.challenger2_junction_probe.tmp'
        probe_global = global_skill_dir / '.challenger2_junction_probe.tmp'

        try:
            # 1. Write from Release Junction
            probe_release.write_text("challenger2_probe_data_1", encoding='utf-8')
            assert probe_global.is_file(), "File written via Junction does not reflect in Global directory!"
            assert probe_global.read_text(encoding='utf-8') == "challenger2_probe_data_1"

            # 2. Modify from Global Target
            probe_global.write_text("challenger2_probe_data_2", encoding='utf-8')
            assert probe_release.read_text(encoding='utf-8') == "challenger2_probe_data_2", "Modification in Global target does not reflect in Junction!"
        finally:
            # Clean up
            if probe_release.exists():
                probe_release.unlink()
            assert not probe_global.exists(), "Cleanup failed to reflect across junction!"
