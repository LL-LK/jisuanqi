# -*- coding: utf-8 -*-
"""模型零隐私剔除模块与伪匿名化等价性回归测试。"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parent.parent
SKILL_SCRIPT = Path.home() / '.agents' / 'skills' / 'attendance-stats-calc' / 'scripts' / 'run_calc.py'
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import cli  # noqa: E402
from core import RunConfig, run  # noqa: E402
from core.excel_safety import generate_diagnostic_log  # noqa: E402
from core.privacy import (  # noqa: E402
    ANON_SID_RE,
    PrivacyRedactor,
    PrivacyViolationError,
    assert_model_privacy_compliance,
    build_model_safe_exempt_summary,
    strip_workbook_privacy,
)
from tests_e2e.helpers import create_detail_workbook  # noqa: E402


def _syn_sid(suffix: str) -> str:
    return '2024' + suffix


_SID_DEMO = _syn_sid('081901234')
_SID_1 = _syn_sid('081901001')
_SID_2 = _syn_sid('081902001')
_SID_3 = _syn_sid('081903001')


class TestPrivacyRedactorCore:
    def test_anonymize_sid_is_deterministic_and_irreversible(self):
        red = PrivacyRedactor(salt=b'unit-test-fixed-salt-32bytes-000')
        raw_sid = _SID_DEMO
        anon1 = red.anonymize_sid(raw_sid, valid_sid=True)
        anon2 = red.anonymize_sid(raw_sid, valid_sid=True)
        assert anon1 == anon2
        assert ANON_SID_RE.match(anon1)
        assert '2024' not in anon1
        assert '1234' not in anon1

        # 异常位数与 NOSID 合成键同样不透出原文
        abnorm = red.anonymize_sid('20240819', valid_sid=False)
        assert abnorm.startswith('ABNORM-')
        assert '20240819' not in abnorm

        nosid = red.anonymize_sid('NOSID:计算机学院:计科241:张秘密')
        assert ANON_SID_RE.match(nosid)
        assert '张秘密' not in nosid
        assert '计科' not in nosid

    def test_redact_text_removes_names_completely_without_surname_leak(self):
        red = PrivacyRedactor(salt=b'unit-test-fixed-salt-32bytes-001')
        raw = (
            f'同学赵铁柱({_SID_DEMO})于计算机24-1班因急性阑尾炎住院开刀请假，'
            '联系电话13800138000，身份证45030020040101123X，辅导员钱老师核实。'
        )
        cleaned = red.redact_text(
            raw,
            extra_names={'赵铁柱'},
            extra_sids={_SID_DEMO},
            extra_reasons={'急性阑尾炎住院开刀'},
            extra_classes={'计算机24-1班'},
        )
        for forbidden in ('赵铁柱', '赵*', _SID_DEMO, '2024', '1234',
                          '13800138000', '45030020040101123X', '急性阑尾炎', '住院开刀',
                          '计算机24-1班', '钱老师'):
            assert forbidden not in cleaned
        assert '[已脱敏]' in cleaned
        assert 'ANON-' in cleaned
        assert '[健康医疗]' in cleaned

    def test_assert_model_privacy_compliance_blocks_every_pii_type(self):
        assert_model_privacy_compliance('{"anon_sid": "ANON-A1B2C3D4E5", "college": "计算机学院", "exempt": 12}')

        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance(json.dumps({"sid": _SID_DEMO}))

        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance('{"phone": "13912345678"}')

        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance('{"id": "45030020040101123X"}')

        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance('学生赵铁柱未计入', forbidden_names={'赵铁柱'})

        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance('原因：左踝扭伤打石膏', forbidden_reasons={'左踝扭伤打石膏'})

        with pytest.raises(PrivacyViolationError):
            assert_model_privacy_compliance('所属：计算机24-1班', forbidden_classes={'计算机24-1班'})


class TestStripPrivacyEquivalenceAndModelGate:
    @staticmethod
    def _build_sensitive_exempt_workbook(path: Path) -> None:
        wb = Workbook()
        ws = wb.active
        ws.title = '早自习'
        ws.append([
            '学号', '姓名', '学院', '专业班级', '所在教室',
            '免勤周次', '免勤时间', '免勤原因', '手机号', '身份证号', '辅导员备注',
        ])
        # 1. 计算机24-1班个人申请（会被同班集体申请覆盖）
        ws.append([
            _SID_1, '赵铁柱', '计算机学院', '计算机24-1班', '8101',
            '第7周', '周一', '急性扁桃体炎校医院发烧输液', '13911112222', '450300200401011111', '王辅导员已核实',
        ])
        # 2. 计算机24-1班集体申请（共4人，覆盖上面个人）
        ws.append([
            '', '计算机24-1班全体学生（共4人）', '计算机学院', '计算机24-1班', '8101',
            '第7周', '周一', '数学建模竞赛集训共4人', '', '', '李老师带队',
        ])
        # 3. 计算机24-2班个人申请（正常计入 1 人） + 重复申报（应去重仍计 1 人）
        ws.append([
            _SID_2, '钱翠花', '计算机学院', '计算机24-2班', '8102',
            '第7周', '周一', '高数重修课程冲突', '13933334444', '45030020040202222X', '教务系统截图',
        ])
        ws.append([
            _SID_2, '钱翠花', '计算机学院', '计算机24-2班', '8102',
            '第7周', '周一', '高数重修课程冲突', '13933334444', '45030020040202222X', '重复提交',
        ])
        # 4. 理学院个人申请（8201）
        ws.append([
            _SID_3, '孙大壮', '理学院', '理学24-1班', '8201',
            '第7周', '周三', '家庭急事回家处理', '13855556666', '450300200403033333', '家长电话确认',
        ])
        wb.save(path)
        wb.close()

    def test_strip_privacy_removes_all_pii_and_preserves_exact_calculation(self, tmp_path: Path, capsys):
        detail_path = tmp_path / 'detail.xlsx'
        exempt_raw = tmp_path / 'exempt_raw.xlsx'
        exempt_stripped = tmp_path / 'exempt_stripped.xlsx'
        out_raw = tmp_path / 'out_raw.xlsx'
        out_stripped = tmp_path / 'out_stripped.xlsx'

        create_detail_workbook(detail_path, week_no=7)
        self._build_sensitive_exempt_workbook(exempt_raw)

        # CLI --strip-privacy 预剔除
        code = cli.main(['--strip-privacy', str(exempt_raw), '--out', str(exempt_stripped), '--json'])
        assert code == 0
        payload = json.loads(capsys.readouterr().out)
        assert payload['action'] == 'strip_privacy'
        assert payload['status'] == 'ok'
        assert payload['residual_formulas'] == 0
        assert payload['sids_anonymized'] >= 3
        assert payload['names_removed'] >= 4

        # 扫描脱敏后的 Excel 全格内容，确认零真实隐私残留（含半打码姓氏）
        wb = load_workbook(exempt_stripped, data_only=False)
        try:
            all_text = '\n'.join(
                str(cell.value)
                for ws in wb.worksheets
                for row in ws.iter_rows()
                for cell in row
                if cell.value is not None
            )
        finally:
            wb.close()

        for leaked in (
            _SID_1, _SID_2, _SID_3,
            '赵铁柱', '钱翠花', '孙大壮', '赵*', '钱*', '孙*',
            '扁桃体', '发烧输液', '家庭急事', '计算机24-1班', '计算机24-2班', '理学24-1班',
            '13911112222', '13933334444', '13855556666',
            '450300200401011111', '45030020040202222X', '450300200403033333',
            '王辅导员', '李老师带队', '家长电话确认',
        ):
            assert leaked not in all_text, f'脱敏 Excel 仍残留隐私词: {leaked}'
        assert 'ANON-' in all_text
        assert 'CLS-' in all_text

        # 原始明文表与脱敏表分别执行完整计算，对比全部学院出勤率与免勤人数 100% 一致
        res_raw = run(RunConfig(
            detail_path=str(detail_path),
            exempt_path=str(exempt_raw),
            output_path=str(out_raw),
            week_no=7,
            enable_storage=False,
        ))
        res_stripped = run(RunConfig(
            detail_path=str(detail_path),
            exempt_path=str(exempt_stripped),
            output_path=str(out_stripped),
            week_no=7,
            enable_storage=False,
        ))
        assert res_raw.ok and res_stripped.ok

        raw_rates = {(r.sheet, r.college, r.day): (r.expected, r.actual, r.exempt, r.rate)
                     for r in res_raw.college_rates}
        str_rates = {(r.sheet, r.college, r.day): (r.expected, r.actual, r.exempt, r.rate)
                     for r in res_stripped.college_rates}
        assert raw_rates == str_rates

    def test_cli_and_skill_model_output_only_expose_anon_sid_and_college_counts(self, tmp_path: Path, capsys):
        detail_path = tmp_path / 'detail.xlsx'
        exempt_raw = tmp_path / 'exempt_raw.xlsx'
        out_path = tmp_path / 'out.xlsx'
        diag_path = tmp_path / 'diag.md'

        create_detail_workbook(detail_path, week_no=7)
        self._build_sensitive_exempt_workbook(exempt_raw)

        # 1. CLI --json + --diagnostic-log 零隐私检查
        code = cli.main([
            '--detail', str(detail_path),
            '--exempt', str(exempt_raw),
            '--week', '7',
            '--out', str(out_path),
            '--diagnostic-log', str(diag_path),
            '--no-storage',
            '--json',
        ])
        assert code == 0
        cli_out = capsys.readouterr().out
        data = json.loads(cli_out)

        assert 'college_exempt_summary' in data
        colleges = {item['college']: item for item in data['college_exempt_summary'] if item['total_exempt_count'] > 0}
        assert '计算机科学与工程学院' in colleges
        assert colleges['计算机科学与工程学院']['total_exempt_count'] == 5  # 集体4 + 个人1
        assert all(ANON_SID_RE.match(s) for item in data['college_exempt_summary'] for s in item['anon_sids'])

        # 严密确认 JSON 与诊断日志中均无任何姓名（含打码姓氏）、原学号、原班级、原病由、手机号
        combined_text = cli_out + '\n' + diag_path.read_text(encoding='utf-8')
        for secret in (
            '赵铁柱', '钱翠花', '孙大壮', '赵*', '钱*', '孙*',
            _SID_1, _SID_2, _SID_3,
            '扁桃体', '发烧输液', '高数重修', '家庭急事',
            '计算机24-1班', '计算机24-2班', '理学24-1班',
            '13911112222', '450300200401011111', '王辅导员',
        ):
            assert secret not in combined_text, f'模型可见输出出现未剔除隐私: {secret}'

        # 2. Skill wrapper run_calc.py 同样隔离输出
        if SKILL_SCRIPT.exists():
            proc = subprocess.run(
                [
                    sys.executable, str(SKILL_SCRIPT),
                    '--app-root', str(ROOT),
                    '--detail', str(detail_path),
                    '--exempt', str(exempt_raw),
                    '--week', '7',
                    '--out', str(tmp_path / 'out_skill.xlsx'),
                    '--no-storage',
                ],
                capture_output=True,
                text=True,
                encoding='utf-8',
                timeout=30,
            )
            assert proc.returncode == 0, proc.stderr
            skill_data = json.loads(proc.stdout)
            assert 'college_exempt_summary' in skill_data
            for secret in ('赵铁柱', '钱翠花', '孙大壮', _SID_1, '扁桃体', '13911112222'):
                assert secret not in proc.stdout
