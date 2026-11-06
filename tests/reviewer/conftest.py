# -*- coding: utf-8 -*-
"""测试审查员专属装置（tests/reviewer/）。

刻意与 tests/conftest.py 解耦：即使上层 conftest 被其他同事改动，
本套件仍可独立运行 ``python -m pytest tests/reviewer -q``。
"""

from __future__ import annotations

import glob
import os
import sys
from pathlib import Path

import pytest

PROJECT_DIR = Path(__file__).resolve().parents[2]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from hypothesis import HealthCheck, settings          # noqa: E402

settings.register_profile(
    'rj-reviewer',
    max_examples=int(os.environ.get('RJ_HYPOTHESIS_EXAMPLES', '40')),
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow, HealthCheck.function_scoped_fixture],
)
settings.load_profile('rj-reviewer')


def pytest_configure(config):
    config.addinivalue_line(
        'markers', 'realdata: 需要真实工作簿的回归用例，文件缺失时自动跳过')
    config.addinivalue_line(
        'markers', 'spec_gap: 已知与既定口径不一致的现状（xfail 登记，修复后必须去掉标记）')


def _find_real(pattern: str):
    hits = sorted(glob.glob(str(PROJECT_DIR.parent / pattern)))
    return hits[0] if hits else None


@pytest.fixture(scope='session')
def synthetic_detail(tmp_path_factory) -> str:
    import synthetic
    path = tmp_path_factory.mktemp('reviewer') / '合成明细表.xlsx'
    synthetic.build_detail_workbook(str(path))
    return str(path)


@pytest.fixture(scope='session')
def synthetic_detail_broken(tmp_path_factory) -> str:
    """带模板缺陷的合成明细表（N 列漏合并、O~R 行块不齐）。"""
    import synthetic
    path = tmp_path_factory.mktemp('reviewer') / '合成明细表_缺陷.xlsx'
    synthetic.build_detail_workbook(str(path), broken_merge=True)
    return str(path)


@pytest.fixture(scope='session')
def synthetic_exempt(tmp_path_factory) -> str:
    import synthetic
    path = tmp_path_factory.mktemp('reviewer') / '合成免勤表.xlsx'
    early, late = synthetic.default_exempt_rows()
    synthetic.build_exempt_workbook(str(path), early, late)
    return str(path)


@pytest.fixture(scope='session')
def synthetic_run(synthetic_detail, synthetic_exempt, tmp_path_factory):
    """对合成数据跑一次完整流水线，返回 (RunResult, 输出文件路径)。"""
    from core import RunConfig, run
    out = tmp_path_factory.mktemp('reviewer-out') / '合成输出.xlsx'
    result = run(RunConfig(detail_path=synthetic_detail,
                            exempt_path=synthetic_exempt,
                            output_path=str(out),
                            week_no=7))
    assert result.ok, '合成数据流水线必须成功：%s' % result.error
    return result, str(out)


@pytest.fixture(scope='session')
def real_detail_path() -> str:
    path = os.environ.get('RJ_DETAIL_XLSX') or _find_real('*明细表.xlsx')
    if not path or not os.path.exists(path):
        pytest.skip('真实明细表不在预期位置，跳过 realdata 用例')
    return path


@pytest.fixture(scope='session')
def real_exempt_path() -> str:
    path = os.environ.get('RJ_EXEMPT_XLSX') or _find_real('*免勤人员统计表*.xlsx')
    if not path or not os.path.exists(path):
        pytest.skip('真实免勤表不在预期位置，跳过 realdata 用例')
    return path
