# -*- coding: utf-8 -*-
"""pytest 公共装置。

本套件的三条硬纪律：

  1. **测试数据只来自合成数据生成器**——纯虚构，无真实学号/姓名/学院名；
  2. **不依赖项目外的真实 Excel**：真实表含学生个人信息，测试既不读也不复制；
  3. 夹具实现集中在 ``tests/_helpers.py``，本文件只做再导出与全局设置，
     这样即使上层 conftest 被改动，各测试模块仍能独立跑通。

夹具实现见 ``tests/_helpers.py``（``synthetic_dir`` / ``detail_path`` /
``exempt_path`` / ``run_result`` 等）。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from hypothesis import HealthCheck, settings          # noqa: E402

settings.register_profile(
    'rj',
    max_examples=int(os.environ.get('RJ_HYPOTHESIS_EXAMPLES', '40')),
    deadline=None,
    suppress_health_check=[HealthCheck.too_slow, HealthCheck.function_scoped_fixture],
)
settings.load_profile('rj')

# 再导出夹具：本套件要求的夹具都在这里提供。
from _helpers import (                                    # noqa: E402,F401
    build_focused_sample,
    detail_path,
    exempt_path,
    focused_sample,
    focused_sheets,
    generator,
    gen_synthetic,
    make_detail_sheet,
    make_exempt_row,
    md5,
    output_path,
    output_workbook,
    run_engine,
    run_result,
    synthetic_dir,
    workbook_digest,
    write_detail,
    write_exempt,
)


@pytest.fixture(scope='session')
def project_dir() -> Path:
    """项目根目录（测试里推导路径时使用）。"""
    return PROJECT_DIR


def pytest_configure(config):
    config.addinivalue_line(
        'markers', 'realdata: 需要真实工作簿的回归用例（本套件不使用）')
