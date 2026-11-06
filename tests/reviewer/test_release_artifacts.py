# -*- coding: utf-8 -*-
"""发布产物一致性验收（标记 `releasegate`）。

v1.0.5 入口条件 7：**随包词典与项目根词典 sha256 相同**。
本模块提供独立于 `tools/verify_release.py` 的第二道交叉校验（测试侧自证），
避免"脚本说一致、实际不一致"。

默认不进打包闸门（发布产物在构建时才存在），发版前单独执行：
    python -X utf8 -m pytest tests -m releasegate -q
"""

from __future__ import annotations

import hashlib
import os

import pytest

pytestmark = pytest.mark.releasegate

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
ROOT_DICT = os.path.join(PROJECT_DIR, '免勤词典.xlsx')
PACKAGED_DICT = os.path.join(PROJECT_DIR, '发布版', '免勤词典.xlsx')


def _sha256(path: str) -> str:
    digest = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b''):
            digest.update(chunk)
    return digest.hexdigest()


def test_packaged_dictionary_sha256_matches_project_root():
    """入口条件 7：随包词典与项目根词典必须逐位一致（否则随包落后一个构建）。"""
    if not os.path.exists(ROOT_DICT):
        pytest.skip('项目根词典不存在：%s' % ROOT_DICT)
    if not os.path.exists(PACKAGED_DICT):
        pytest.skip('尚未构建发布包（%s 不存在），发版前需重跑 12 道闸门' % PACKAGED_DICT)
    assert _sha256(ROOT_DICT) == _sha256(PACKAGED_DICT), \
        '随包词典与项目根词典 sha256 不一致 ⇒ 随包落后一个构建'


def test_release_dir_contains_expected_artifacts():
    """发布目录必须齐备：EXE + 词典 + 说明 + 隐私报告 + 校验清单。"""
    if not os.path.isdir(os.path.join(PROJECT_DIR, '发布版')):
        pytest.skip('尚未构建发布包')
    names = set(os.listdir(os.path.join(PROJECT_DIR, '发布版')))
    for required in ('免勤词典.xlsx', '隐私扫描报告.md', 'SHA256.txt'):
        assert required in names, '发布目录缺少 %s' % required
    assert any(n.lower().endswith('.exe') for n in names), '发布目录缺少 EXE'
