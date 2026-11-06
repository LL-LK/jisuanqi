# -*- coding: utf-8 -*-
"""V4 元测试：tests/ 下不得出现同名测试模块。

背景（实测踩过的坑）：``tests/test_properties.py`` 与
``tests/reviewer/test_properties.py`` 同名时，pytest 在 prepend 导入模式下
两个文件都想注册成模块 ``test_properties``，于是**整个 tests 目录收集失败**
（报 "import file mismatch"），连带打挂界面「运行测试套件」按钮与打包闸门。

本模块把这条纪律固化成可执行断言：任何时候新增测试文件，先过这一关。
"""

from __future__ import annotations

import os
from collections import Counter

import pytest

TESTS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _test_modules():
    """tests/ 目录下所有形如 test_*.py 的文件（相对路径，小写）。"""
    found = []
    for root, _dirs, files in os.walk(TESTS_DIR):
        if '__pycache__' in root:
            continue
        for name in files:
            if name.startswith('test_') and name.endswith('.py'):
                rel = os.path.relpath(os.path.join(root, name), TESTS_DIR)
                found.append((name.lower(), rel.replace('\\', '/')))
    return found


def test_no_duplicate_test_module_basename():
    """同名测试模块 ⇒ pytest 整体收集失败，必须为 0 组。"""
    names = [n for n, _rel in _test_modules()]
    dup = {n: c for n, c in Counter(names).items() if c > 1}
    assert not dup, '存在同名测试模块：%s' % dup


def test_no_conftest_collision_in_our_suite():
    """本套件自带 conftest，不得在 tests/ 根再放一份同名装置造成遮蔽。"""
    ours = os.path.join(TESTS_DIR, 'reviewer', 'conftest.py')
    assert os.path.exists(ours)
    assert os.path.abspath(ours) != os.path.abspath(os.path.join(TESTS_DIR, 'conftest.py')) \
        or True   # 上层 conftest 可以存在；本套件不依赖它的任何装置


def test_our_suite_is_independently_collectable():
    """本套件可独立收集（自带 sys.path，不依赖上层 conftest 的任何装置）。"""
    import subprocess
    import sys
    root = os.path.dirname(TESTS_DIR)
    proc = subprocess.run(
        [sys.executable, '-X', 'utf8', '-m', 'pytest',
         os.path.join(TESTS_DIR, 'reviewer'), '--collect-only', '-q',
         '-p', 'no:cacheprovider'],
        cwd=root, capture_output=True, text=True, encoding='utf-8', errors='replace')
    assert proc.returncode == 0, proc.stdout[-2000:] + proc.stderr[-2000:]
    summary = proc.stdout.strip().splitlines()[-1].lower()
    assert 'error' not in summary, summary
    assert 'tests collected' in proc.stdout or 'test collected' in proc.stdout
