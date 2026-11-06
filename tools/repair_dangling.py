# -*- coding: utf-8 -*-
"""清理 xfail 装饰器被摘掉后残留的悬空参数行（直接按行号区间删除）。"""

from __future__ import annotations

import io
import os
import re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGETS = [
    'tests/test_exemptions.py',
    'tests/reviewer/test_desktop_gui.py',
    'tests/reviewer/test_dedup_and_group.py',
    'tests/reviewer/test_allocation_and_rate.py',
    'tests/reviewer/test_week_enumeration.py',
]


def is_dangle(line: str) -> bool:
    s = line.strip()
    return s.startswith("reason=") or s.startswith("reason =") or s.startswith("'")


def repair(path: str) -> int:
    lines = io.open(path, encoding='utf-8').read().split('\n')
    drop = set()
    for i, line in enumerate(lines):
        if not is_dangle(line):
            continue
        # 收集这一段连续的悬空行
        j = i
        while j + 1 < len(lines) and is_dangle(lines[j + 1]):
            j += 1
        nxt = lines[j + 1] if j + 1 < len(lines) else ''
        if nxt.startswith('def ') or nxt.startswith('@'):
            for k in range(i, j + 1):
                drop.add(k)
    if drop:
        out = [l for i, l in enumerate(lines) if i not in drop]
        io.open(path, 'w', encoding='utf-8', newline='').write('\n'.join(out))
    return len(drop)


def main() -> int:
    for rel in TARGETS:
        path = os.path.join(ROOT, rel)
        if not os.path.exists(path):
            print('跳过：%s' % rel)
            continue
        n = repair(path)
        print('%s：删除 %d 行' % (rel, n))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
