# -*- coding: utf-8 -*-
"""把已经修好的用例上的 xfail 标记摘掉（xfail 装饰器位于 def 上方）。

已修复、应当转为普通回归用例的用例见 REMOVE 列表。
保留 xfail 的：test_room_allocation_sum_equals_college_exempt_even_when_overflow
    —— 教室级按容量封顶是**对话口径明文要求**（220>219 锚点），不是缺陷。
"""

from __future__ import annotations

import io
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TESTS = os.path.join(ROOT, 'tests')

REMOVE = {
    'tests/test_week_enumeration.py': [],
    'tests/reviewer/test_week_enumeration.py': ['test_parse_weeks_chinese_numeral_range_without_prefix',
                                               'test_parse_weeks_space_separated_enumeration'],
    'tests/test_exemptions.py': ['test_group_and_personal_of_other_class_both_count'],
    'tests/reviewer/test_dedup_and_group.py': ['test_group_application_without_size_does_not_cover_personal',
                                                'test_group_application_does_not_swallow_other_class_personal',
                                                'test_same_sid_across_colleges_counted_once'],
    'tests/reviewer/test_allocation_and_rate.py': ['test_late_sheet_friday_exemption_is_surfaced'],
    'tests/reviewer/test_desktop_gui.py': ['test_selfcheck_command_is_wellformed'],
}


def strip_xfail(path: str, funcs) -> int:
    lines = io.open(path, encoding='utf-8').read().split('\n')
    out = []
    removed = 0
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.lstrip().startswith('@pytest.mark.xfail'):
            depth = line.count('(') - line.count(')')
            j = i
            while depth > 0 and j + 1 < len(lines):
                j += 1
                depth += lines[j].count('(') - lines[j].count(')')
            nxt = lines[j + 1] if j + 1 < len(lines) else ''
            keep = any(nxt.startswith('def ' + f + '(') for f in KEEP)
            if not keep:
                removed += 1
                i = j + 1          # 整块跳过
                continue
        out.append(line)
        i += 1
    if removed:
        io.open(path, 'w', encoding='utf-8', newline='').write('\n'.join(out))
    return removed


KEEP = {
    # 教室级按容量封顶是**对话口径明文要求**（220>219 锚点），不是缺陷，保留 xfail 登记
    'test_room_allocation_sum_equals_college_exempt_even_when_overflow',
}


def main() -> int:
    for filename, funcs in REMOVE.items():
        path = os.path.join(ROOT, filename)
        if not os.path.exists(path):
            print('跳过（文件不存在）：%s' % filename)
            continue
        n = strip_xfail(path, funcs)
        print('%s：摘掉 %d 处 xfail 标记' % (filename, n))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
