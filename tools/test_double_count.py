# -*- coding: utf-8 -*-
"""验证假设：基线是否用「班级 → 教室」匹配，把同一个人重复计入多间教室。

对每个 (学院, 星期, 自习)：
    H1 = 按班级匹配到教室后，各教室人数之和（允许重复计入）
    H2 = 唯一学号去重人数（对话口径）
    基线 = 基线文件里该学院各教室免勤之和
若 H1 == 基线，则基线确实用了班级匹配口径。
"""

from __future__ import annotations

import glob
import os
import re
import sys
from collections import defaultdict

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_exempt import load_exemptions              # noqa: E402
from core.parsing import normalize_class               # noqa: E402

EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}


def strip_count(label):
    return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(label or '')).strip()


def room_classes(text):
    """从 '数字26-2（39）、数字26-3（38）' 拆出班级 token。"""
    out = []
    for part in re.split(r'[、,，;；]', str(text or '')):
        part = re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', part).strip()
        if part:
            out.append(normalize_class(part))
    return [x for x in out if x]


def main() -> int:
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    src = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '明细表' in os.path.basename(p):
            src = p
    exempt_path = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '免勤人员统计表' in os.path.basename(p):
            exempt_path = p

    wb_src = load_workbook(src, data_only=True)
    wb_base = load_workbook(baseline, data_only=True)

    # 明细表：每个学院的 (教室 -> 班级集合)
    rooms_by_college = defaultdict(list)
    for sheet in ('雁山（早）', '雁山（晚）'):
        ws = wb_src[sheet]
        current = None
        for row in range(3, ws.max_row + 1):
            label = ws.cell(row, 2).value
            if label:
                current = strip_count(label)
            room = ws.cell(row, 3).value
            if not room:
                continue
            classes = room_classes(ws.cell(row, 4).value)
            expected = ws.cell(row, 5).value
            rooms_by_college[(sheet, current)].append((row, str(room), classes, expected or 0))

    # 基线：各学院各星期教室免勤合计
    base_totals = defaultdict(int)
    for sheet in ('雁山（早）', '雁山（晚）'):
        ws = wb_base[sheet]
        current = None
        for row in range(3, ws.max_row + 1):
            label = ws.cell(row, 2).value
            if label:
                current = strip_count(label)
            if ws.cell(row, 3).value is None:
                continue
            for day, col in EXEMPT_COLS.items():
                v = ws.cell(row, col).value
                if isinstance(v, (int, float)):
                    base_totals[(sheet, current, day)] += int(v)

    records = load_exemptions(exempt_path)

    print('%-8s %-16s %-4s %-6s %-6s %-6s' % ('工作表', '学院', '星期', '基线', 'H1班级匹配', 'H2唯一学号'))
    hit1 = hit2 = total = 0
    for (sheet, college), rooms in sorted(rooms_by_college.items()):
        session = '早自习' if '早' in sheet else '晚自习'
        for day in EXEMPT_COLS:
            base = base_totals.get((sheet, college, day), 0)
            if base == 0:
                continue
            # 命中该学院该天该次自习、且第7周有效的记录
            hits = [r for r in records
                    if r.college == college and r.session == session and day in r.weekdays
                    and (not r.weeks or 7 in r.weeks)]
            h2 = len({r.sid for r in hits if r.sid})
            h1 = 0
            for r in hits:
                if not r.sid:
                    continue
                key = r.class_key
                for _row, _room, classes, _exp in rooms:
                    if key and any(key == c or key in c or c in key for c in classes):
                        h1 += 1
                        break
            total += 1
            mark1 = '=' if h1 == base else ''
            mark2 = '=' if h2 == base else ''
            hit1 += 1 if h1 == base else 0
            hit2 += 1 if h2 == base else 0
            if h1 != base or h2 != base:
                print('%-8s %-16s %-4s %-6d %-11d%s %-10d%s'
                      % (sheet[2:-1], college, day, base, h1, mark1, h2, mark2))
    print('\n组合总数 %d：H1(班级匹配) 命中 %d，H2(唯一学号) 命中 %d' % (total, hit1, hit2))
    wb_src.close(); wb_base.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
