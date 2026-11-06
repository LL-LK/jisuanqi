# -*- coding: utf-8 -*-
"""定位某个 (学院, 星期, 自习) 范围内所有「边界记录」。

边界记录 = 占位学号 / 周次无法解析 / 星期无法解析 / 校区非雁山 / 集体申请
这些正是造成 ±1 人差异的来源。

用法：python -X utf8 tools/edge_cases.py 外国语学院 周一 晚自习
"""

from __future__ import annotations

import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_exempt import load_exemptions      # noqa: E402


def main() -> int:
    college, day, session = sys.argv[1], sys.argv[2], sys.argv[3]
    week = int(sys.argv[4]) if len(sys.argv) > 4 else 7
    path = [p for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx'))
            if '免勤人员统计表' in os.path.basename(p)][0]
    records = load_exemptions(path)

    hits = [r for r in records if r.college == college and r.session == session
            and day in r.weekdays]
    print('范围 %s / %s / %s：候选记录 %d 条' % (college, day, session, len(hits)))

    in_week = [r for r in hits if week in r.weeks]
    sids = {r.sid for r in in_week if r.sid}
    print('第%d周有效 %d 条；唯一学号 %d 个' % (week, len(in_week), len(sids)))

    edges = []
    for r in hits:
        if not r.weeks:
            edges.append(('周次无法解析', r))
        elif week not in r.weeks:
            edges.append(('周次不含本周', r))
        if not r.weekdays:
            edges.append(('星期无法解析', r))
        if r.campus != '雁山':
            edges.append(('校区非雁山(%s)' % r.campus, r))
        if r.sid and not re.fullmatch(r'\d{13}', r.sid):
            edges.append(('非数字学号(%s)' % r.sid, r))
        if r.is_group:
            edges.append(('集体申请(%d人)' % r.group_size, r))
        if not r.sid and not r.is_group:
            edges.append(('无学号且非集体', r))

    print('\n边界记录 %d 条：' % len(edges))
    for kind, r in edges:
        print('  [%-18s] row=%-5s 班级=%-22s 周次=%-14s 星期=%-20s 校区=%-6s 人数=%s'
              % (kind, r.row, r.class_text[:22],
                 ','.join(str(x) for x in sorted(r.weeks)) or '(空)',
                 ','.join(sorted(r.weekdays)) or '(空)', r.campus, r.group_size))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
