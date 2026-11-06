# -*- coding: utf-8 -*-
"""针对某 (学院, 星期, 自习) 组合，列出免勤表里的全部候选记录，并对比多种口径的计数。

用法：
    python tools/inspect_exempt.py 商学院 周三 晚自习
    python tools/inspect_exempt.py 数学与统计学院 周一 晚自习
"""

from __future__ import annotations

import glob
import os
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_exempt import load_exemptions      # noqa: E402
from core.parsing import normalize_class        # noqa: E402


def find_exempt():
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '免勤人员统计表' in os.path.basename(path):
            return path
    return None


def main() -> int:
    if len(sys.argv) < 4:
        print(__doc__)
        return 1
    college_want, day_want, session_want = sys.argv[1], sys.argv[2], sys.argv[3]
    week = int(sys.argv[4]) if len(sys.argv) > 4 else 7

    records = load_exemptions(find_exempt())
    hits = []
    for rec in records:
        if college_want not in rec.college:
            continue
        if session_want and rec.session != session_want:
            continue
        if day_want not in rec.weekdays:
            continue
        hits.append(rec)

    print('命中候选记录 %d 条' % len(hits))
    in_week = [r for r in hits if week in r.weeks]
    print('其中免勤周次含第 %d 周：%d 条' % (week, len(in_week)))

    sids = {r.sid for r in in_week if r.sid}
    groups = defaultdict(int)
    for r in in_week:
        if r.is_group:
            groups[(r.class_key, r.group_size)] += 1
    print('唯一学号数：%d' % len(sids))
    print('集体申请：%s' % (dict(groups) or '无'))

    print('\n--- 记录明细 ---')
    for r in hits[:40]:
        print('  row=%-5s sid=%-14s 班级=%-22s 周次=%-14s 星期=%-22s 自习=%-6s 校区=%-6s 人数=%s'
              % (r.row, r.sid or '(集体)', r.class_text[:22],
                 ','.join(str(x) for x in sorted(r.weeks)) or '-',
                 ','.join(sorted(r.weekdays)), r.session, r.campus, r.group_size))

    print('\n--- 班级 key 分布（用于判断集体申请覆盖是否生效）---')
    by_key = defaultdict(list)
    for r in in_week:
        by_key[r.class_key].append(r)
    for key, rs in sorted(by_key.items())[:20]:
        kinds = ['集体%s' % r.group_size if r.is_group else ('学号' if r.sid else '无学号')
                 for r in rs]
        print('  %-26s 共 %2d 条  %s' % (key or '(空)', len(rs), kinds[:12]))

    print('\n--- 全表：该学院所有记录的周次/星期分布 ---')
    allc = [r for r in records if college_want in r.college]
    dist = defaultdict(int)
    for r in allc:
        for d in r.weekdays:
            dist[(r.session, d)] += 1
    for key in sorted(dist, key=lambda x: (str(x[0]), str(x[1]))):
        print('  %-8s %-4s %d 条' % (key[0], key[1], dist[key]))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
