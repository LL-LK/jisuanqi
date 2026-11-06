# -*- coding: utf-8 -*-
"""排查免勤表里的占位学号（'全班' / '全体' / 空等）与集体申请的真实构成。"""

from __future__ import annotations

import glob
import os
import re
import sys
from collections import Counter, defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_exempt import load_exemptions      # noqa: E402

PLACEHOLDER = re.compile(r'^(全班|全体|全部|集体|无|/|待定)?$')


def main() -> int:
    path = None
    for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        if '免勤人员统计表' in os.path.basename(p):
            path = p
    records = load_exemptions(path)
    print('总记录 %d' % len(records))

    sid_kind = Counter()
    placeholder_rows = []
    for rec in records:
        sid = rec.sid
        if not sid:
            kind = '空学号'
        elif PLACEHOLDER.match(sid):
            kind = '占位学号:%s' % sid
        elif re.fullmatch(r'\d+', sid):
            kind = '数字学号'
        else:
            kind = '非数字学号'
        sid_kind[kind] += 1
        if kind != '数字学号':
            placeholder_rows.append(rec)

    print('\n学号形态分布：')
    for k, v in sid_kind.most_common():
        print('  %-20s %d' % (k, v))

    print('\n非数字学号记录（前 30 条，脱敏显示）：')
    for rec in placeholder_rows[:30]:
        print('  row=%-5s 学号占位=%-8s 学院=%-14s 班级=%-24s 星期=%-18s 周次=%-10s 人数=%s'
              % (rec.row, rec.sid or '(空)', rec.college[:14], rec.class_text[:24],
                 ','.join(sorted(rec.weekdays)), ','.join(str(x) for x in sorted(rec.weeks)) or '-',
                 rec.group_size))

    print('\n=== 商学院 / 周三 / 晚自习 细分 ===')
    hits = [r for r in records if '商学院' in r.college and r.session == '晚自习'
            and '周三' in r.weekdays]
    print('命中记录 %d 条' % len(hits))
    kinds = Counter()
    numeric_sids = set()
    placeholders = 0
    for rec in hits:
        if rec.sid and re.fullmatch(r'\d+', rec.sid):
            numeric_sids.add(rec.sid)
        elif rec.sid:
            placeholders += 1
        else:
            kinds['空学号'] += 1
    print('唯一数字学号 %d 个，非数字学号记录 %d 条' % (len(numeric_sids), placeholders))
    print('按班级分组的非数字学号：')
    byclass = defaultdict(list)
    for rec in hits:
        if not rec.sid or not re.fullmatch(r'\d+', rec.sid):
            byclass[(rec.class_text, rec.sid or '(空)')].append(rec.row)
    for k, rows in list(byclass.items())[:20]:
        print('   %-28s %-10s %d 行 %s' % (k[0][:28], k[1], len(rows), rows[:8]))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
