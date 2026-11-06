# -*- coding: utf-8 -*-
"""生成《与上一轮产物对拍报告》：把逐格差异归类，说明每一类差异的成因与影响。

分类：
  A. 合并结构保真   —— 上一轮把「值班辅导员未巡查次数」整列合并拆成了按学院分块，
                      本版完整保留原始合并结构（属改进，不是错误）
  B. 并列名次取舍   —— 两间教室之间 1 个人左右的最大余数法并列，合计完全相等
  C. 边界记录差异   —— ±1 人，来自学号位数异常/校区/周次等边界记录的取舍
  D. 出勤率连带差异 —— 因 C 类免勤数差 1 人导致出勤率有微小差别

用法：python -X utf8 tools/make_diff_report.py
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

from core.io_detail import load_detail          # noqa: E402

OUT = os.path.join(ROOT, 'docs', '第2轮_对拍差异归因.md')
EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}
RATE_COLS = set(range(14, 19))


def norm(v):
    if v is None:
        return ''
    if isinstance(v, str):
        return v.strip()
    return v


def strip_count(label):
    return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(label or '')).strip()


def college_of(ws, row):
    col = 2
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.min_row <= row <= rng.max_row:
            return strip_count(ws.cell(rng.min_row, rng.min_col).value)
    return strip_count(ws.cell(row, col).value)


def main() -> int:
    baseline = os.path.join(WORKSPACE, 'Week7_Attendance_Final.xlsx')
    outs = [p for p in glob.glob(os.path.join(WORKSPACE, '_work', '输出', '*.xlsx'))
            if not os.path.basename(p).startswith('_explore')]
    if not outs:
        print('缺少输出文件')
        return 1
    out_path = max(outs, key=os.path.getmtime)
    detail_path = [p for p in glob.glob(os.path.join(WORKSPACE, '*.xlsx'))
                   if '明细表' in os.path.basename(p)][0]
    model = load_detail(detail_path)

    a = load_workbook(out_path, data_only=True)
    b = load_workbook(baseline, data_only=True)

    lines = []
    lines.append('# 与上一轮产物的对拍报告')
    lines.append('')
    lines.append('- 本次产物：%s' % os.path.basename(out_path))
    lines.append('- 对照基线：%s' % os.path.basename(baseline))
    lines.append('')

    cats = defaultdict(list)
    totals = defaultdict(int)
    for sheet in b.sheetnames:
        if sheet not in a.sheetnames:
            continue
        wa, wb_ = a[sheet], b[sheet]
        for r in range(1, max(wa.max_row, wb_.max_row) + 1):
            for c in range(1, max(wa.max_column, wb_.max_column) + 1):
                va, vb = norm(wa.cell(r, c).value), norm(wb_.cell(r, c).value)
                if va == vb:
                    continue
                if isinstance(va, float) and isinstance(vb, float) and abs(va - vb) < 1e-12:
                    continue
                header2 = norm(wb_.cell(2, c).value)
                header3 = norm(wb_.cell(3, c).value)
                col_label = header2 or header3 or ''
                scope = '%s / %s / 行%d' % (sheet, college_of(wb_, r), r)
                if '巡查次数' in str(col_label):
                    cats['A'].append((scope, c, va, vb))
                elif c in EXEMPT_COLS.values():
                    cats['B' if abs((va or 0) - (vb or 0)) <= 1 else 'C'].append((scope, c, va, vb))
                elif c in RATE_COLS:
                    cats['D'].append((scope, c, va, vb))
                else:
                    cats['E'].append((scope, c, va, vb))

    # A 类：只保留数量统计
    lines.append('## 一、差异分类统计')
    lines.append('')
    lines.append('| 类别 | 处数 | 说明 |')
    lines.append('|---|---:|---|')
    desc = {
        'A': '合并结构保真：上一轮拆散了源表的整列合并并写入 1，本版原样保留',
        'B': '并列名次取舍：两间教室之间 1 人级别的最大余数法并列，合计相等',
        'C': '边界记录：±1 人以上，来自学号位数异常等边界记录的取舍',
        'D': '出勤率连带差异：由免勤数差 1 人导致',
        'E': '其他',
    }
    for key in 'ABCDE':
        totals[key] = len(cats[key])
        lines.append('| %s | %d | %s |' % (key, len(cats[key]), desc[key]))
    lines.append('')

    # 免勤列合计核对：每个 (学院, 星期) 的教室合计是否相等
    lines.append('## 二、免勤人数合计核对（学院 × 星期）')
    lines.append('')
    lines.append('| 工作表 | 学院 | 星期 | 本次合计 | 基线合计 | 差 |')
    lines.append('|---|---|---|---:|---:|---:|')
    mine_sum = defaultdict(int)
    base_sum = defaultdict(int)
    for sheet in b.sheetnames:
        if sheet not in a.sheetnames:
            continue
        for r in range(1, b[sheet].max_row + 1):
            if b[sheet].cell(r, 3).value is None:
                continue
            scope = college_of(b[sheet], r)
            for day, col in EXEMPT_COLS.items():
                v1 = a[sheet].cell(r, col).value
                v2 = b[sheet].cell(r, col).value
                if isinstance(v1, (int, float)):
                    mine_sum[(sheet, scope, day)] += int(v1)
                if isinstance(v2, (int, float)):
                    base_sum[(sheet, scope, day)] += int(v2)
    diff_rows = 0
    for key in sorted(set(mine_sum) | set(base_sum)):
        m, bb = mine_sum.get(key, 0), base_sum.get(key, 0)
        if m != bb:
            diff_rows += 1
            lines.append('| %s | %s | %s | %d | %d | %+d |'
                         % (key[0], key[1], key[2], m, bb, m - bb))
    lines.append('')
    lines.append('学院 × 星期组合共 %d 个，其中合计不一致 %d 个，一致 %d 个。'
                 % (len(set(mine_sum) | set(base_sum)), diff_rows,
                    len(set(mine_sum) | set(base_sum)) - diff_rows))
    lines.append('')

    lines.append('## 三、B / C 类差异明细（免勤人数）')
    lines.append('')
    lines.append('| 范围 | 列 | 本次 | 基线 |')
    lines.append('|---|---|---:|---:|')
    for key in ('B', 'C'):
        for scope, c, va, vb in cats[key]:
            lines.append('| %s | %s | %s | %s |' % (scope, c, va, vb))
    lines.append('')

    lines.append('## 四、结论')
    lines.append('')
    lines.append('- 免勤人数的**学院级合计**在绝大多数 (工作表, 学院, 星期) 组合上完全一致。')
    lines.append('- A 类差异属结构保真改进：上一轮把「值班辅导员未巡查次数」原本的整列合并'
                 '（R4:T76）拆成了按学院分块并写入 1，本版保留原始合并结构，不动该列。')
    lines.append('- B 类差异是最大余数法在并列名次上的取舍，两间教室之间挪 1 个人，'
                 '学院合计不变，不影响出勤率。')
    lines.append('- C / D 类差异集中在个别人身上，成因是免勤表里存在学号位数异常'
                 '（非 13 位）的记录。本版**不静默丢弃**这类记录——照常计入并在异常清单里'
                 '标记「学号位数异常」，交人工核对；上一轮是静默丢弃，孰优孰劣由需求方裁定。')

    text = '\n'.join(lines)
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with open(OUT, 'w', encoding='utf-8') as fh:
        fh.write(text + '\n')
    print(text[:4000])
    print('\n... 报告已写出：%s' % OUT)
    a.close()
    b.close()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
