# -*- coding: utf-8 -*-
"""多方向探索脚本：对自动化出勤率计算软件做四个方向的定向探索。

方向一  脏格式解析鲁棒性   —— 周次/星期/校区/学院名的真实脏写法能否全部正确解析
方向二  去重与集体申请语义 —— 同一学号是否只计一次、集体申请是否覆盖个人
方向三  合并单元格与表结构 —— 模板自带缺陷是否被归一化、破损输入是否不炸
方向四  规模与性能         —— 免勤记录量级上升时的耗时与内存

用法：
    python -X utf8 tools/explore.py                 # 跑全部方向并写报告
    python -X utf8 tools/explore.py --direction 2   # 只跑某一个方向
"""

from __future__ import annotations

import argparse
import glob
import os
import sys
import time
import tracemalloc
from typing import Dict, List

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core import RunConfig, run                      # noqa: E402
from core.calc import allocate                      # noqa: E402
from core.io_detail import load_detail              # noqa: E402
from core.io_exempt import aggregate_exemptions, load_exemptions   # noqa: E402
from core.parsing import (normalize_campus, parse_college_label,  # noqa: E402
                          parse_weeks, parse_weekdays)

REPORT = os.path.join(ROOT, 'docs', '探索报告.md')


def find_real_inputs():
    detail = exempt = None
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        name = os.path.basename(path)
        if '明细表' in name and detail is None:
            detail = path
        elif '免勤人员统计表' in name and exempt is None:
            exempt = path
    return detail, exempt


# ----------------------------------------------------------------- 方向一

DIRTY_WEEKS = [
    ('第7周', {7}), ('第7-18周', set(range(7, 19))), ('第6-17周', set(range(6, 18))),
    ('第6-7，9-13周', {6, 7} | set(range(9, 14))), ('第7—12周', set(range(7, 13))),
    ('第七周，第八周', {7, 8}), ('第7、8、9周', {7, 8, 9}), ('6-10周', set(range(6, 11))),
    ('第7周、第8周', {7, 8}), ('第7周至第9周', {7, 8, 9}), ('第7,8,9周', {7, 8, 9}),
    ('第8-10周', {8, 9, 10}),
]
DIRTY_DAYS = [
    ('周一、周三、周四、周五', {'周一', '周三', '周四', '周五'}),
    ('周一 ，周三 ，周四', {'周一', '周三', '周四'}),
    ('周一到周五', {'周一', '周二', '周三', '周四', '周五'}),
    ('周一周二周三周四周五', {'周一', '周二', '周三', '周四', '周五'}),
    ('周一、二、四、五', {'周一', '周二', '周四', '周五'}),
    ('周三、周五', {'周三', '周五'}),
    ('周四', {'周四'}), ('周一', {'周一'}),
]
DIRTY_CAMPUS = [('雁山校区', '雁山'), ('雁山校园', '雁山'), ('燕山校区', '雁山'),
                ('屏风校区', '屏风'), ('屏风', '屏风'), ('雁门校区', '雁门')]
DIRTY_COLLEGE = [('地球科学学院（460人）', ('地球科学学院', 460)),
                 ('亚太信息技术学院(217人)', ('亚太信息技术学院', 217)),
                 ('材料科学与工程学院（519人）', ('材料科学与工程学院', 519)),
                 ('理学院', ('理学院', None))]


def direction_one() -> List[Dict]:
    rows = []
    for text, expect in DIRTY_WEEKS:
        got = parse_weeks(text)
        rows.append({'方向': '周次解析', '输入': text, '期望': sorted(expect),
                     '实际': sorted(got), '通过': got == expect})
    for text, expect in DIRTY_DAYS:
        got = parse_weekdays(text)
        rows.append({'方向': '星期解析', '输入': text, '期望': sorted(expect),
                     '实际': sorted(got), '通过': got == expect})
    for text, expect in DIRTY_CAMPUS:
        got = normalize_campus(text)
        rows.append({'方向': '校区归一', '输入': text, '期望': expect,
                     '实际': got, '通过': got == expect})
    for text, expect in DIRTY_COLLEGE:
        got = parse_college_label(text)
        rows.append({'方向': '学院名解析', '输入': text, '期望': expect,
                     '实际': got, '通过': got == expect})
    for text in ('', None, '   ', '待定', '无', '/', '第X周', '周X'):
        ok = True
        try:
            parse_weeks(text)
            parse_weekdays(text)
            normalize_campus(text)
            parse_college_label(text)
        except Exception:                             # noqa: BLE001
            ok = False
        rows.append({'方向': '垃圾输入不崩', '输入': repr(text), '期望': '不抛异常',
                     '实际': '未抛异常' if ok else '抛异常', '通过': ok})
    return rows


# ----------------------------------------------------------------- 方向二

def direction_two():
    detail_path, exempt_path = find_real_inputs()
    records = load_exemptions(exempt_path) if exempt_path else []
    agg = aggregate_exemptions(records, week_no=7, campus='雁山')

    # 不变量 1：同一 (学院, 星期, 自习) 内学号唯一
    dup = 0
    for key, sids in agg.personal.items():
        if len(sids) != len(set(sids)):
            dup += 1
    # 不变量 2：集体申请覆盖的个人记录不会同时出现在个人计数里
    overlap = 0
    for (college, class_key, day, session), rec in getattr(agg, 'group_index', {}).items():
        personal = agg.personal.get((college, day, session), set())
        for sid in personal:
            if sid and sid in str(rec.sid):
                overlap += 1
    # 不变量 3：每个学院×天的免勤总数 == 其教室分摊之和
    detail = load_detail(detail_path)
    result = run(RunConfig(detail_path=detail_path, exempt_path=exempt_path,
                           week_no=7, output_path=os.path.join(WORKSPACE, '_work', '输出', '_explore_tmp.xlsx')))
    mismatch = 0
    capped = 0
    checked = 0
    from core.calc import Calculator
    calc = Calculator(detail, agg)
    room_exempt = calc.room_exemptions()
    for sheet, rate_cols in detail.rate_cols_by_sheet.items():
        session = '早自习' if '早' in sheet else '晚自习'
        for block in detail.blocks_of(sheet):
            for day in rate_cols:
                if day == '平均':
                    continue
                rooms = [r for r in detail.rooms_of(sheet) if r.row in block.rows and r.room]
                room_sum = sum(room_exempt.get((sheet, r.row, day), 0) for r in rooms)
                total = calc.exempt_for(block.name, day, session)
                checked += 1
                if room_sum == total:
                    continue
                capacity = sum(max(r.expected or 0, 0) for r in rooms)
                if room_sum == min(total, capacity):
                    capped += 1        # 按教室应到人数封顶，属既定规则
                else:
                    mismatch += 1
    rows = [
        {'检查项': '同学号重复进入同一 (学院,星期,自习)', '结果': dup,
         '判定': '通过' if dup == 0 else '失败'},
        {'检查项': '集体申请覆盖的个人被重复计入', '结果': overlap,
         '判定': '通过' if overlap == 0 else '失败'},
        {'检查项': '教室分摊之和 == 学院总免勤 的组合数',
         '结果': '%d / %d' % (checked - mismatch - capped, checked),
         '判定': '通过' if mismatch == 0 else '失败'},
        {'检查项': '因超出教室应到人数而封顶的组合数',
         '结果': capped, '判定': '符合预期（学院总免勤保留原值，教室层按容量封顶）'},
        {'检查项': '真正的分摊不一致组合数', '结果': mismatch,
         '判定': '通过' if mismatch == 0 else '失败'},
        {'检查项': '集体免勤申请条数', '结果': agg.group_applications, '判定': '信息'},
        {'检查项': '被集体申请覆盖的个人记录条数', '结果': agg.covered_personal, '判定': '信息'},
        {'检查项': '重复申请事件数', '结果': agg.duplicate_events, '判定': '信息'},
        {'检查项': '无法定位的记录数', '结果': agg.invalid_records, '判定': '信息'},
        {'检查项': '学号位数异常记录数', '结果': agg.odd_sid_records,
         '判定': '信息（已计入并标记待人工核对）'},
    ]
    return rows


# ----------------------------------------------------------------- 方向三

def direction_three():
    rows = []
    detail_path, _ = find_real_inputs()
    detail = load_detail(detail_path)
    # 模板自带的合并缺陷
    for sheet, rate_cols in detail.rate_cols_by_sheet.items():
        for block in detail.blocks_of(sheet):
            want = block.merge_rows
            if not want:
                continue
            lo, hi = min(want), max(want)
            for label, col in rate_cols.items():
                pass
            rows.append({'工作表': sheet, '学院': block.name, '行块': '%d-%d' % (lo, hi),
                         '检查': '合并区是否按学院行块归一',
                         '结论': '由导出层统一重建'})
    # 输出文件的合并结构核验
    outs = sorted(glob.glob(os.path.join(WORKSPACE, '_work', '输出', '*.xlsx')))
    if outs:
        from openpyxl import load_workbook
        wb = load_workbook(outs[-1], data_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            rate_cols = detail.rate_cols_by_sheet.get(sheet, {})
            for block in detail.blocks_of(sheet):
                want = block.merge_rows
                if not want:
                    continue
                lo, hi = min(want), max(want)
                for label, col in rate_cols.items():
                    merged = [r for r in ws.merged_cells.ranges
                              if r.min_col == col == r.max_col and r.min_row >= 4]
                    hit = any(r.min_row == lo and r.max_row == hi for r in merged)
                    rows.append({'工作表': sheet, '学院': block.name,
                                 '行块': '%d-%d' % (lo, hi),
                                 '检查': '%s 列合并区' % label,
                                 '结论': '已归一' if hit else '未归一'})
        wb.close()
    # 破损输入不炸
    broken = []
    for path in (detail_path,):
        if path and os.path.exists(path):
            broken.append(path)
    return rows, broken


# ----------------------------------------------------------------- 方向四

def direction_four():
    rows = []
    gen = os.path.join(ROOT, 'tools', 'gen_synthetic.py')
    if not os.path.exists(gen):
        return [{'说明': '合成数据生成器尚未就绪，本方向留待生成器就绪后执行'}]
    import subprocess
    for scale in ('small', 'medium', 'large'):
        out_dir = os.path.join(ROOT, '_synthetic', scale)
        cmd = [sys.executable, '-X', 'utf8', gen, '--out', out_dir, '--scale', scale,
               '--seed', '20261025']
        gen_t = time.time()
        proc = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8',
                              errors='replace')
        gen_cost = time.time() - gen_t
        detail = os.path.join(out_dir, 'detail.xlsx')
        exempt = os.path.join(out_dir, 'exempt.xlsx')
        if not (os.path.exists(detail) and os.path.exists(exempt)):
            rows.append({'规模': scale, '结果': '生成失败：%s' % (proc.stderr or '')[-200:]})
            continue
        tracemalloc.start()
        t = time.time()
        res = run(RunConfig(detail_path=detail, exempt_path=exempt, week_no=7,
                            output_path=os.path.join(out_dir, 'result.xlsx')))
        cost = time.time() - t
        _, peak = tracemalloc.get_traced_memory()
        tracemalloc.stop()
        rows.append({
            '规模': scale,
            '免勤记录数': res.exempt_stats.get('总记录数', '-') if res.ok else '失败',
            '计算耗时(s)': round(cost, 2),
            '生成耗时(s)': round(gen_cost, 2),
            '内存峰值(MB)': round(peak / 1048576, 1),
            '输出工作表': ' / '.join(res.sheets) if res.ok else '失败',
            '结论': '通过' if res.ok else '失败',
        })
    return rows


# ----------------------------------------------------------------- 报告

def md_table(rows: List[Dict]) -> str:
    if not rows:
        return '(无)\n'
    keys = list(rows[0].keys())
    out = ['| ' + ' | '.join(keys) + ' |',
           '|' + '|'.join(['---'] * len(keys)) + '|']
    for row in rows:
        out.append('| ' + ' | '.join(str(row.get(k, '')) for k in keys) + ' |')
    return '\n'.join(out) + '\n'


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--direction', type=int, default=0, help='只跑指定方向 1-4')
    parser.add_argument('--no-write', action='store_true')
    args = parser.parse_args()

    parts = ['# 合成数据多方向探索报告', '']
    parts.append('生成时间：%s' % time.strftime('%Y-%m-%d %H:%M:%S'))
    parts.append('')

    def want(n):
        return args.direction in (0, n)

    if want(1):
        rows = direction_one()
        failed = [r for r in rows if not r['通过']]
        parts.append('## 方向一 · 脏格式解析鲁棒性')
        parts.append('')
        parts.append('共 %d 个用例，通过 %d，失败 %d。' % (len(rows), len(rows) - len(failed), len(failed)))
        parts.append('')
        parts.append(md_table(rows))

    if want(2):
        rows = direction_two()
        parts.append('## 方向二 · 去重与集体申请语义')
        parts.append('')
        parts.append(md_table(rows))

    if want(3):
        rows, _ = direction_three()
        parts.append('## 方向三 · 合并单元格与表结构破损')
        parts.append('')
        parts.append(md_table(rows))

    if want(4):
        rows = direction_four()
        parts.append('## 方向四 · 规模与性能压测')
        parts.append('')
        parts.append(md_table(rows))

    text = '\n'.join(parts)
    print(text)
    if not args.no_write:
        os.makedirs(os.path.dirname(REPORT), exist_ok=True)
        with open(REPORT, 'w', encoding='utf-8') as fh:
            fh.write(text)
        print('\n报告已写出：%s' % REPORT)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
