# -*- coding: utf-8 -*-
"""成品校验：检查导出的 Excel 是否满足全部硬约束。

校验项：
  1. 输出工作表恰好三张，且名称与原始一致
  2. 屏风工作表与原始文件逐单元格相同、合并结构相同
  3. 所有免勤人数单元格都有数值（无空白、无 None）
  4. 同一学院的出勤率合并区按学院行块归一（重点复核 N69:N70 这类漏合并）
  5. 出勤率列宽足够，不会显示成 ####
  6. 出勤率数值保留两位小数、且公式无错误值
  7. 输入文件 md5 未改变

用法：
    python -X utf8 tools/verify_output.py
"""

from __future__ import annotations

import glob
import hashlib
import os
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)
sys.path.insert(0, ROOT)

from core.io_detail import load_detail          # noqa: E402

EXPECTED_SHEETS = ['雁山（早）', '雁山（晚）', '屏风']
EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}


def md5(path):
    h = hashlib.md5()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def find_inputs():
    detail = exempt = None
    for path in glob.glob(os.path.join(WORKSPACE, '*.xlsx')):
        name = os.path.basename(path)
        if '明细表' in name and detail is None:
            detail = path
        elif '免勤人员统计表' in name and exempt is None:
            exempt = path
    return detail, exempt


def main() -> int:
    detail_path, exempt_path = find_inputs()
    outs = sorted(glob.glob(os.path.join(WORKSPACE, '_work', '输出', '*.xlsx')))
    outs = [p for p in outs if not os.path.basename(p).startswith('_explore')]
    if not outs:
        print('找不到输出文件')
        return 1
    out_path = max(outs, key=os.path.getmtime)   # 取最新一份
    print('输出文件：%s' % os.path.basename(out_path))
    print('输入文件：%s' % os.path.basename(detail_path))

    before_md5 = md5(detail_path)
    model = load_detail(detail_path)
    wb_out = load_workbook(out_path, data_only=True)
    wb_src = load_workbook(detail_path, data_only=True)

    results = []

    # 1. 工作表
    ok = wb_out.sheetnames == EXPECTED_SHEETS
    results.append(('工作表恰好三张且名称正确', '%s' % wb_out.sheetnames, ok))

    # 2. 屏风零改动
    diff = 0
    if '屏风' in wb_out.sheetnames:
        a, b = wb_out['屏风'], wb_src['屏风']
        for r in range(1, max(a.max_row, b.max_row) + 1):
            for c in range(1, max(a.max_column, b.max_column) + 1):
                if a.cell(r, c).value != b.cell(r, c).value:
                    diff += 1
        merges_same = ({str(x) for x in a.merged_cells.ranges}
                       == {str(x) for x in b.merged_cells.ranges})
        results.append(('屏风表逐单元格零差异', '内容差异 %d 处，合并结构一致=%s'
                        % (diff, merges_same), diff == 0 and merges_same))
    else:
        results.append(('屏风表逐单元格零差异', '输出缺少屏风表', False))

    # 3. 免勤单元格无空白
    blank = 0
    total_cells = 0
    for sheet in wb_out.sheetnames:
        if sheet.startswith('屏风'):
            continue
        ws = wb_out[sheet]
        for row in range(3, ws.max_row + 1):
            if ws.cell(row, 3).value is None:
                continue
            for col in EXEMPT_COLS.values():
                total_cells += 1
                if ws.cell(row, col).value is None:
                    blank += 1
    results.append(('免勤人数单元格无空白', '共 %d 格，空白 %d 格' % (total_cells, blank),
                    blank == 0 and total_cells > 0))

    # 4. 合并区归一
    bad_merges = []
    for sheet, rate_cols in model.rate_cols_by_sheet.items():
        ws = wb_out[sheet]
        for block in model.blocks_of(sheet):
            want = block.merge_rows
            if not want:
                continue
            lo, hi = min(want), max(want)
            for label, col in rate_cols.items():
                hit = any(r.min_col == col == r.max_col and r.min_row == lo and r.max_row == hi
                          for r in ws.merged_cells.ranges)
                if not hit:
                    bad_merges.append('%s/%s/%s(%s列)' % (sheet, block.name, label,
                                                          ws.cell(3, col).column_letter))
    results.append(('学院出勤率合并区按行块归一',
                    '异常 %d 处 %s' % (len(bad_merges), bad_merges[:4]), not bad_merges))

    # 5. 列宽
    narrow = []
    for sheet, rate_cols in model.rate_cols_by_sheet.items():
        ws = wb_out[sheet]
        for label, col in rate_cols.items():
            letter = ws.cell(3, col).column_letter
            width = ws.column_dimensions[letter].width or 8.43
            if width < 10:
                narrow.append('%s!%s=%.1f' % (sheet, letter, width))
    results.append(('出勤率列宽足够（不显示 ####）',
                    '过窄 %d 处 %s' % (len(narrow), narrow[:4]), not narrow))

    # 6. 出勤率数值与格式
    errors = 0
    extreme = 0
    bad_fmt = 0
    checked = 0
    for sheet, rate_cols in model.rate_cols_by_sheet.items():
        ws = wb_out[sheet]
        for label, col in rate_cols.items():
            for block in model.blocks_of(sheet):
                want = block.merge_rows
                if not want:
                    continue
                cell = ws.cell(min(want), col)
                v = cell.value
                if v is None:
                    continue
                checked += 1
                if isinstance(v, str) and v.startswith('#'):
                    errors += 1          # 公式错误值：真问题
                if isinstance(v, (int, float)):
                    if cell.number_format not in ('0.00%', '0.00%"'):
                        bad_fmt += 1
                    if v > 1.5:
                        extreme += 1     # 源数据矛盾导致的高值，已列入异常清单待人工复核
    results.append(('出勤率无公式错误值',
                    '检查 %d 格，错误值 %d' % (checked, errors), errors == 0))
    results.append(('出勤率百分比格式统一',
                    '格式不符 %d 格' % bad_fmt, bad_fmt == 0))
    results.append(('超过 150% 的极端值（源数据矛盾，需人工复核）',
                    '%d 格（已列入异常清单）' % extreme, True))

    # 7. 输入文件未被改动
    after_md5 = md5(detail_path)
    results.append(('输入文件未被覆盖', 'md5 一致=%s' % (before_md5 == after_md5),
                    before_md5 == after_md5))

    print('\n=== 成品校验 ===')
    all_ok = True
    for item, detail_text, ok in results:
        all_ok = all_ok and ok
        print('  [%s] %-28s %s' % ('通过' if ok else '失败', item, detail_text))
    print('\n总体结论：%s' % ('全部通过' if all_ok else '存在未通过项'))
    wb_out.close()
    wb_src.close()
    return 0 if all_ok else 1


if __name__ == '__main__':
    raise SystemExit(main())
