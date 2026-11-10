# -*- coding: utf-8 -*-
"""明细表读取：把「执行情况明细表」解析成教室行 / 学院块 / 出勤率区。

明细表结构（来自真实样本）：
    第2行表头  序号 | 学院（人数） | 教室 | 分配专业班级 | 应到人数 |
              周一 | 周一免勤人数 | 周三 | 周三免勤人数 | 周四 | 周四免勤人数 |
              周五 | 周五免勤人数 | 学生出勤率（%） | ... | 早自习考勤时间 | 执行效果
    第3行子表头  N=周一 O=周三 P=周四 Q=周五 R=平均（晚自习表无周五列）

第3行同时也是第一个学院（地球科学学院）的首行数据——这是模板自带的怪癖，
所以学院的合并区一律**排除第3行**。
"""

from __future__ import annotations

import re
from collections import Counter
from contextlib import closing
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .model import CollegeBlock, DetailModel, RoomRow
from .parsing import (WEEKDAY_NAMES, normalize_text, parse_college_label,
                      parse_int, parse_weeks, strip_spaces)

RATE_BLOCK_LABELS = set(WEEKDAY_NAMES) | {'平均'}
WEEK_LABEL = re.compile(r'第\s*[0-9零〇一二三四五六七八九十两]+\s*周')


def week_from_label(text):
    """从文件名或表标题的明确周次标签提取唯一周次。"""
    weeks = {week for label in WEEK_LABEL.findall(normalize_text(text))
             for week in parse_weeks(label) if 1 <= week <= 30}
    return next(iter(weeks)) if len(weeks) == 1 else None


def infer_workbook_week(path):
    """优先使用表内标题。多个标题冲突且无多数时，要求明确指定周次。"""
    with open(path, 'rb') as source, closing(load_workbook(source, read_only=True, data_only=True)) as book:
        sheets = [sheet for sheet in book if sheet.title in ('雁山（早）', '雁山（晚）', '屏风')]
        if not sheets:
            sheets = [sheet for sheet in book if sheet.sheet_state == "visible"]
        weeks = Counter(week for sheet in sheets if (week := week_from_label(sheet['A1'].value)))
    if weeks:
        week, count = weeks.most_common(1)[0]
        if count * 2 <= sum(weeks.values()):
            raise ValueError('考勤表内标题周次冲突，请核对并手动指定统计周次')
        return week
    return week_from_label(Path(path).name)


def actual_cells(workbook):
    """以原表的表头和坐标定位雁山实到栏，保留空白、零、文字及公式。"""
    return {
        (sheet.title, row, col): sheet.cell(row, col).value
        for sheet in workbook if sheet.title in ('雁山（早）', '雁山（晚）')
        for header in (_find_header_row(sheet),)
        for col in _detect_data_cols(sheet, header)['actuals'].values()
        for row in range(header + 1, sheet.max_row + 1)
    }


def _merged_value(ws, col: int, row: int):
    """取合并区左上角的值；非合并单元格返回自身。"""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return ws.cell(rng.min_row, rng.min_col).value
    return ws.cell(row, col).value


def _merged_ranges_in_col(ws, col: int):
    return [rng for rng in ws.merged_cells.ranges
            if rng.min_col == col == rng.max_col]


def _find_header_row(ws, max_probe: int = 8) -> int:
    """表头行 = 第一行同时含 '应到人数' 与 '教室' 的行。"""
    for row in range(1, min(max_probe, ws.max_row) + 1):
        labels = [normalize_text(ws.cell(row, c).value) for c in range(1, ws.max_column + 1)]
        if any('应到人数' in x for x in labels) and any('教室' in x for x in labels):
            return row
    return 2


# R14：这些词出现在主表头里，说明该列是**其它指标**（不是出勤率），
# 必须排除在「出勤率列重建范围」之外。
NON_RATE_HEADER_KEYS = ('巡查', '执行效果', '考核', '值班', '考勤时间', '考察时间')


def _effective_header(ws, row: int, col: int) -> str:
    """取某列在主表头行的**有效标题**（命中合并区时取左上角）。"""
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            return normalize_text(ws.cell(rng.min_row, rng.min_col).value)
    return normalize_text(ws.cell(row, col).value)


def _detect_rate_cols(ws, header_row: int, actual_days):
    """找出出勤率区各列：'周一' '周三' '周四' '周五' '平均' → 列号。

    返回 (cols, excluded)：
      * cols     参与合并区重建与出勤率写入的列
      * excluded 被明确排除的**非出勤率指标列**（R14 逐列实证的留痕）

    R14：范围**不得硬编码** N..R / N..Q，必须按列标题逐列反查：
      * 出勤率区的列范围 = 主表头合并区（如 N2:R2 / N2:Q2）
      * 每列再核对主表头语义，确认不是「值班辅导员未巡查次数」这类指标列
    """
    start_col = None
    for col in range(1, ws.max_column + 1):
        head = normalize_text(ws.cell(header_row, col).value)
        if '出勤率' in head and not any(key in head for key in NON_RATE_HEADER_KEYS):
            start_col = col
            break
    if start_col is None:
        return {}, [], None

    # 出勤率区的列范围 = 主表头合并区（早表 N2:R2 / 晚表 N2:Q2）
    end_col = ws.max_column
    merged_header = False
    for rng in ws.merged_cells.ranges:
        if (rng.min_row <= header_row <= rng.max_row
                and rng.min_col <= start_col <= rng.max_col):
            end_col = rng.max_col
            merged_header = True
            break

    # 最新版模板的子表头是空的，但仍留有独立的一行及完整的出勤率合并标题。
    # 只有列数与实际星期数加一完全一致、空格下一行是率区纵向合并的起点时才补列名。
    label_row = header_row + 1
    rate_columns = range(start_col, end_col + 1)
    labels = [normalize_text(ws.cell(label_row, col).value) for col in rate_columns]
    label_cells = {ws.cell(label_row, col).coordinate for col in rate_columns}
    dedicated_row = all(not any(cell in area for area in ws.merged_cells.ranges)
                        for cell in label_cells)
    merge_starts = {r.min_col for r in ws.merged_cells.ranges
                    if r.min_col == r.max_col and r.min_row == label_row + 1}
    inferred_row = None
    expected_labels = [*actual_days, '周平均']
    if (merged_header and actual_days and dedicated_row and len(labels) == len(expected_labels)
            and set(rate_columns) <= merge_starts
            and (not any(labels) or labels == expected_labels)):
        labels = [*actual_days, '平均']
        inferred_row = label_row
    elif not any(label in RATE_BLOCK_LABELS for label in labels):
        for row in range(label_row + 1, min(header_row + 4, ws.max_row) + 1):
            candidate = [normalize_text(ws.cell(row, col).value) for col in rate_columns]
            if candidate[0] in RATE_BLOCK_LABELS:
                labels = candidate
                break

    cols: Dict[str, int] = {}
    excluded: List[Tuple[int, str]] = []
    for col, label in zip(rate_columns, labels):
        if label not in RATE_BLOCK_LABELS:
            continue
        head = _effective_header(ws, header_row, col)
        if any(k in head for k in NON_RATE_HEADER_KEYS):
            excluded.append((col, head))
            continue
        if head and '出勤率' not in head:
            continue
        cols[label] = col
    # 出勤率区之外、但语义上属于其它指标列的（如晚表 R 列「值班辅导员未巡查次数」）
    for col in range(end_col + 1, ws.max_column + 1):
        head = _effective_header(ws, header_row, col)
        if head and any(k in head for k in NON_RATE_HEADER_KEYS):
            excluded.append((col, head))
    return cols, excluded, inferred_row


def _parse_headcount(value):
    """保留数值型文本，拒绝把备注里的数字、负数或小数当作人数。"""
    text = normalize_text(value)
    match = re.fullmatch(r'([0-9]+)(?:\.0+)?\s*人?', text)
    return int(match[1]) if match else None


def _detect_data_cols(ws, header_row: int):
    """识别序号 / 学院 / 教室 / 班级 / 应到 / 各星期实到列与免勤列。"""
    seq_col = college_col = room_col = classes_col = expected_col = None
    actual_cols: Dict[str, int] = {}
    exempt_cols: Dict[str, int] = {}

    for col in range(1, ws.max_column + 1):
        head = strip_spaces(ws.cell(header_row, col).value)
        if not head:
            continue
        if head == '序号':
            seq_col = col
        elif head.startswith('学院') and '人数' in head:
            college_col = col
        elif head == '教室':
            room_col = col
        elif '专业班级' in head:
            classes_col = col
        elif head == '应到人数':
            expected_col = col
        elif '免勤人数' in head:
            day = head.replace('免勤人数', '')
            if day in RATE_BLOCK_LABELS:
                exempt_cols[day] = col
        elif head in WEEKDAY_NAMES:
            actual_cols[head] = col

    return dict(seq=seq_col, college=college_col, room=room_col,
                classes=classes_col, expected=expected_col,
                actuals=actual_cols, exempts=exempt_cols)


def load_detail(path: str) -> DetailModel:
    """读取明细表。只解析雁山校区的工作表；屏风表原样保留、不参与任何计算。"""
    model = DetailModel(path=path)
    with open(path, 'rb') as f:
        wb = load_workbook(f, data_only=True)
        try:
            model.sheet_names = list(wb.sheetnames)

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                title_week = week_from_label(ws['A1'].value)
                if title_week is not None:
                    model.title_weeks[sheet_name] = title_week
                if not sheet_name.startswith('雁山'):
                    # 屏风校区：完全排除，只保留工作表本身
                    continue
                header_row = _find_header_row(ws)
                cols = _detect_data_cols(ws, header_row)
                rate_cols, excluded_cols, inferred_row = _detect_rate_cols(ws, header_row, tuple(cols['actuals']))
                model.rate_cols_by_sheet[sheet_name] = rate_cols
                model.exempt_cols_by_sheet[sheet_name] = cols['exempts']
                model.protected_cols[sheet_name] = {col for col in range(1, ws.max_column + 1)
                                                  if any(key in _effective_header(ws, header_row, col)
                                                         for key in NON_RATE_HEADER_KEYS)}
                if inferred_row:
                    model.rate_label_rows[sheet_name] = inferred_row
                exempt_columns = set(cols['exempts'].values())
                model.college_exempt_layout |= any(
                    r.min_col == r.max_col and r.min_col in exempt_columns
                    and r.min_row > header_row and r.max_row > r.min_row
                    for r in ws.merged_cells.ranges)
                for col, head in excluded_cols:
                    model.warnings.append(
                        '工作表「%s」第 %s 列（%s）是非出勤率指标列，已排除在合并区重建范围外'
                        % (sheet_name, get_column_letter(col), head))

                if cols['college'] is None or cols['room'] is None:
                    model.warnings.append(f'工作表「{sheet_name}」缺少学院列或教室列，已跳过')
                    continue

                # ---- 学院块：优先用学院列的合并区；未合并的单元格各算一块 ----
                merged_rows = set()
                blocks: List[CollegeBlock] = []
                for rng in sorted(_merged_ranges_in_col(ws, cols['college']),
                                  key=lambda r: r.min_row):
                    if rng.max_row <= header_row:
                        continue
                    rows = list(range(rng.min_row, rng.max_row + 1))
                    merged_rows.update(rows)
                    raw = ws.cell(rng.min_row, rng.min_col).value
                    name, total = parse_college_label(raw)
                    blocks.append(CollegeBlock(sheet=sheet_name, name=name, total=total,
                                               rows=rows, rate_cols=dict(rate_cols)))
                for row in range(header_row + 1, ws.max_row + 1):
                    if row in merged_rows:
                        continue
                    raw = ws.cell(row, cols['college']).value
                    if not normalize_text(raw):
                        continue
                    name, total = parse_college_label(raw)
                    blocks.append(CollegeBlock(sheet=sheet_name, name=name, total=total,
                                               rows=[row], rate_cols=dict(rate_cols)))
                blocks.sort(key=lambda b: b.rows[0])
                # 兼容未合并 B 列的连续多教室行：若某行有教室号但 B 列为空且不在合并区，
                # 自动归入紧邻的上一学院块，防止丢失学院归属。
                for row in range(header_row + 1, ws.max_row + 1):
                    if row in merged_rows:
                        continue
                    raw_col = ws.cell(row, cols['college']).value
                    if normalize_text(raw_col):
                        continue
                    room_val = strip_spaces(ws.cell(row, cols['room']).value)
                    if not room_val:
                        continue
                    for blk in reversed(blocks):
                        if blk.rows and blk.rows[-1] == row - 1:
                            blk.rows.append(row)
                            break
                model.blocks.extend(blocks)

                block_of_row = {}
                for blk in blocks:
                    for r in blk.rows:
                        block_of_row[r] = blk

                # ---- 教室行 ----
                for row in range(header_row + 1, ws.max_row + 1):
                    room_raw = ws.cell(row, cols['room']).value
                    room = strip_spaces(room_raw)
                    blk = block_of_row.get(row)
                    college_label = normalize_text(
                        _merged_value(ws, cols['college'], row)) if cols['college'] else ''
                    name, total = parse_college_label(college_label)
                    if not name and blk:
                        name = blk.name
                        total = blk.total
                    if not room and not blk:
                        continue
                    if not room:
                        continue
                    actuals, actual_raw = {}, {}
                    for day, col in cols['actuals'].items():
                        raw = ws.cell(row, col).value
                        actual_raw[day] = '' if raw is None else str(raw).strip()
                        actuals[day] = _parse_headcount(raw)
                    classes_text = normalize_text(
                        ws.cell(row, cols['classes']).value) if cols['classes'] else ''
                    model.rooms.append(RoomRow(
                        sheet=sheet_name,
                        row=row,
                        seq=parse_int(ws.cell(row, cols['seq']).value) if cols['seq'] else None,
                        college_name=name,
                        college_total=total if total is not None else (blk.total if blk else None),
                        room=room,
                        classes_text=classes_text,
                        expected=_parse_headcount(ws.cell(row, cols['expected']).value) if cols['expected'] else None,
                        actuals=actuals,
                        actual_raw=actual_raw,
                        actual_cols=dict(cols['actuals']),
                        exempt_cols=dict(cols['exempts']),
                    ))

                # ---- 块内教室数据行（R7 ①）----
                # 教室行 = 含有效教室号且有应到数值的行；纯表头行剔除。
                # 合并区起点 = 该学院首个教室数据行（第 3 行含 6504，必须在区内）。
                for blk in blocks:
                    rows = [r.row for r in model.rooms
                            if r.sheet == sheet_name and r.row in blk.rows
                            and r.room and r.expected is not None]
                    blk.data_rows = sorted(rows)
        finally:
            wb.close()
    return model
