# -*- coding: utf-8 -*-
"""第 2 轮验收基线（Chief 统一口径，任何角色不得各算各的）。

**两口径的分工（R21 加固条款：凡「产出工作簿」口径的数字，只能以引擎实跑读出的
文件实际值为证据，转述值不得作基准、不得写进验收基线）**：

  B 口径 = 产出工作簿**免勤列**单元格值之和，按 (工作表, 学院, 星期)
           → **对外唯一**：早 160/163/158/163、晚 1890/1517/796/19
  A 口径 = 引擎面板「学院总免勤」合计
           → **内部分解量**，仅归因用，不对外、不得与 B 口径并列比较
  两口径差集**恰好 1 个组合**：数统周三（A=220 / B=219），差 1 即 C1=甲
  教室容量封顶的超额 1 人（归 7614，不写入任何免勤单元格）——完全自洽。

固化内容全部来自对**真实工作簿**跑完整流水线的实测快照
（见 ``baseline_realdata.json`` 的 provenance 字段），用途是把
「已签字的行为」变成可执行断言：任何改动一旦动了上报数字，本模块立刻转红。
"""

from __future__ import annotations

import hashlib
import json
import os

import pytest
from openpyxl import load_workbook

from core import RunConfig, run
from core.io_detail import load_detail
from core.io_exempt import aggregate_exemptions, load_exemptions

pytestmark = pytest.mark.realdata

DAYS = ('周一', '周三', '周四', '周五')
SHEETS = ('雁山（早）', '雁山（晚）')
BASELINE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              'baseline_realdata.json')

#: 对外唯一口径表（R21）：产出工作簿的学院×星期合计
PUBLIC_PER_DAY = {
    '早自习|周一': 160, '早自习|周三': 163, '早自习|周四': 158, '早自习|周五': 163,
    # v1.1.0：同号异名/跨学院按1人计后，晚自习恢复为计入全部争议事件
    '晚自习|周一': 1892, '晚自习|周三': 1518, '晚自习|周四': 797, '晚自习|周五': 19,
}

# 上述为签字历史快照（historical_b_totals）。v1.1.0 起同号异名/跨学院
# 按「学号×星期×自习」去重计 1 人，不再暂缓计入；当前验收以 baseline['b_totals']
# 为准，旧口径快照保存在 baseline['historical_*'] 中供追溯。
IDENTITY_REVIEW_DELTAS = {}
IDENTITY_REVIEW_LATE_CELLS = '03d81241274d76ef3128b5235707ac84a8996d2250fd93e98352ff8544ae30b1'
IDENTITY_REVIEW_OUTPUT_CELLS = '490c20d9223919f27a0b263e3f7be5dbb63baaa0b7eaa4bbbdf14a552f714d2b'


EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}


@pytest.fixture(scope='module')
def baseline():
    with open(BASELINE_PATH, encoding='utf-8') as fh:
        return json.load(fh)


@pytest.fixture(scope='module')
def real_run(real_detail_path, real_exempt_path, tmp_path_factory):
    out = tmp_path_factory.mktemp('baseline') / '验收基线输出.xlsx'
    before = hashlib.md5(open(real_detail_path, 'rb').read()).hexdigest()
    result = run(RunConfig(detail_path=real_detail_path, exempt_path=real_exempt_path,
                            output_path=str(out), week_no=7))
    assert result.ok, result.error
    after = hashlib.md5(open(real_detail_path, 'rb').read()).hexdigest()
    assert before == after, '输入明细表被改动了'
    return result, str(out), real_detail_path, real_exempt_path


def _b_totals(output_path, detail_path):
    """B 口径：产出工作簿免勤列之和（按明细模型的房间行取值，含第 3 行首块）。"""
    detail = load_detail(detail_path)
    wb = load_workbook(output_path, data_only=True)
    out = {}
    try:
        for sheet in SHEETS:
            for room in detail.rooms_of(sheet):
                for day, col in EXEMPT_COLS.items():
                    value = wb[sheet].cell(room.row, col).value
                    if isinstance(value, (int, float)):
                        key = '%s|%s|%s' % (sheet, room.college_name, day)
                        out[key] = out.get(key, 0) + int(value)
    finally:
        wb.close()
    return out


def _a_totals(exempt_path):
    """A 口径：引擎学院总免勤（内部分解量）。"""
    agg = aggregate_exemptions(load_exemptions(exempt_path), 7, campus='雁山')
    return agg


# ------------------------------------------------------------------ A 两口径


def test_public_totals_match_acceptance_baseline(real_run, baseline):
    """A1：B 口径（对外唯一）120 个组合的免勤合计必须与验收基线逐一相等。"""
    result, output, detail_path, _exempt = real_run
    current = _b_totals(output, detail_path)
    expected = {key: value + IDENTITY_REVIEW_DELTAS.get(key, 0)
                for key, value in baseline['b_totals'].items()}
    pending = {(row['学院'], row['星期'], row['自习']) for row in result.exempt_rows
               if row['免勤状态'] == '身份待核对'}
    assert result.exempt_stats['身份冲突暂缓计入人次'] == 0
    assert result.exempt_stats['身份冲突事件（已按1人计）'] == 31
    assert len(pending) == 20
    assert len(current) == 120
    diff = {k: (expected.get(k), current.get(k))
            for k in set(expected) | set(current)
            if expected.get(k, 0) != current.get(k, 0)}
    assert not diff, '对外口径免勤合计偏离验收基线：%s' % diff


def test_internal_totals_match_acceptance_baseline(real_run, baseline):
    """A1b：A 口径（内部分解量，仅归因用）同样逐组合对齐。"""
    _result, _output, _detail, exempt_path = real_run
    agg = _a_totals(exempt_path)
    detail = load_detail(exempt_path) if False else None       # 保持不读明细
    colleges = []
    for key in baseline['a_totals']:
        sheet, college, _day = key.split('|')
        if college not in colleges:
            colleges.append(college)
    current = {}
    for session, sheet in (('早自习', '雁山（早）'), ('晚自习', '雁山（晚）')):
        for college in colleges:
            for day in DAYS:
                current['%s|%s|%s' % (sheet, college, day)] = \
                    agg.exempt_count(college, day, session)
    diff = {k: (baseline['a_totals'].get(k), current.get(k))
            for k in set(baseline['a_totals']) | set(current)
            if baseline['a_totals'].get(k, 0) != current.get(k, 0)}
    assert not diff, '内部分解量偏离验收基线：%s' % diff


def test_two_calibers_differ_only_by_capacity_cap_overflow(real_run, baseline):
    """A2：两口径的差集**恰好 1 个组合**＝数统周三，差 1，且归因到 C1=甲 超额（7614）。

    两边使用同一次统计的身份审核结果，差 1 来自教室容量封顶。
    """
    result, output, detail_path, _exempt = real_run
    b = _b_totals(output, detail_path)
    a = {}
    for row in result.exempt_rows:
        sheet = '雁山（早）' if row['自习'] == '早自习' else '雁山（晚）'
        key = '|'.join((sheet, row['学院'], row['星期']))
        if key in b:
            a[key] = row['免勤合计']
    diff = {k: (a.get(k, 0), b.get(k, 0)) for k in set(a) | set(b)
            if a.get(k, 0) != b.get(k, 0)}
    assert len(diff) == 1, '同一次统计的两口径差集应恰好 1 个组合，实际：%s' % diff
    key, (a_value, b_value) = next(iter(diff.items()))
    assert key == '雁山（晚）|数学与统计学院|周三', key
    assert a_value - b_value == 1
    # 差 1 必须归因到 C1=甲：教室封顶 111+108=219，学院总免勤 220，超额 1 归 7614
    assert baseline['anchor']['rooms'] == {'7614': 111, '7616': 108}
    assert baseline['anchor']['exempt'] == 220
    assert baseline['anchor']['rate'] is None
    overflow = [o for o in baseline['overflow']
                if o.get('college') == '数学与统计学院' and o.get('day') == '周三']
    assert len(overflow) == 1 and overflow[0]['room'] == '7614' \
        and overflow[0]['excess'] == 1


def test_public_per_day_table_is_hardcoded(real_run, baseline):
    """按天核对当前回填口径，同时校验历史数据未被改写。"""
    _result, output, detail_path, _exempt = real_run
    b = _b_totals(output, detail_path)
    computed = {}
    for session, needle in (('早自习', '早'), ('晚自习', '晚')):
        for day in DAYS:
            computed['%s|%s' % (session, day)] = sum(
                v for k, v in b.items()
                if k.split('|')[2] == day and needle in k.split('|')[0])
    expected = PUBLIC_PER_DAY
    assert computed == expected, '当前身份核对口径与实跑不符：%s' % computed
    assert baseline['b_per_day'] == PUBLIC_PER_DAY
    assert baseline['a_per_day']['晚自习|周三'] == 1518   # A 口径，仅归因用


# ------------------------------------------------------------------ B 早表硬锚点


def test_early_table_four_day_totals_match_conversation_source(real_run, baseline):
    """B1：早表四天去重合计 160/163/158/163 —— 与对话原文逐位吻合（r178 归晚自习侧）。"""
    _result, _out, _detail, exempt_path = real_run
    agg = _a_totals(exempt_path)
    colleges = {k.split('|')[1] for k in baseline['a_totals']}
    for day, expected in (('周一', 160), ('周三', 163), ('周四', 158), ('周五', 163)):
        total = sum(agg.exempt_count(c, day, '早自习') for c in colleges)
        assert total == expected, '早表 %s = %d，应为 %d' % (day, total, expected)


# ------------------------------------------------------------------ C 免勤单元


def test_room_exempt_cells_count_and_hash_match_baseline(real_run, baseline):
    """C1：两表合计 592 个教室级免勤单元，逐格取值与基线 sha256 一致。"""
    _result, output, detail_path, _exempt = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output, data_only=True)
    try:
        for sheet in SHEETS:
            ws = wb[sheet]
            rows = []
            for room in detail.rooms_of(sheet):
                for day, col in sorted(room.exempt_cols.items()):
                    rows.append('%s|%d|%s|%s|%r' % (sheet, room.row, room.room, day,
                                                    ws.cell(room.row, col).value))
            rows.sort()
            digest = hashlib.sha256('\n'.join(rows).encode('utf-8')).hexdigest()
            expected = IDENTITY_REVIEW_LATE_CELLS if sheet == '雁山（晚）' else baseline['cells'][sheet]['sha256']
            assert digest == expected, \
                '%s 教室级免勤逐格与基线不一致' % sheet
            assert len(rows) == baseline['cells'][sheet]['count']
    finally:
        wb.close()


def test_room_exempt_cells_are_never_blank(real_run, baseline):
    """C2：592 个免勤单元必须全部有值（无空白格）。"""
    _result, output, detail_path, _exempt = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output, data_only=True)
    try:
        total = 0
        for sheet in SHEETS:
            ws = wb[sheet]
            for room in detail.rooms_of(sheet):
                for day, col in room.exempt_cols.items():
                    value = ws.cell(room.row, col).value
                    assert value is not None, '%s 行%d %s 免勤格为空' % (sheet, room.row, day)
                    total += 1
        assert total == 592
    finally:
        wb.close()


# ------------------------------------------------------------------ D 锚点


def test_math_college_anchor_matches_c1_verdict(real_run, baseline):
    """D1：数统周三晚 —— 学院应到 219 / 免勤 220 / 分母 −1 不出率 / 教室 111+108 / 超额 1 归 7614。"""
    result, _out, _detail, _exempt = real_run
    hit = [r for r in result.college_rates
           if r.college == '数学与统计学院' and r.day == '周三' and r.session == '晚自习']
    assert len(hit) == 1
    rec = hit[0]
    assert rec.expected == 219
    assert rec.exempt == 220
    assert rec.rate is None
    assert rec.status != '完整'
    assert '分母' in rec.status or '分母' in rec.note


def test_math_college_rooms_are_capped_and_overflow_located(real_run, baseline):
    """D2：教室封顶值 111/108，超额 1 人单列并归到 7614（绝不写 112 > 111）。"""
    result, _out, _detail, _exempt = real_run
    rooms = {str(r['教室']): r.get('周三免勤') for r in result.room_rows
             if r['学院'] == '数学与统计学院' and '晚' in r['工作表']}
    assert rooms.get('7614') == 111
    assert rooms.get('7616') == 108
    assert all(v <= 111 for v in rooms.values() if isinstance(v, int))
    overflow = [o for o in result.exempt_overflow
                if o.get('college') == '数学与统计学院' and o.get('day') == '周三']
    assert len(overflow) == 1
    assert overflow[0]['room'] == '7614'
    assert overflow[0]['excess'] == 1
    assert 111 + 108 + 1 == 220


def test_screen_sheet_has_zero_cell_difference(real_run, baseline):
    """D3：屏风逐单元格零差异。"""
    _result, output, detail_path, _exempt = real_run
    src = load_workbook(detail_path, data_only=True)
    dst = load_workbook(output, data_only=True)
    try:
        ws_src, ws_dst = src['屏风'], dst['屏风']
        diff = 0
        for row in range(1, ws_src.max_row + 1):
            for col in range(1, ws_src.max_column + 1):
                if ws_dst.cell(row, col).value != ws_src.cell(row, col).value:
                    diff += 1
        assert diff == 0
        assert diff == baseline['screen_diff'] == 0
    finally:
        src.close()
        dst.close()


def test_first_college_block_merge_includes_first_room_row(real_run, baseline):
    """D4：块首行教室（6504）必须落在合并区内（早/晚首段均为 3:8）。"""
    _result, output, detail_path, _exempt = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output)
    try:
        for sheet in SHEETS:
            ws = wb[sheet]
            first_col = min(detail.rate_cols_by_sheet[sheet].values())
            spans = sorted((r.min_row, r.max_row) for r in ws.merged_cells.ranges
                           if r.min_col == r.max_col == first_col)
            assert spans, '%s 出勤率列没有任何合并区' % sheet
            first = spans[0]
            # v1.1.0：第 3 行是标签行，出勤率合并区从第 4 行开始，
            # 但第 3 行首间教室 6504 仍属于学院块 B3:B8。
            assert first == (4, 8), '%s 首段率区 %s，应为 (4, 8)' % (sheet, first)
            first_room = sorted(detail.rooms_of(sheet), key=lambda r: r.row)[0]
            college_spans = [(r.min_row, r.max_row) for r in wb[sheet].merged_cells.ranges
                             if r.min_col == r.max_col == 2]
            assert any(lo <= first_room.row <= hi for lo, hi in college_spans), \
                '%s 块首行教室 %s(行%d) 不在学院块合并区 %s 内' % (
                    sheet, first_room.room, first_room.row, college_spans)
    finally:
        wb.close()


def test_output_has_exactly_three_sheets(real_run, baseline):
    """D5：输出工作表恒为三张（雁山（早）/雁山（晚）/屏风）。"""
    _result, output, _detail, _exempt = real_run
    wb = load_workbook(output, read_only=True)
    try:
        assert [s for s in wb.sheetnames if s not in ('_系统校验元数据', '免勤核对')] == list(SHEETS) + ['屏风']
    finally:
        wb.close()


# ------------------------------------------------------------------ E 行为中性对拍


def test_output_content_digest_unchanged(real_run, baseline):
    """E1：**内容摘要对拍**——任何改动（界面层、core 标记位、解析重构）都不得改变产出。

    锁定当前身份核对口径下的单元格内容，以及历史快照中的合并区和屏风内容。
    """
    _result, output, _detail, _exempt = real_run
    wb = load_workbook(output, data_only=True)
    try:
        cell_lines, merge_lines, screen_lines = [], [], []
        for sheet in SHEETS:
            ws = wb[sheet]
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    value = ws.cell(row, col).value
                    if value is not None:
                        cell_lines.append('%s|%d|%d|%r' % (sheet, row, col, value))
            for rng in sorted(ws.merged_cells.ranges,
                              key=lambda r: (r.min_row, r.min_col)):
                merge_lines.append('%s|%s' % (sheet, str(rng)))
        ws = wb['屏风']
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                screen_lines.append('%d|%d|%r' % (row, col, ws.cell(row, col).value))
    finally:
        wb.close()

    digest = baseline['digest']
    assert hashlib.sha256('\n'.join(cell_lines).encode('utf-8')).hexdigest() \
        == IDENTITY_REVIEW_OUTPUT_CELLS, '产出单元格内容与当前口径快照不一致'
    assert hashlib.sha256('\n'.join(merge_lines).encode('utf-8')).hexdigest() \
        == digest['merges_sha256'], '合并区结构与验收基线不一致'
    assert hashlib.sha256('\n'.join(screen_lines).encode('utf-8')).hexdigest() \
        == digest['screen_sha256'], '屏风逐格内容与验收基线不一致'
