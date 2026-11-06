# -*- coding: utf-8 -*-
"""归因文档验收门（R16 / R19 / R21 / R23 + 需求方遗留待办）。

验收项（Chief 点名）：
  ① 六类映射完整：容量封顶差 / 无明细列 / 全班无人数 / 全学期 / indel 合并 / tie-break 变更
  ② 每个差异有行级定位（学院 × 星期 × 源行号 + 脱敏学号）
  ③ 预期差异 与 缺陷 **分列**
  ④ 「口径变更」独立栏
  ⑤ 无归因差异数 = 0
  ⑥ §四不得再留「孰优孰劣由需求方裁定」（已裁定：本版照常计入并标记正确，上一轮静默丢弃才是缺陷）
  ⑦ 数字必须与**独立复算**一致（R21：对外只认产出工作簿的学院×星期合计）

默认不进打包闸门（文档层收尾不阻塞发布）；需要验收时运行：
    python -X utf8 -m pytest tests -m docgate -q
"""

from __future__ import annotations

import os
import re
from collections import defaultdict

import pytest

pytestmark = pytest.mark.docgate

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DOC_PATH = os.path.join(PROJECT_DIR, 'docs', '第2轮_对拍差异归因.md')
PREV_ROUND = os.path.join(os.path.dirname(PROJECT_DIR), 'Week7_Attendance_Final.xlsx')

#: R16 的六个归因类目 → 允许的书写形式
SIX_CATEGORIES = {
    '容量封顶差': ('容量封顶', '按容量封顶'),
    '无明细列': ('无明细列', '无处落表'),
    '全班无人数': ('全班无人数', '待补人数', '全班'),
    '全学期': ('全学期',),
    'indel 合并': ('indel',),
    'tie-break 变更': ('tie-break', '并列'),
}

EXEMPT_COLS = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}


# ------------------------------------------------------------------ 工具


def _doc_text():
    if not os.path.exists(DOC_PATH):
        pytest.skip('归因文档尚未产出：%s（文档层待办，不阻塞发布）' % DOC_PATH)
    with open(DOC_PATH, encoding='utf-8') as fh:
        return fh.read()


def _headings_and_headers(text):
    """返回（标题行集合, 表头行集合），用于判定"独立栏/成节"而非散文提及。"""
    heads, headers = [], []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith('#'):
            heads.append(stripped)
        elif stripped.startswith('|') and stripped.count('|') >= 3:
            headers.append(stripped)
    return heads, headers


def _college_of(ws, row):
    col = 2
    for rng in ws.merged_cells.ranges:
        if rng.min_col <= col <= rng.max_col and rng.min_row <= row <= rng.max_row:
            raw = ws.cell(rng.min_row, rng.min_col).value
            return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(raw or '')).strip()
    raw = ws.cell(row, col).value
    return re.sub(r'[（(]\s*\d+\s*人\s*[)）]', '', str(raw or '')).strip()


def _room_sums(path, sheet):
    """按 (学院, 星期) 汇总教室级免勤（与 tools/diagnose_diff.py 同口径）。"""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb[sheet]
    out = defaultdict(int)
    for row in range(4, ws.max_row + 1):
        college = _college_of(ws, row)
        room = ws.cell(row, 3).value
        if not college or room is None:
            continue
        for day, col in EXEMPT_COLS.items():
            value = ws.cell(row, col).value
            if isinstance(value, (int, float)):
                out[(college, day)] += int(value)
    wb.close()
    return out


def _parse_diff_table(text):
    """从文档里解析「工作表 / 学院 / 星期 × 本次 / 基线 / 差」明细表。"""
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if not line.strip().startswith('|'):
            continue
        cells = [c.strip() for c in line.strip().strip('|').split('|')]
        if not all(cells):
            continue
        if not ('工作表' in cells and '学院' in cells and '星期' in cells):
            continue

        def find_col(*needles):
            for idx, head in enumerate(cells):
                if any(n in head for n in needles):
                    return idx
            return None

        i_sheet = find_col('工作表')
        i_college = find_col('学院')
        i_day = find_col('星期')
        i_cur = find_col('本次', '产出', '本版')
        i_base = find_col('基线', '上一轮', '对照')
        i_delta = find_col('差', 'Δ')
        if None in (i_sheet, i_college, i_day, i_cur, i_base):
            continue
        rows = []
        for line2 in lines[i + 1:]:
            if not line2.strip().startswith('|'):
                break
            row = [c.strip() for c in line2.strip().strip('|').split('|')]
            if len(row) <= max(i_sheet, i_college, i_day, i_cur, i_base):
                continue
            try:
                cur = int(row[i_cur])
                base = int(row[i_base])
            except ValueError:
                continue
            delta = None
            if i_delta is not None and len(row) > i_delta:
                try:
                    delta = int(row[i_delta])
                except ValueError:
                    delta = None
            rows.append({'sheet': row[i_sheet], 'college': row[i_college],
                         'day': row[i_day], 'cur': cur, 'base': base, 'delta': delta})
        if rows:
            return rows
    return None


def _recompute(real_detail_path, real_exempt_path, tmp_path):
    """独立复算：本次产物 vs 上一轮产物，逐 (工作表, 学院, 星期) 的免勤合计。"""
    from core import RunConfig, run
    out = str(tmp_path / '本次产物.xlsx')
    result = run(RunConfig(detail_path=real_detail_path, exempt_path=real_exempt_path,
                            output_path=out, week_no=7))
    assert result.ok, result.error
    mine, base = {}, {}
    for sheet in ('雁山（早）', '雁山（晚）'):
        for (college, day), value in _room_sums(out, sheet).items():
            mine[(sheet, college, day)] = value
        for (college, day), value in _room_sums(PREV_ROUND, sheet).items():
            base[(sheet, college, day)] = value
    keys = set(mine) | set(base)
    diff = {k: (mine.get(k, 0), base.get(k, 0)) for k in keys
            if mine.get(k, 0) != base.get(k, 0)}
    return mine, base, diff


# ------------------------------------------------------------------ 验收项


def test_doc_has_all_six_attribution_categories():
    """① 六类映射完整。"""
    text = _doc_text()
    missing = [name for name, words in SIX_CATEGORIES.items()
               if not any(w in text for w in words)]
    assert not missing, '归因文档缺少类目：%s' % missing


def test_doc_separates_expected_differences_from_defects():
    """③ 预期差异 与 缺陷 分列（两者都要成栏/成节，不能只塞在散文里）。"""
    text = _doc_text()
    heads, headers = _headings_and_headers(text)
    for word in ('预期差异', '缺陷'):
        assert word in text, '文档缺少「%s」' % word
        assert any(word in h for h in heads + headers), \
            '「%s」未独立成栏/成节' % word


def test_doc_has_standalone_caliber_change_section():
    """④ 「口径变更」独立栏（R21/R23：三套数字并存时必须单列口径差异）。"""
    text = _doc_text()
    heads, headers = _headings_and_headers(text)
    assert '口径变更' in text
    assert any('口径变更' in h for h in heads + headers), \
        '「口径变更」未独立成栏/成节'


def test_doc_has_row_level_attribution_columns():
    """② 行级定位：必须有「源行」与「学号」两列（脱敏）。"""
    text = _doc_text()
    _heads, headers = _headings_and_headers(text)
    assert any('源行' in h for h in headers), '缺少「源行」列'
    assert any('学号' in h for h in headers), '缺少「学号」列'


def test_doc_has_no_unattributed_differences():
    """⑤ 无归因差异数 = 0。"""
    text = _doc_text()
    zero_lines = [l for l in text.splitlines()
                  if '未归因' in l and re.search(r'\b0\b', l)]
    assert zero_lines or ('未归因' in text and '无未归因' in text), \
        '文档未声明「未归因差异数 = 0」'


def test_doc_has_no_undecided_verdict_paragraph():
    """⑥ §四不得再留已被裁定的「孰优孰劣由需求方裁定」。"""
    text = _doc_text()
    assert '孰优孰劣由需求方裁定' not in text, \
        '结论章仍留着已被裁定的「孰优孰劣由需求方裁定」'


def test_doc_numbers_match_independent_recomputation(real_detail_path, real_exempt_path,
                                                     tmp_path):
    """⑦ 数字必须与独立复算一致，且差异组合数符合 R19（6 个组合）。"""
    if not os.path.exists(PREV_ROUND):
        pytest.skip('上一轮产物不在预期位置：%s' % PREV_ROUND)
    text = _doc_text()
    rows = _parse_diff_table(text)
    assert rows, '文档缺少「工作表 / 学院 / 星期 × 本次 / 基线 / 差」明细表（R19/R21 要求）'
    mine, base, diff = _recompute(real_detail_path, real_exempt_path, tmp_path)

    assert len(diff) == 6, '独立复算得到 %d 个差异组合，R19 裁定为 6 个' % len(diff)
    assert len(rows) == len(diff), \
        '文档列出 %d 行差异，独立复算为 %d 个组合' % (len(rows), len(diff))

    for row in rows:
        key = (row['sheet'], row['college'], row['day'])
        assert key in diff, '文档多出差异组合：%s' % (key,)
        assert row['cur'] == mine.get(key, 0), \
            '%s 本次合计 %d，独立复算为 %d' % (key, row['cur'], mine.get(key, 0))
        assert row['base'] == base.get(key, 0), \
            '%s 基线合计 %d，独立复算为 %d' % (key, row['base'], base.get(key, 0))
        if row['delta'] is not None:
            assert row['delta'] == row['cur'] - row['base'], \
                '%s 差值 %d 与 本次−基线 不符' % (key, row['delta'])

    doc_keys = {(r['sheet'], r['college'], r['day']) for r in rows}
    assert doc_keys == set(diff), \
        '文档的差异组合与独立复算不一致：缺 %s / 多 %s' % (set(diff) - doc_keys,
                                                          doc_keys - set(diff))


def test_doc_does_not_quote_non_public_totals_as_headline(real_detail_path,
                                                           real_exempt_path, tmp_path):
    """R23：对外只认产出工作簿的学院×星期合计；引用其它口径必须附「口径变更」说明。"""
    text = _doc_text()
    for token in ('1888', '1265', '1891', '1517'):
        if token in text:
            assert '口径变更' in text, \
                '文档引用了非对外口径数字 %s，但没有口径变更说明' % token
