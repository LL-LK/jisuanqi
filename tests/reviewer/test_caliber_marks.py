# -*- coding: utf-8 -*-
"""R 系列口径断言（需求方 R1/R2/R5/R6/R7 在真实工作簿上的落地校验）。

这六条是需求方在第 1 轮点名的必测项：
  R1 估算样式标记（教室级免勤=估算 ⇒ 固定斜体 + 固定填充色；非估算无标记；禁文字后缀/批注/新增列）
  R2 缺失实到 ⇒ 导出**真空白**（不写 0、不写「—」、不写占位符）
  R5 免勤星期无对应明细列 ⇒ 逐条进清单，且**不计入任何日**（严禁并入其它日）
  R6 无人数「全班」集体申请 ⇒ 不覆盖个人、不计免勤、11 行全进清单
  R7 块首行教室（6504）必须在合并区内
"""

from __future__ import annotations

import pytest
from openpyxl import load_workbook

from core import RunConfig, run
from core.io_detail import load_detail

pytestmark = pytest.mark.realdata

SHEETS = ('雁山（早）', '雁山（晚）')


@pytest.fixture(scope='module')
def real_run(real_detail_path, real_exempt_path, tmp_path_factory):
    out = tmp_path_factory.mktemp('caliber') / '口径输出.xlsx'
    result = run(RunConfig(detail_path=real_detail_path, exempt_path=real_exempt_path,
                            output_path=str(out), week_no=7))
    assert result.ok, result.error
    return result, str(out), real_detail_path


def _anomalies(result, kind):
    return [a for a in result.anomalies if a.get('type') == kind]


def _rate_row(detail, sheet, block):
    """出勤率写入行：有标签行时从标签行下一行开始，否则取块首行。"""
    rows = block.merge_rows
    label = detail.rate_label_rows.get(sheet, 0)
    return max(min(rows), label + 1) if label else min(rows)


# ------------------------------------------------------------------ R1 估算标记

#: R17 固化的估算标记色值（与 core/export.py 的 ESTIMATE_FILL 同源）
ESTIMATE_FILL_RGB = 'FFFFF3CD'


def _fill_rgb(cell):
    fill = cell.fill
    if fill is None:
        return None
    rgb = getattr(fill.start_color, 'rgb', None)
    return rgb if isinstance(rgb, str) else None


def test_room_exempt_cells_are_marked_as_estimated(real_run):
    """R1/R17：斜体 + 指定色值 FFFFF3CD + **与模板底色不同**。

    真实免勤表 3372 条**全部无教室号** ⇒ 教室级免勤全部为估算 ⇒ 592 格必须全部带标记。
    """
    result, output, detail_path = real_run
    detail = load_detail(detail_path)
    src = load_workbook(detail_path)
    wb = load_workbook(output)
    try:
        total = 0
        marked = 0
        for sheet in SHEETS:
            ws = wb[sheet]
            for room in detail.rooms_of(sheet):
                for _day, col in room.exempt_cols.items():
                    cell = ws.cell(room.row, col)
                    total += 1
                    if not (cell.font is not None and cell.font.italic):
                        continue
                    marked += 1
                    assert _fill_rgb(cell) == ESTIMATE_FILL_RGB, \
                        '%s 行%d 估算格填充色 %s，应为 %s' % (sheet, room.row,
                                                           _fill_rgb(cell), ESTIMATE_FILL_RGB)
                    template_rgb = _fill_rgb(src[sheet].cell(room.row, col))
                    assert _fill_rgb(cell) != template_rgb, \
                        '%s 行%d 估算色与模板底色相同（%s），标记不可辨识' % (sheet, room.row,
                                                                            template_rgb)
        assert total == 592
        assert marked == total, '只有 %d/%d 个教室级免勤格带估算标记（应为 592）' % (marked, total)
    finally:
        src.close()
        wb.close()


def test_zero_valued_exempt_cells_are_marked_as_estimated(real_run):
    """R17：**值为 0 的免勤格同样带估算标记**（0 也是比例分摊结果，不是逐人核对过的名单）。

    这条最容易被后人当 bug「优化」掉（0 看着像没数据），必须固化成断言。
    """
    _result, output, detail_path = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output)
    try:
        zeros = 0
        marked = 0
        for sheet in SHEETS:
            ws = wb[sheet]
            for room in detail.rooms_of(sheet):
                for _day, col in room.exempt_cols.items():
                    cell = ws.cell(room.row, col)
                    if cell.value != 0:
                        continue
                    zeros += 1
                    if cell.font is not None and cell.font.italic \
                            and _fill_rgb(cell) == ESTIMATE_FILL_RGB:
                        marked += 1
        assert zeros > 0, '合成/真实数据里应存在值为 0 的免勤格'
        assert marked == zeros, '有 %d/%d 个 0 值免勤格没带估算标记' % (zeros - marked, zeros)
    finally:
        wb.close()


def test_screen_sheet_has_no_estimate_mark(real_run):
    """R1：屏风表逐单元格零改动 ⇒ 不得出现任何估算标记。"""
    _result, output, detail_path = real_run
    src = load_workbook(detail_path)
    wb = load_workbook(output)
    try:
        ws_src, ws_dst = src['屏风'], wb['屏风']
        for row in range(1, ws_src.max_row + 1):
            for col in range(1, ws_src.max_column + 1):
                assert not (ws_dst.cell(row, col).font and ws_dst.cell(row, col).font.italic), \
                    '屏风表 (%d,%d) 出现了估算标记' % (row, col)
    finally:
        src.close()
        wb.close()


def test_rate_cells_are_not_marked_as_estimated(real_run):
    """R1：出勤率格不是估算值 ⇒ 不得带估算标记。"""
    _result, output, detail_path = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output)
    try:
        checked = 0
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if not rows:
                    continue
                for _label, col in rate_cols.items():
                    cell = ws.cell(_rate_row(detail, sheet, block), col)
                    assert not (cell.font is not None and cell.font.italic), \
                        '%s 出勤率格被误标为估算' % sheet
                    checked += 1
        assert checked > 0
    finally:
        wb.close()


def test_estimate_mark_is_not_written_as_text_or_comment(real_run):
    """R1：禁文字后缀、禁批注 —— 免勤格值必须是纯数字。"""
    _result, output, detail_path = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output)
    try:
        for sheet in SHEETS:
            ws = wb[sheet]
            for room in detail.rooms_of(sheet):
                for day, col in room.exempt_cols.items():
                    value = ws.cell(room.row, col).value
                    assert isinstance(value, int), \
                        '%s 行%d %s 免勤格不是整数：%r' % (sheet, room.row, day, value)
                    assert ws.cell(room.row, col).comment is None
    finally:
        wb.close()


def test_no_extra_column_added_for_estimate(real_run):
    """R1：禁新增列 —— 输出工作表列数必须与输入一致。"""
    _result, output, detail_path = real_run
    src = load_workbook(detail_path, read_only=True)
    dst = load_workbook(output, read_only=True)
    try:
        for sheet in SHEETS:
            assert dst[sheet].max_column == src[sheet].max_column, \
                '%s 列数变了（疑似为估算新增了列）' % sheet
    finally:
        src.close()
        dst.close()


# ------------------------------------------------------------------ R2 真空白


def test_missing_actual_is_truly_blank_in_export(real_run):
    """R2：实到缺失 ⇒ 出勤率格真空白（None），不得写 0、不得写「—」。"""
    result, output, detail_path = real_run
    detail = load_detail(detail_path)
    by_key = {(r.sheet, r.college, r.day): r for r in result.college_rates}
    wb = load_workbook(output)
    try:
        blanks = 0
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            ws = wb[sheet]
            for block in detail.blocks_of(sheet):
                rows = block.merge_rows
                if not rows:
                    continue
                for label, col in rate_cols.items():
                    if label == '平均':
                        continue
                    rec = by_key.get((sheet, block.name, label))
                    if rec is None or rec.actual is not None:
                        continue
                    value = ws.cell(_rate_row(detail, sheet, block), col).value
                    assert value is None, \
                        '实到缺失却写了 %r（%s / %s / %s）' % (value, sheet, block.name, label)
                    blanks += 1
        assert blanks > 0, '真实数据里应存在实到缺失的学院×天组合'
    finally:
        wb.close()


def test_missing_actual_is_listed(real_run):
    """R2：实到缺失必须进异常清单（不得静默）。"""
    result, _output, _detail = real_run
    assert _anomalies(result, '实到数据未完整填写'), '实到缺失未进异常清单'


# ------------------------------------------------------------------ R5 无列星期


def test_days_without_detail_column_are_listed(real_run):
    """R5：免勤表有、但明细表无该日列 ⇒ 逐条进清单（早表周二 / 晚表周二 / 晚表周五）。"""
    result, _output, _detail = real_run
    nowhere = _anomalies(result, '该天免勤无处落表')
    not_in_rate = _anomalies(result, '该天免勤已填表但不参与出勤率计算')
    scopes = {a.get('scope', '') for a in nowhere}
    assert '雁山（早） / 周二' in scopes, scopes
    assert '雁山（晚） / 周二' in scopes, scopes
    assert any('周五' in a.get('scope', '') for a in not_in_rate), \
        '晚表周五免勤未进清单'


def test_days_without_detail_column_do_not_leak_into_other_days(real_run):
    """R5：周二/周五的免勤严禁并入周一/周三/周四任一计算。"""
    result, output, detail_path = real_run
    detail = load_detail(detail_path)
    by_key = {(r.sheet, r.college, r.day): r for r in result.college_rates}
    for sheet, rate_cols in detail.rate_cols_by_sheet.items():
        assert '周二' not in rate_cols, '%s 竟出现了周二率列' % sheet
    for (sheet, college, day), rec in by_key.items():
        assert day in ('周一', '周三', '周四', '周五')
        if day == '周五' and sheet == '雁山（晚）':
            assert rec.rate is None, '晚表不该有周五率'


def test_late_friday_exempt_is_written_but_not_counted_in_rate(real_run):
    """R5：晚表周五免勤**写入周五免勤列**，但因该表无周五出勤率列，不参与率计算。

    数据核对员取证：晚表 L/M 两列「表头在、整表 0 格数据」⇒ 属模板结构缺失，
    引擎把 19 人次写进周五免勤列（留痕可查），并在清单注明「不参与出勤率计算」。
    """
    result, output, detail_path = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output, data_only=True)
    try:
        written = 0
        for room in detail.rooms_of('雁山（晚）'):
            col = room.exempt_cols.get('周五')
            if not col:
                continue
            value = wb['雁山（晚）'].cell(room.row, col).value
            if value:
                written += int(value)
        assert written == 19, '晚表周五免勤列合计 %d，应为 19' % written
        # 晚表没有周五出勤率列 ⇒ 不得产出周五率
        assert '周五' not in detail.rate_cols_by_sheet['雁山（晚）']
        for rec in result.college_rates:
            if rec.sheet == '雁山（晚）':
                assert rec.day != '周五'
    finally:
        wb.close()


# ------------------------------------------------------------------ R6 全班无人数


def test_whole_class_rows_are_all_listed(real_run):
    """R6：11 行「全班」占位申请必须逐行进清单（行级条目 + 受影响学院×天条目）。"""
    result, _output, _detail = real_run
    items = _anomalies(result, '免勤不完整（待补人数）')
    row_level = [a for a in items if '源行' in a.get('scope', '')]
    college_level = [a for a in items if '源行' not in a.get('scope', '')]
    assert len(row_level) == 11, '行级「全班」条目 %d 条，应为 11' % len(row_level)
    assert len(college_level) >= 1, '缺少「免勤不完整」的学院×天标注'


def test_whole_class_rows_do_not_cover_personal(real_run):
    """R6/R12：「全班」缺人数 ⇒ 不计入免勤、不覆盖同班个人；该学院该次状态强制为「免勤不完整」。"""
    result, _output, _detail = real_run
    hit = [r for r in result.college_rates
           if r.college == '商学院' and r.day == '周一' and r.session == '晚自习']
    assert hit
    rec = hit[0]
    assert rec.status != '完整', '商学院周一晚存在无人数占位申请，状态不得为「完整」'
    assert '不完整' in rec.status or '待补人数' in rec.status or \
        '不完整' in rec.note or '待补人数' in rec.note


def test_whole_class_records_never_enter_sid_keyspace(real_run, real_exempt_path):
    """R6：11 行「全班」占位申请不得进入去重键空间（不得被当成一个人、不得贡献 1 人次）。"""
    from core.io_exempt import aggregate_exemptions, load_exemptions
    records = load_exemptions(real_exempt_path)
    placeholders = [r for r in records
                    if (r.sid or '').strip() in ('全班',) or
                    ('全班' in (r.name or '') and not r.group_size)]
    assert len(placeholders) == 11, '占位申请 %d 条，应为 11' % len(placeholders)
    for r in placeholders:
        assert not r.group_size, '占位申请不得带人数：%r' % (r.group_size,)

    agg = aggregate_exemptions(records, 7, campus='雁山')
    key = ('商学院', '周一', '晚自习')
    # 口径说明：该桶原始不同学号 89 个，引擎计数 88 —— 差 1 是 R11 的 indel 合并
    # （仅删/仅插一位视为同一人），属已裁定口径，不是占位申请造成的。
    # 因此这里用「移除 11 行占位后计数不变」来证明占位申请贡献 0 人次。
    without = [r for r in records if r not in placeholders]
    agg_without = aggregate_exemptions(without, 7, campus='雁山')
    assert agg.exempt_count(*key) == agg_without.exempt_count(*key), \
        '占位申请被计入了免勤人数'
    assert agg.exempt_count(*key) == 88


# ------------------------------------------------------------------ R7 块首行


def test_first_block_room_is_inside_merge_region(real_run):
    """R7：块首行教室（6504）必须落在本表第一个合并区内。"""
    _result, output, detail_path = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output)
    try:
        for sheet in SHEETS:
            ws = wb[sheet]
            first_col = min(detail.rate_cols_by_sheet[sheet].values())
            spans = sorted((r.min_row, r.max_row) for r in ws.merged_cells.ranges
                           if r.min_col == r.max_col == first_col)
            assert spans[0] == (4, 8), \
                '%s 首段率区 %s，应为 (4, 8)（第3行保留星期标签）' % (sheet, spans[0])
            first_room = sorted(detail.rooms_of(sheet), key=lambda r: r.row)[0]
            college_spans = [(r.min_row, r.max_row) for r in ws.merged_cells.ranges
                             if r.min_col == r.max_col == 2]
            assert any(lo <= first_room.row <= hi for lo, hi in college_spans), \
                '%s 块首行教室 %s(行%d) 不在学院块合并区 %s 内' % (
                    sheet, first_room.room, first_room.row, college_spans)
    finally:
        wb.close()


# ------------------------------------------------------------------ R22 晚表周五留痕


def test_late_friday_trace_cells_are_marked_and_excluded(real_run):
    """R22：晚表周五免勤 19 人次是**留痕**，边界是「有值 ≠ 出数」。

    * 留痕格必须带 R1 估算标记（斜体 + FFFFF3CD）；
    * 该日不产出教室率、不产出学院率、**不进平均列**；
    * 晚表周五实到整表为空 ⇒ 依 R2 不得凭空生成实到。
    """
    result, output, detail_path = real_run
    detail = load_detail(detail_path)
    wb = load_workbook(output, data_only=True)
    try:
        ws = wb['雁山（晚）']
        marked = 0
        written = 0
        for room in detail.rooms_of('雁山（晚）'):
            col = room.exempt_cols.get('周五')
            if not col:
                continue
            cell = ws.cell(room.row, col)
            if cell.value:
                written += int(cell.value)
                assert cell.font is not None and cell.font.italic, \
                    '晚表周五留痕格(行%d) 缺估算标记' % room.row
                assert _fill_rgb(cell) == ESTIMATE_FILL_RGB
                marked += 1
        assert written == 19, '晚表周五留痕合计 %d，应为 19' % written
        assert marked > 0

        # 不出率：晚表出勤率区无周五列
        assert '周五' not in detail.rate_cols_by_sheet['雁山（晚）']
        for rec in result.college_rates:
            if rec.sheet == '雁山（晚）':
                assert rec.day != '周五'

        # 不进平均列：晚表平均 == 周一/周三/周四三天的算术平均（不含周五）
        by_key = {(r.sheet, r.college, r.day): r for r in result.college_rates}
        rate_cols = detail.rate_cols_by_sheet['雁山（晚）']
        avg_col = rate_cols.get('平均')
        assert avg_col
        for block in detail.blocks_of('雁山（晚）'):
            rows = block.merge_rows
            if not rows:
                continue
            cell = ws.cell(_rate_row(detail, '雁山（晚）', block), avg_col)
            if cell.value is None:
                continue
            days = [d for d in ('周一', '周三', '周四')
                    if by_key.get(('雁山（晚）', block.name, d)) is not None
                    and by_key[('雁山（晚）', block.name, d)].rate is not None]
            expected = sum(by_key[('雁山（晚）', block.name, d)].rate for d in days) / len(days)
            assert abs(cell.value - expected) < 1e-9, \
                '%s 平均列未按「缺任一天即留空、只算已算出的天」计算' % block.name

        # 不得凭空生成实到：晚表周五实到列在输出里仍为空
        for room in detail.rooms_of('雁山（晚）'):
            actual_col = room.actual_cols.get('周五')
            if actual_col:
                assert ws.cell(room.row, actual_col).value in (None, ''), \
                    '晚表周五实到被凭空生成了（行%d）' % room.row
    finally:
        wb.close()


# ------------------------------------------------------------------ R24 源数据异常


def test_over_actual_cells_are_not_auto_corrected(real_run):
    """R24：实到 > 应到的 8 处**不自动改源数据**。

    早表 r55(5306)=39 / r56(5304)=192 的周五实到虽可判定为两行对调，
    但授权边界是「不凭空生成」，不是「替人改源表」⇒ 输出必须原样保留。
    """
    _result, output, detail_path = real_run
    src = load_workbook(detail_path, data_only=True)
    wb = load_workbook(output, data_only=True)
    try:
        for row, expected in ((55, 39), (56, 192)):
            cell = wb['雁山（早）'].cell(row, 12)      # L 列 = 周五实到
            original = src['雁山（早）'].cell(row, 12).value
            assert cell.value == original, \
                '早表行%d 的周五实到被自动改写了：%r -> %r' % (row, original, cell.value)
            assert cell.value == expected
    finally:
        src.close()
        wb.close()


def test_over_actual_cells_are_surfaced_in_anomalies(real_run):
    """R24：实到 > 应到的格子必须进异常清单（照算出率 + 复核标记），不得静默。"""
    result, _output, _detail = real_run
    notes = ' | '.join('%s %s' % (a.get('type', ''), a.get('note', ''))
                        for a in result.anomalies)
    assert '超过100%' in notes or '超过 100%' in notes or '复核' in notes, \
        '实到 > 应到未进异常清单'
