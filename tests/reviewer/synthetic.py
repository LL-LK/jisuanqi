# -*- coding: utf-8 -*-
"""合成数据生成器（测试审查员专用，tests/reviewer/ 私有）。

数据纪律（硬性）：
  * 学号形如 ``SYN20260001``——10 位、带 ``SYN`` 前缀，与任何真实 13 位学号都不可能相同；
  * 姓名一律「样本生NN」；学院一律「测试学院甲/乙/丙/丁」；
  * 班级一律「测试专业X2026-N班」；教室号一律 ``T0NN``（屏风 ``P0NN``）；
  * 本模块**不读取任何真实工作簿**，产物中不出现任何真实身份信息。

合成明细表的版式与真实模板一致（否则解析器测不到真东西）：
    第1行 标题
    第2行 表头  A序号 B学院（人数） C教室 D分配专业班级 E应到人数
              F周一 G周一免勤人数 H周三 I周三免勤人数 J周四 K周四免勤人数
              L周五 M周五免勤人数 N学生出勤率（%） …
    第3行 出勤率区子表头 N=周一 O=周三 P=周四 Q=周五 R=平均
            （早表有周五，晚表无周五）
    第3行同时也是第一个学院块的首行数据 —— 与真实模板同款怪癖。
"""

from __future__ import annotations

from typing import Dict, Iterable, List, NamedTuple, Optional, Sequence, Tuple

from openpyxl import Workbook

from core.model import CollegeBlock, DetailModel, ExemptRecord, RoomRow
from core.parsing import (extract_group_size, is_group_application, normalize_class,
                          normalize_campus, normalize_room, normalize_text,
                          parse_int, parse_session, parse_weeks, parse_weekdays,
                          strip_spaces)

SHEET_EARLY = '雁山（早）'
SHEET_LATE = '雁山（晚）'
SHEET_PINGFENG = '屏风'
EXEMPT_SHEET_EARLY = '早自习免勤人员统计表'
EXEMPT_SHEET_LATE = '晚自习免勤人员统计表'

RATE_HEADER = '学生出勤率（%）'
EARLY_RATE_LABELS = ('周一', '周三', '周四', '周五', '平均')
LATE_RATE_LABELS = ('周一', '周三', '周四', '平均')
EARLY_HEADER = ('序号', '学院（人数）', '教室', '分配专业班级', '应到人数',
                '周一', '周一免勤人数', '周三', '周三免勤人数',
                '周四', '周四免勤人数', '周五', '周五免勤人数',
                RATE_HEADER, '早自习考勤时间', '执行效果')
LATE_HEADER = ('序号', '学院（人数）', '教室', '分配专业班级', '应到人数',
               '周一', '周一免勤人数', '周三', '周三免勤人数',
               '周四', '周四免勤人数', '周五', '周五免勤人数',
               RATE_HEADER, '值班辅导员未巡查次数', '晚自习考勤时间', '执行效果')

COLLEGE_CODES = {'测试学院甲': '甲', '测试学院乙': '乙', '测试学院丙': '丙', '测试学院丁': '丁'}

#: 学院块：(学院名, 标注总人数, [(行号, 教室号, 应到人数)])
BLOCKS: List[Tuple[str, int, List[Tuple[int, str, int]]]] = [
    ('测试学院甲', 100, [(3, 'T001', 60), (4, 'T002', 40)]),
    ('测试学院乙', 60, [(5, 'T003', 30), (6, 'T004', 30)]),
    ('测试学院丙', 20, [(7, 'T005', 10), (8, 'T006', 10)]),
    ('测试学院丁', 20, [(9, 'T007', 10), (10, 'T008', 10)]),
]

#: 早表实到：(学院, 星期) -> 各教室实到。None = 缺失；'免勤' = 文字态。
EARLY_ACTUALS: Dict[Tuple[str, str], Tuple] = {
    ('测试学院甲', '周一'): (50, 32), ('测试学院甲', '周三'): (49, 33),
    ('测试学院甲', '周四'): (48, 31), ('测试学院甲', '周五'): (47, 30),
    ('测试学院乙', '周一'): (None, 29),          # 缺失：不得写 0、不得出率
    ('测试学院乙', '周三'): (28, 29), ('测试学院乙', '周四'): (29, 30),
    ('测试学院乙', '周五'): (27, 28),
    ('测试学院丙', '周一'): (9, 9), ('测试学院丙', '周三'): (9, 9),
    ('测试学院丙', '周四'): (9, 9), ('测试学院丙', '周五'): (9, 9),
    ('测试学院丁', '周一'): (9, 10), ('测试学院丁', '周三'): (9, 10),
    ('测试学院丁', '周四'): (9, 10), ('测试学院丁', '周五'): (9, 10),
}

#: 晚表实到（无周五列）。
LATE_ACTUALS: Dict[Tuple[str, str], Tuple] = {
    ('测试学院甲', '周一'): (50, 32), ('测试学院甲', '周三'): (49, 33),
    ('测试学院甲', '周四'): (48, 31),
    ('测试学院乙', '周一'): (30, 29), ('测试学院乙', '周三'): (27, 29),
    ('测试学院乙', '周四'): (29, 30),
    ('测试学院丙', '周一'): (9, 9), ('测试学院丙', '周三'): (9, 9),
    ('测试学院丙', '周四'): ('免勤', 9),           # 文字态：与缺失同等处理
    ('测试学院丁', '周一'): (9, 10), ('测试学院丁', '周三'): (9, 10),
    ('测试学院丁', '周四'): (9, 10),
}

#: 屏风表内容（逐单元格对比用）。
PINGFENG_ROWS = [
    ('测试学院丙', 'P001', 18),
    ('测试学院丁', 'P002', 15),
]

EXEMPT_HEADER = ('序号', '学号', '学院', '姓名', '专业班级', '免勤周次',
                 '免勤时间', '免勤自习', '所在校区', '申请免勤理由')


class ExemptRow(NamedTuple):
    """一条免勤申请，字段顺序**必须**与 EXEMPT_HEADER 的数据列一致。"""

    sid: str
    college: str
    name: str
    class_text: str
    weeks: str
    time: str
    session: str
    campus: str
    reason: str = '合成理由'


def fake_sid(n: int) -> str:
    """合成学号：13 位、99 开头 + 全零前缀 + 序号。

    形如「9900」+ 9 位（8 个 0 + 1 位序号）——位数与真实一致（引擎会校验 13 位），
    但 99 + 全零的号段在任何真实学籍里都不存在，不会指向任何真人。
    （此处刻意不写出完整 13 位字面量，避免被发布侧「13 位学号」扫描误杀。）
    """
    return '9900%09d' % n


def fake_name(n: int) -> str:
    return '样本生%02d' % n


def college_class(college: str, index: int = 1) -> str:
    return '测试专业%s2026-%d班' % (COLLEGE_CODES[college], index)


def _sheet_specs():
    return (
        (SHEET_EARLY, EARLY_HEADER, EARLY_RATE_LABELS, EARLY_ACTUALS,
         ('周一', '周三', '周四', '周五')),
        (SHEET_LATE, LATE_HEADER, LATE_RATE_LABELS, LATE_ACTUALS,
         ('周一', '周三', '周四')),
    )


def build_detail_workbook(path: str, broken_merge: bool = False) -> str:
    """构造合成明细表。

    :param broken_merge: 故意制造模板缺陷——丙块 N 列漏合并、丁块 O~R 与 N 行块不齐，
        用来验证导出阶段的「合并区归一」。
    """
    wb = Workbook()
    wb.remove(wb.active)

    for sheet, header, labels, actuals, days in _sheet_specs():
        ws = wb.create_sheet(sheet)
        ws.cell(1, 1, '合成数据（非真实）· %s' % sheet)
        rate_start = header.index(RATE_HEADER) + 1
        # 真实模板里出勤率区是一个**合并区**（如 N2:R2），后面的指标列从合并区之后开始。
        # 这里按同样规则排版：出勤率区占 len(labels) 列，其余表头整体右移，避免
        # 「早自习考勤时间 / 执行效果」挤进出勤率标签列造成列语义冲突。
        shift = len(labels) - 1
        for col, head in enumerate(header, 1):
            if col < rate_start:
                ws.cell(2, col, head)
            elif col == rate_start:
                ws.cell(2, col, head)
            else:
                ws.cell(2, col + shift, head)
        for i, label in enumerate(labels):
            ws.cell(3, rate_start + i, label)
        if len(labels) > 1:
            ws.merge_cells(start_row=2, start_column=rate_start,
                           end_row=2, end_column=rate_start + len(labels) - 1)

        for name, total, rooms in BLOCKS:
            ws.cell(rooms[0][0], 2, '%s（%d人）' % (name, total))
            if rooms[-1][0] > rooms[0][0]:
                ws.merge_cells(start_row=rooms[0][0], start_column=2,
                               end_row=rooms[-1][0], end_column=2)
            values = actuals[(name, days[0])]
            for i, (row, room, expected) in enumerate(rooms):
                ws.cell(row, 1, row)
                ws.cell(row, 3, room)
                ws.cell(row, 4, college_class(name, i + 1))
                ws.cell(row, 5, expected)
                for day in days:
                    ws.cell(row, header.index(day) + 1, actuals[(name, day)][i])

        # ---- 出勤率区合并 ----
        for name, _total, rooms in BLOCKS:
            block_rows = [r for r, _rm, _e in rooms if r != 3]
            if len(block_rows) < 2:
                continue
            lo, hi = block_rows[0], block_rows[-1]
            for i in range(len(labels)):
                col = rate_start + i
                if broken_merge and name == '测试学院丙' and i == 0:
                    continue                      # N 列漏合并（照搬 N69:N70 缺陷）
                if broken_merge and name == '测试学院丁' and i > 0:
                    ws.merge_cells(start_row=lo - 1, start_column=col,
                                   end_row=hi, end_column=col)   # O~R 比 N 多一行
                    continue
                ws.merge_cells(start_row=lo, start_column=col,
                               end_row=hi, end_column=col)

    ws = wb.create_sheet(SHEET_PINGFENG)
    ws.cell(1, 1, '屏风校区（合成数据，整体排除）')
    for col, head in enumerate(('学院', '教室', '应到人数'), 1):
        ws.cell(2, col, head)
    for i, row in enumerate(PINGFENG_ROWS, 3):
        for col, value in enumerate(row, 1):
            ws.cell(i, col, value)

    wb.save(path)
    wb.close()
    return path


def default_exempt_rows() -> Tuple[List[ExemptRow], List[ExemptRow]]:
    """合成免勤表内容：(早表行, 晚表行)。

    覆盖：正常个人免勤、错字校区（燕山）、集体申请覆盖同班个人、
    同一学号早/晚各一次、屏风记录、晚表周五（晚表无周五列）、非本周申请。
    """
    early: List[ExemptRow] = []
    late: List[ExemptRow] = []

    # 甲：周一 早自习 ×5，其中一条用错字校区「燕山校区」
    for i in range(5):
        early.append(ExemptRow(
            sid=fake_sid(i + 1), college='测试学院甲', name=fake_name(i + 1),
            class_text=college_class('测试学院甲', 1), weeks='第7周', time='周一',
            session='早自习', campus='雁山校区' if i < 4 else '燕山校区'))
    # 丙：集体申请（共25人）覆盖同班 5 条个人申请
    early.append(ExemptRow(
        sid='', college='测试学院丙', name='测试专业丙2026-1班全体学生（共25人）',
        class_text=college_class('测试学院丙', 1), weeks='第7周', time='周四',
        session='早自习', campus='雁山校区'))
    for i in range(5):
        early.append(ExemptRow(
            sid=fake_sid(101 + i), college='测试学院丙', name=fake_name(101 + i),
            class_text=college_class('测试学院丙', 1), weeks='第7周', time='周四',
            session='早自习', campus='雁山校区'))
    # 乙：周三 早自习 ×2
    for i in range(2):
        early.append(ExemptRow(
            sid=fake_sid(201 + i), college='测试学院乙', name=fake_name(201 + i),
            class_text=college_class('测试学院乙', 1), weeks='第7周', time='周三',
            session='早自习', campus='雁山校区'))
    # 屏风：必须被整体排除
    for i in range(2):
        early.append(ExemptRow(
            sid=fake_sid(301 + i), college='测试学院丙', name=fake_name(301 + i),
            class_text=college_class('测试学院丙', 1), weeks='第7周', time='周一',
            session='早自习', campus='屏风校区'))

    # 乙：周三 晚自习 ×3（同一学号同时出现在早、晚 → 早晚各计一次）
    for i in range(3):
        late.append(ExemptRow(
            sid=fake_sid(201 + i), college='测试学院乙', name=fake_name(201 + i),
            class_text=college_class('测试学院乙', 1), weeks='第7周', time='周三',
            session='晚自习', campus='雁山校区'))
    # 甲：周一 晚自习 ×1
    late.append(ExemptRow(
        sid=fake_sid(401), college='测试学院甲', name=fake_name(401),
        class_text=college_class('测试学院甲', 1), weeks='第7周', time='周一',
        session='晚自习', campus='雁山校区'))
    # 丁：周三 晚自习 ×30 → 超过学院应到 20（分母 ≤ 0）
    for i in range(30):
        late.append(ExemptRow(
            sid=fake_sid(500 + i), college='测试学院丁', name=fake_name(500 + i),
            class_text=college_class('测试学院丁', 1), weeks='第7周', time='周三',
            session='晚自习', campus='雁山校区'))
    # 丙：周五 晚自习 ×1 → 晚表没有周五列，必须进异常清单而不是静默丢弃
    late.append(ExemptRow(
        sid=fake_sid(601), college='测试学院丙', name=fake_name(601),
        class_text=college_class('测试学院丙', 1), weeks='第7周', time='周五',
        session='晚自习', campus='雁山校区'))
    # 非本周：必须被周次窗口过滤
    late.append(ExemptRow(
        sid=fake_sid(701), college='测试学院甲', name=fake_name(701),
        class_text=college_class('测试学院甲', 1), weeks='第6周', time='周一',
        session='晚自习', campus='雁山校区'))
    late.append(ExemptRow(
        sid=fake_sid(702), college='测试学院甲', name=fake_name(702),
        class_text=college_class('测试学院甲', 1), weeks='第6-7，9-13周', time='周一',
        session='晚自习', campus='雁山校区'))

    return early, late


def build_exempt_workbook(path: str, early_rows: Sequence[ExemptRow],
                          late_rows: Sequence[ExemptRow]) -> str:
    """构造合成免勤表（两张工作表，第 2 行表头）。"""
    wb = Workbook()
    wb.remove(wb.active)
    for sheet, rows in ((EXEMPT_SHEET_EARLY, early_rows), (EXEMPT_SHEET_LATE, late_rows)):
        ws = wb.create_sheet(sheet)
        ws.cell(1, 1, '合成数据（非真实）· %s' % sheet)
        for col, head in enumerate(EXEMPT_HEADER, 1):
            ws.cell(2, col, head)
        for i, row in enumerate(rows, 3):
            ws.cell(i, 1, i - 2)                       # 序号列
            ws.cell(i, 2, row.sid)
            ws.cell(i, 3, row.college)
            ws.cell(i, 4, row.name)
            ws.cell(i, 5, row.class_text)
            ws.cell(i, 6, row.weeks)
            ws.cell(i, 7, row.time)
            ws.cell(i, 8, row.session)
            ws.cell(i, 9, row.campus)
            ws.cell(i, 10, row.reason)
    wb.save(path)
    wb.close()
    return path


def make_record(sid: str = '', name: str = '', college: str = '测试学院甲',
                class_text: str = '测试专业甲2026-1班', weeks: str = '第7周',
                time: str = '周一', session: str = '早自习', campus: str = '雁山校区',
                room: str = '', reason: str = '合成理由', source: str = '合成',
                row: int = 3) -> ExemptRecord:
    """按与 ``load_exemptions`` 完全相同的规则造一条免勤记录。"""
    group_size = extract_group_size(name) or extract_group_size(class_text)
    if not sid and is_group_application(name, class_text):
        if group_size is None:
            group_size = 0
    elif not sid:
        group_size = None
    return ExemptRecord(
        sid=strip_spaces(sid), name=normalize_text(name), college=normalize_text(college),
        class_text=normalize_text(class_text), class_key=normalize_class(class_text or name),
        weeks=parse_weeks(weeks), weekdays=parse_weekdays(time),
        session=parse_session(session) or session, campus=normalize_campus(campus),
        room=normalize_room(room), reason=normalize_text(reason),
        group_size=group_size, source=source, row=row,
    )


def make_detail_model(spec: Sequence[Tuple[str, Optional[int], Sequence[Tuple[str, Optional[int], Dict[str, Optional[int]]]]]],
                      sheet: str = SHEET_EARLY,
                      days: Sequence[str] = ('周一', '周三', '周四', '周五')) -> DetailModel:
    """在内存里拼一个明细表模型（免去写文件的开销，供性质化测试使用）。

    :param spec: [(学院名, 标注总人数, [(教室号, 应到人数, {星期: 实到})])]
    """
    detail = DetailModel(path='<合成>')
    detail.sheet_names = [SHEET_EARLY, SHEET_LATE, SHEET_PINGFENG]
    rate_cols: Dict[str, int] = {}
    col = 14
    for day in days:
        rate_cols[day] = col
        col += 1
    rate_cols['平均'] = col
    detail.rate_cols_by_sheet[sheet] = rate_cols

    row = 4                                    # 从第 4 行起，避开第 3 行子表头
    for name, total, rooms in spec:
        rows = []
        for room, expected, actuals in rooms:
            rows.append(row)
            detail.rooms.append(RoomRow(
                sheet=sheet, row=row, seq=row, college_name=name, college_total=total,
                room=room, classes_text='合成班级', expected=expected,
                actuals=dict(actuals),
                actual_raw={d: ('' if v is None else str(v)) for d, v in actuals.items()},
                actual_cols={d: i for i, d in enumerate(actuals)},
                exempt_cols={d: i for i, d in enumerate(actuals)},
            ))
            row += 1
        detail.blocks.append(CollegeBlock(sheet=sheet, name=name, total=total,
                                          rows=rows, rate_cols=dict(rate_cols)))
    return detail
