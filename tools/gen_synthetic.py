# -*- coding: utf-8 -*-
"""合成数据生成器：为「三早一晚免勤自动统计系统」生成**纯虚构**的测试样本。

生成两个 Excel 文件（默认写到 ``<项目根>/_synthetic/``）：

    detail.xlsx   执行情况明细表
                  工作表：雁山（早） / 雁山（晚） / 屏风
    exempt.xlsx   免勤人员统计表
                  工作表：早自习免勤人员统计表 / 晚自习免勤人员统计表

数据全部为虚构占位，**不读取也不引用工作区里的任何真实 Excel**：

  * 学院名使用通用学科名（理学院 / 工学院 / 商学院 / 文学院 / 计算机学院 …）；
  * 学号统一为 ``99`` 开头的 13 位假号（9900000000001 …）；
  * 姓名为「学生甲01」「学生乙02」式占位；
  * 教室号使用 81xx / 82xx 明显虚构的号段。

明细表刻意复现真实模板的两个自带缺陷（否则导出端没有东西可归一化）：

  * ``文学院`` 块位于第 57..59 行：N 列只合并到 58 行，O/P/Q 列合并到 59 行
    —— 同一学院各列合并行数不一致；
  * ``医学院`` 块位于第 69..70 行：N 列 **漏合并**（N69:N70 不合并），
    其余列正常合并。

用法::

    python -X utf8 tools/gen_synthetic.py --out _synthetic --scale small --dirty
    python -X utf8 tools/gen_synthetic.py --out _synthetic --scale large --seed 1
"""

from __future__ import annotations

import argparse
import math
import random
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Set, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_OUT_DIR = ROOT / '_synthetic'

# ------------------------------------------------------------------ 工作表名

SHEET_EARLY = '雁山（早）'
SHEET_LATE = '雁山（晚）'
SHEET_PINGFENG = '屏风'
DETAIL_SHEETS = (SHEET_EARLY, SHEET_LATE, SHEET_PINGFENG)

EXEMPT_SHEET_EARLY = '早自习免勤人员统计表'
EXEMPT_SHEET_LATE = '晚自习免勤人员统计表'

# ------------------------------------------------------------------ 列布局

COL_SEQ = 1            # A 序号
COL_COLLEGE = 2        # B 学院（人数）
COL_ROOM = 3           # C 教室
COL_CLASSES = 4        # D 分配专业班级
COL_EXPECTED = 5       # E 应到人数
DAYS = ('周一', '周三', '周四', '周五')
ACTUAL_COL = {'周一': 6, '周三': 8, '周四': 10, '周五': 12}     # F H J L
EXEMPT_COL = {'周一': 7, '周三': 9, '周四': 11, '周五': 13}     # G I K M
RATE_START_COL = 14    # N 学生出勤率（%）
COL_UNCHECKED = 19     # S 值班辅导员未巡查次数
COL_ATTEND_TIME = 22   # V 早/晚自习考勤时间
COL_EFFECT = 26        # Z 执行效果

HEADER_ROW = 2
RATE_SUBHEADER_ROW = 3
FIRST_DATA_ROW = 4

DETAIL_HEADERS = {
    COL_SEQ: '序号',
    COL_COLLEGE: '学院（人数）',
    COL_ROOM: '教室',
    COL_CLASSES: '分配专业班级',
    COL_EXPECTED: '应到人数',
    6: '周一', 7: '周一免勤人数',
    8: '周三', 9: '周三免勤人数',
    10: '周四', 11: '周四免勤人数',
    12: '周五', 13: '周五免勤人数',
    RATE_START_COL: '学生出勤率（%）',
    COL_UNCHECKED: '值班辅导员未巡查次数',
    COL_ATTEND_TIME: '早自习考勤时间',
    COL_EFFECT: '执行效果',
}

# 早自习表出勤率区：N 周一 O 周三 P 周四 Q 周五 R 平均
RATE_LABELS_EARLY = (('周一', 14), ('周三', 15), ('周四', 16), ('周五', 17), ('平均', 18))
# 晚自习表没有周五列：N 周一 O 周三 P 周四 Q 平均
RATE_LABELS_LATE = (('周一', 14), ('周三', 15), ('周四', 16), ('平均', 17))

EXEMPT_HEADERS = ('序号', '学号', '学院', '姓名', '专业班级', '免勤周次',
                  '免勤时间', '免勤自习', '所在校区', '申请免勤理由')
ROOM_HEADER = '教室号'

# ------------------------------------------------------------------ 学院 / 专业

COLLEGE_MAJORS: Dict[str, Tuple[str, ...]] = {
    '理学院': ('数学与应用数学', '应用统计学'),
    '工学院': ('机械工程', '土木工程'),
    '商学院': ('工商管理', '会计学'),
    '文学院': ('汉语言文学', '新闻学'),
    '计算机学院': ('计算机科学与技术', '软件工程'),
    '医学院': ('临床医学', '护理学'),
    '环境学院': ('环境科学', '地理科学'),
    '土木学院': ('建筑环境与能源应用工程', '给排水科学与工程'),
    '艺术学院': ('视觉传达设计', '音乐学'),
    '体育学院': ('体育教育', '社会体育指导与管理'),
    # 屏风校区专用：只出现在屏风工作表，永远不出现在免勤表里
    '化工学院': ('化学工程与工艺', '制药工程'),
    '交通学院': ('交通工程', '交通运输'),
}

# 前 6 个学院块的位置固定，保证两处模板缺陷落在真实行号上：
#   文学院 → 行 57..59（各列合并行数不一致）
#   医学院 → 行 69..70（N69:N70 漏合并）
FIXED_BLOCKS: Tuple[Tuple[str, int], ...] = (
    ('理学院', 20),        # 行 4..23
    ('工学院', 18),        # 行 24..41
    ('商学院', 15),        # 行 42..56
    ('文学院', 3),         # 行 57..59   ← 缺陷块一
    ('计算机学院', 9),     # 行 60..68
    ('医学院', 2),         # 行 69..70   ← 缺陷块二
)
TRAILING_COLLEGES = ('环境学院', '土木学院', '艺术学院', '体育学院')
PINGFENG_BLOCKS: Tuple[Tuple[str, int], ...] = (
    ('化工学院', 6),
    ('交通学院', 5),
)

# ------------------------------------------------------------------ 脏格式

WEEK_CLEAN = '第7周'
WEEK_DIRTY = ('第7周', '第7-18周', '第6-7，9-13周', '第7—12周',
              '第七周，第八周', '第7、8、9周', '6-10周', '第7周、第8周')
WEEK_OTHER = ('第6周', '第8周', '第5周', '第9周')

WEEKDAY_CLEAN = ('周一', '周三', '周四', '周五')
WEEKDAY_DIRTY = ('周一、周三、周四、周五', '周一 ，周三 ，周四', '周一到周五',
                 '周一周二周三周四周五', '周一、二、四、五')

CAMPUS_CLEAN = ('雁山校区',)
CAMPUS_DIRTY = ('雁山校区', '雁山校区', '雁山校区', '雁山校园', '燕山校区',
                 '屏风校区', '雁门校区')

REASONS = ('身体不适', '外出参加竞赛', '课程冲突', '家中有事', '参加校级活动')

NAME_PREFIX = ('甲', '乙', '丙', '丁', '戊', '己')

# ------------------------------------------------------------------ 规模

SCALES: Dict[str, Dict[str, int]] = {
    'small': {'rooms': 72, 'exempt_per_session': 150},
    'medium': {'rooms': 150, 'exempt_per_session': 1500},
    'large': {'rooms': 150, 'exempt_per_session': 30000},
}
# rooms = **每张**明细表的教室行数。固定块已经占掉 67 行（行 4..70），
# 缺陷锚点 57..59 / 69..70 必须存在，所以下限就是 67 + 5 = 72 行。


# ================================================================== 数据结构


@dataclass
class DetailRow:
    """明细表的一行（一间教室）。"""
    row: int
    seq: int
    college: str
    college_total: int
    room: str
    classes_text: str
    expected: int
    actuals: Dict[str, object] = field(default_factory=dict)


@dataclass
class BlockSpec:
    """一个学院块（连续若干行）。"""
    start_row: int
    end_row: int
    college: str
    college_total: int = 0
    skip_cols: Set[int] = field(default_factory=set)          # 完全不合并的列
    partial_merges: Dict[int, Tuple[int, int]] = field(default_factory=dict)
    dirty_actuals: bool = False


@dataclass
class SheetSpec:
    """一张明细工作表的完整描述。"""
    name: str
    title: str
    rows: List[DetailRow] = field(default_factory=list)
    blocks: List[BlockSpec] = field(default_factory=list)
    rate_labels: Tuple[Tuple[str, int], ...] = RATE_LABELS_EARLY
    attendance_time: str = '07:00-07:30'


@dataclass
class CollegePlan:
    """学院维度的规划信息（房间、班级池、应到总数）。"""
    name: str
    total: int = 0
    classes: List[str] = field(default_factory=list)
    rooms: List[str] = field(default_factory=list)
    class_rooms: Dict[str, str] = field(default_factory=dict)


@dataclass
class RoomPlan:
    """一间教室的静态信息（早/晚两张表共用同一批教室）。"""
    room: str
    classes_text: str
    expected: int


@dataclass
class ExemptRow:
    """免勤表的一行。"""
    seq: int
    sid: str
    college: str
    name: str
    class_text: str
    week_text: str
    day_text: str
    session: str
    campus: str
    reason: str
    room: str = ''


# ================================================================== 布局规划


def split_evenly(total: int, parts: int) -> List[int]:
    """把 total 尽量平均地拆成 parts 份（和恒等于 total）。"""
    if parts <= 0 or total <= 0:
        return []
    base, extra = divmod(total, parts)
    return [base + (1 if i < extra else 0) for i in range(parts)]


def class_pool(college: str) -> List[str]:
    """学院的专业班级池（虚构班号）。"""
    pool: List[str] = []
    for major in COLLEGE_MAJORS.get(college, (college + '专业',)):
        pool.append('%s26-1班' % major)
        pool.append('%s26-2班' % major)
    return pool


def room_number(index: int) -> str:
    """明显虚构的教室号：8101/8102/8103/8104/8201/…（每 4 间换一次楼栋号）。"""
    return '%d%02d' % (81 + index // 4, index % 4 + 1)


def plan_blocks(total_rooms: int) -> List[BlockSpec]:
    """规划一张明细表的学院块布局。

    前 6 个块固定（合计 67 行，行 4..70），后面按剩余行数切分到
    ``TRAILING_COLLEGES``。第一个尾部块负责注入「实到缺失/文字/0」三种脏值。
    """
    blocks: List[BlockSpec] = []
    row = FIRST_DATA_ROW
    for name, count in FIXED_BLOCKS:
        blocks.append(BlockSpec(row, row + count - 1, name))
        row += count
    used = row - FIRST_DATA_ROW
    if total_rooms < used:
        raise ValueError('教室行数不足：至少需要 %d 行才能放下缺陷锚点' % used)

    remaining = total_rooms - used
    if remaining:
        parts = max(1, min(len(TRAILING_COLLEGES), int(math.ceil(remaining / 25.0))))
        sizes = split_evenly(remaining, parts)
        for idx, size in enumerate(sizes):
            blocks.append(BlockSpec(row, row + size - 1, TRAILING_COLLEGES[idx],
                                    dirty_actuals=(idx == 0)))
            row += size
    return blocks


def apply_template_defects(blocks: List[BlockSpec]) -> None:
    """在固定位置上复现真实模板的两个自带缺陷。"""
    for blk in blocks:
        if blk.college == '文学院' and blk.start_row == 57:
            # N 列只合并到 58 行，O/P/Q 列合并到 59 行 → 合并行数不一致
            blk.partial_merges[RATE_START_COL] = (57, 58)
        elif blk.college == '医学院' and blk.start_row == 69:
            # N69:N70 漏合并
            blk.skip_cols.add(RATE_START_COL)


# ================================================================== 明细表生成


def _actual_style(rng: random.Random, style: str, expected: int) -> Dict[str, object]:
    """按风格生成一周的实到人数。"""
    out: Dict[str, object] = {}

    def deficit() -> int:
        # 缺勤幅度取 8~20 人：保证学院合计实到明显低于「应到 − 免勤」，
        # 干净样本里不会出现 >100% 的出勤率。
        return min(rng.randint(8, 20), expected)

    if style == 'missing':
        out['周一'] = None
        out['周三'] = max(expected - deficit(), 0)
        out['周四'] = max(expected - deficit(), 0)
        out['周五'] = max(expected - deficit(), 0)
    elif style == 'text':
        out['周一'] = '免勤'
        out['周三'] = max(expected - deficit(), 0)
        out['周四'] = max(expected - deficit(), 0)
        out['周五'] = max(expected - deficit(), 0)
    elif style == 'zero':
        out['周一'] = 0
        out['周三'] = max(expected - deficit(), 0)
        out['周四'] = max(expected - deficit(), 0)
        out['周五'] = max(expected - deficit(), 0)
    elif style == 'mixed':
        out['周一'] = None
        out['周三'] = '免勤'
        out['周四'] = 0
        out['周五'] = max(expected - deficit(), 0)
    else:   # full
        for day in DAYS:
            out[day] = max(expected - deficit(), 0)
    return out


def plan_rooms(blocks: Sequence[BlockSpec], rng: random.Random) -> Dict[str, List[RoomPlan]]:
    """规划各学院的教室骨架（早/晚两张表共用同一批教室）。"""
    rooms: Dict[str, List[RoomPlan]] = {}
    room_index = 0
    for blk in blocks:
        pool = class_pool(blk.college)
        items: List[RoomPlan] = []
        for offset in range(blk.start_row, blk.end_row + 1):
            items.append(RoomPlan(
                room=room_number(room_index),
                classes_text=pool[(offset - blk.start_row) % len(pool)],
                expected=rng.randint(30, 55),
            ))
            room_index += 1
        rooms[blk.college] = items
    return rooms


def build_sheet(name: str, blocks: List[BlockSpec], rng: random.Random,
                plans: Dict[str, CollegePlan],
                skeleton: Dict[str, List[RoomPlan]]) -> SheetSpec:
    """生成一张明细工作表的内容。"""
    spec = SheetSpec(name=name, title='合成数据·%s执行情况明细表' % name)
    spec.rate_labels = RATE_LABELS_EARLY if name == SHEET_EARLY else RATE_LABELS_LATE
    spec.attendance_time = '07:00-07:30' if name == SHEET_EARLY else '19:00-20:30'

    seq = 1
    for blk in blocks:
        pool = class_pool(blk.college)
        plan = plans.setdefault(blk.college, CollegePlan(name=blk.college))
        plan.classes = pool
        items = skeleton.get(blk.college, [])
        rows: List[DetailRow] = []
        for offset in range(blk.start_row, blk.end_row + 1):
            item = items[(offset - blk.start_row) % len(items)] if items else RoomPlan(
                room=room_number(offset), classes_text=pool[0], expected=30)
            if blk.dirty_actuals:
                style = ('missing', 'text', 'zero', 'mixed')[(offset - blk.start_row) % 4]
            else:
                style = 'full'
            rows.append(DetailRow(
                row=offset, seq=seq, college=blk.college, college_total=0,
                room=item.room, classes_text=item.classes_text,
                expected=item.expected,
                actuals=_actual_style(rng, style, item.expected)))
            seq += 1
            plan.rooms.append(item.room)
            plan.class_rooms.setdefault(item.classes_text, item.room)
        blk.college_total = sum(r.expected for r in rows)
        for r in rows:
            r.college_total = blk.college_total
        plan.total = blk.college_total
        spec.rows.extend(rows)
        spec.blocks.append(blk)
    return spec


def build_pingfeng_sheet(rng: random.Random) -> SheetSpec:
    """屏风工作表：有真实数据行，但整体不参与任何计算。"""
    spec = SheetSpec(name=SHEET_PINGFENG, title='合成数据·屏风校区执行情况明细表')
    spec.rate_labels = RATE_LABELS_EARLY
    spec.attendance_time = '07:00-07:30'
    row = FIRST_DATA_ROW
    seq = 1
    room_index = 0
    for college, count in PINGFENG_BLOCKS:
        pool = class_pool(college)
        blk = BlockSpec(row, row + count - 1, college)
        rows = []
        for offset in range(row, row + count):
            room = room_number(room_index)
            room_index += 1
            cls = pool[(offset - row) % len(pool)]
            expected = rng.randint(25, 45)
            rows.append(DetailRow(
                row=offset, seq=seq, college=college, college_total=0, room=room,
                classes_text=cls, expected=expected,
                actuals=_actual_style(rng, 'full', expected)))
            seq += 1
        blk.college_total = sum(r.expected for r in rows)
        for r in rows:
            r.college_total = blk.college_total
        spec.rows.extend(rows)
        spec.blocks.append(blk)
        row += count
    return spec


# ================================================================== 免勤表生成


def _week_text(rng: random.Random, dirty: bool) -> str:
    if not dirty:
        return WEEK_CLEAN
    roll = rng.random()
    if roll < 0.06:
        return rng.choice(WEEK_OTHER)          # 不含第 7 周 → 本次统计不计入
    return rng.choice(WEEK_DIRTY)


def _day_text(rng: random.Random, dirty: bool) -> str:
    if not dirty:
        return rng.choice(WEEKDAY_CLEAN)
    if rng.random() < 0.75:
        return rng.choice(WEEKDAY_DIRTY)
    return rng.choice(WEEKDAY_CLEAN)


def _campus_text(rng: random.Random, dirty: bool) -> str:
    if not dirty:
        return '雁山校区'
    return rng.choice(CAMPUS_DIRTY)


def build_exempt_rows(plans: Sequence[CollegePlan], count: int, session: str,
                      rng: random.Random, dirty: bool,
                      sid_start: int = 1) -> List[ExemptRow]:
    """生成某一次自习的免勤申请行。

    学院按**教室数加权**抽取：学院越大免勤人数越多，避免小学院被灌入
    远超其规模的免勤人数（那会让干净样本出现 >100% 的出勤率）。
    """
    rows: List[ExemptRow] = []
    if not plans:
        return rows
    weighted: List[int] = []
    for idx, plan in enumerate(plans):
        weighted.extend([idx] * max(len(plan.rooms), 1))
    for i in range(count):
        plan = plans[rng.choice(weighted)]
        cls = plan.classes[rng.randrange(len(plan.classes))] if plan.classes else ''
        sid = '99%011d' % (sid_start + i)
        name = '学生%s%02d' % (NAME_PREFIX[i % len(NAME_PREFIX)], (i % 99) + 1)
        rows.append(ExemptRow(
            seq=i + 1,
            sid=sid,
            college=plan.name,
            name=name,
            class_text=cls,
            week_text=_week_text(rng, dirty),
            day_text=_day_text(rng, dirty),
            session='早' if session == EXEMPT_SHEET_EARLY else '晚',
            campus=_campus_text(rng, dirty),
            reason=rng.choice(REASONS),
            room=plan.class_rooms.get(cls, ''),
        ))
    return rows


def inject_dirty_exempt_rows(rows: List[ExemptRow], plans: Sequence[CollegePlan],
                             rng: random.Random, session: str) -> List[ExemptRow]:
    """注入第 3 类脏数据：重复、姓名/班级写法不一致、集体申请、越界、无法解析。"""
    extra: List[ExemptRow] = []
    by_college = {p.name: p for p in plans}
    base = list(rows)

    # ---- 1. 重复申请：同一学号、同一(星期, 自习) 多条 ----
    for row in base:
        if not row.sid or row.campus != '雁山校区':
            continue
        if rng.random() > 0.05:
            continue
        dup = ExemptRow(
            seq=row.seq,
            sid=row.sid,
            college=row.college,
            name=row.name.replace('学生', '学生 ') if rng.random() < 0.5 else row.name,
            class_text=row.class_text,
            week_text=rng.choice(WEEK_DIRTY),
            day_text=row.day_text,
            session=row.session,
            campus=rng.choice(('雁山校区', '雁山校园')),
            reason=rng.choice(REASONS),
            room=row.room,
        )
        # 一部分重复行还把班级写成另一种写法（2026 级 → 26 级）
        if rng.random() < 0.4 and row.class_text:
            dup.class_text = row.class_text.replace('26-', '2026-')
        extra.append(dup)

    # ---- 2. 集体申请 + 同班个人申请（个人应被集体覆盖） ----
    for college in ('理学院', '工学院'):
        plan = by_college.get(college)
        if not plan or not plan.classes:
            continue
        cls = plan.classes[0]
        day = rng.choice(WEEKDAY_CLEAN)
        group_size = rng.randint(8, 24)
        extra.append(ExemptRow(
            seq=0, sid='', college=college,
            name='%s全体学生（共%d人）' % (cls, group_size),
            class_text=cls, week_text=rng.choice(WEEK_DIRTY), day_text=day,
            session='早' if session == EXEMPT_SHEET_EARLY else '晚',
            campus='雁山校区', reason='集体外出参加实践活动',
            room=plan.class_rooms.get(cls, '')))
        for k in range(3):
            extra.append(ExemptRow(
                seq=0, sid='99%011d' % (9000000 + k + rng.randrange(1000)),
                college=college, name='学生甲%02d' % (k + 1), class_text=cls,
                week_text=rng.choice(WEEK_DIRTY), day_text=day,
                session='早' if session == EXEMPT_SHEET_EARLY else '晚',
                campus='雁山校区', reason=rng.choice(REASONS),
                room=plan.class_rooms.get(cls, '')))

    # ---- 3. 越界：集体声明人数超过学院应到总人数（只在早自习表注入一次） ----
    if session == EXEMPT_SHEET_EARLY:
        plan = by_college.get('文学院')
        if plan and plan.classes:
            cls = plan.classes[0]
            extra.append(ExemptRow(
                seq=0, sid='', college='文学院',
                name='%s全体学生（共%d人）' % (cls, plan.total + 40),
                class_text=cls, week_text='第7周', day_text='周三',
                session='早', campus='雁山校区', reason='集体外出参加实践活动',
                room=plan.class_rooms.get(cls, '')))

    # ---- 4. 无法解析：一条周次为空、一条星期为空 ----
    if session == EXEMPT_SHEET_EARLY:
        extra.append(ExemptRow(
            seq=0, sid='9900009999999', college='商学院', name='学生戊99',
            class_text=class_pool('商学院')[0], week_text='', day_text='周一',
            session='早', campus='雁山校区', reason='手写记录，周次栏空白',
            room=''))
        extra.append(ExemptRow(
            seq=0, sid='9900009999998', college='商学院', name='学生丁98',
            class_text=class_pool('商学院')[1], week_text='第7周', day_text='',
            session='早', campus='雁山校区', reason='手写记录，星期栏空白',
            room=''))
    return extra


# ================================================================== 写 Excel


def write_detail_workbook(path: Path, sheets: Sequence[SheetSpec]) -> Path:
    """写明细表：表头、子表头、学院块合并、模板缺陷合并。"""
    wb = Workbook()
    wb.remove(wb.active)
    for spec in sheets:
        ws = wb.create_sheet(spec.name)
        ws.cell(1, 1, spec.title)
        ws.merge_cells(start_row=1, start_column=1,
                       end_row=1, end_column=max(RATE_START_COL + 4, COL_EFFECT))
        ws.cell(1, 1).font = Font(bold=True, size=13)
        ws.cell(1, 1).alignment = Alignment(horizontal='center')

        headers = dict(DETAIL_HEADERS)
        headers[COL_ATTEND_TIME] = ('早' if spec.name == SHEET_EARLY else '晚') + '自习考勤时间'
        for col, text in headers.items():
            cell = ws.cell(HEADER_ROW, col, text)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center', vertical='center')

        for label, col in spec.rate_labels:
            cell = ws.cell(RATE_SUBHEADER_ROW, col, label)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        for row in spec.rows:
            ws.cell(row.row, COL_SEQ, row.seq)
            ws.cell(row.row, COL_ROOM, row.room)
            ws.cell(row.row, COL_CLASSES, row.classes_text)
            ws.cell(row.row, COL_EXPECTED, row.expected)
            for day in DAYS:
                if day in row.actuals:
                    ws.cell(row.row, ACTUAL_COL[day], row.actuals[day])
                ws.cell(row.row, EXEMPT_COL[day], 0)
            ws.cell(row.row, COL_UNCHECKED, 0)
            ws.cell(row.row, COL_ATTEND_TIME, spec.attendance_time)
            ws.cell(row.row, COL_EFFECT, '正常')

        # ---- 合并区 ----
        for blk in spec.blocks:
            ws.cell(blk.start_row, COL_COLLEGE,
                    '%s（%d人）' % (blk.college, blk.college_total))
            ws.cell(blk.start_row, COL_COLLEGE).alignment = Alignment(
                horizontal='center', vertical='center')
            ws.merge_cells(start_row=blk.start_row, start_column=COL_COLLEGE,
                           end_row=blk.end_row, end_column=COL_COLLEGE)
            for _label, col in spec.rate_labels:
                if col in blk.skip_cols:
                    continue
                lo, hi = blk.partial_merges.get(col, (blk.start_row, blk.end_row))
                if hi > lo:
                    ws.merge_cells(start_row=lo, start_column=col,
                                   end_row=hi, end_column=col)

        for col, width in ((COL_SEQ, 6), (COL_COLLEGE, 18), (COL_ROOM, 10),
                           (COL_CLASSES, 26), (COL_EXPECTED, 10)):
            ws.column_dimensions[get_column_letter(col)].width = width
        for _label, col in spec.rate_labels:
            ws.column_dimensions[get_column_letter(col)].width = 12
    wb.save(path)
    wb.close()
    return path


def write_exempt_workbook(path: Path, sheets: Dict[str, List[ExemptRow]],
                           with_room: bool = False) -> Path:
    """写免勤表（大批量用 write_only 模式，避免内存爆炸）。"""
    wb = Workbook(write_only=True)
    for sheet_name, rows in sheets.items():
        ws = wb.create_sheet(sheet_name)
        ws.append(['合成数据·%s' % sheet_name])
        headers = list(EXEMPT_HEADERS)
        if with_room:
            headers.append(ROOM_HEADER)
        ws.append(headers)
        for row in rows:
            values = [row.seq, row.sid, row.college, row.name, row.class_text,
                      row.week_text, row.day_text, row.session, row.campus,
                      row.reason]
            if with_room:
                values.append(row.room)
            ws.append(values)
    wb.save(path)
    wb.close()
    return path


# ================================================================== 主流程


def generate(out_dir: Path, scale: str = 'small', dirty: bool = False,
             seed: int = 20261025, with_room: bool = False) -> Dict[str, object]:
    """生成一整套合成样本，返回统计信息。"""
    if scale not in SCALES:
        raise ValueError('未知规模：%s' % scale)
    spec = SCALES[scale]
    rng = random.Random(seed)

    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # ---- 明细表 ----
    blocks = plan_blocks(spec['rooms'])
    apply_template_defects(blocks)
    skeleton = plan_rooms(blocks, rng)
    plans: Dict[str, CollegePlan] = {}
    early = build_sheet(SHEET_EARLY, blocks, rng, plans, skeleton)

    late_blocks = plan_blocks(spec['rooms'])
    apply_template_defects(late_blocks)
    late = build_sheet(SHEET_LATE, late_blocks, rng, plans, skeleton)
    pingfeng = build_pingfeng_sheet(rng)

    detail_path = out_dir / 'detail.xlsx'
    write_detail_workbook(detail_path, [early, late, pingfeng])

    # ---- 免勤表 ----
    college_plans = [plans[name] for name in
                     [b.college for b in early.blocks] if name in plans]
    early_rows = build_exempt_rows(college_plans, spec['exempt_per_session'],
                                   EXEMPT_SHEET_EARLY, rng, dirty, 1)
    late_rows = build_exempt_rows(college_plans, spec['exempt_per_session'],
                                  EXEMPT_SHEET_LATE, rng, dirty,
                                  1 + spec['exempt_per_session'])
    if dirty:
        early_rows += inject_dirty_exempt_rows(early_rows, college_plans, rng,
                                                EXEMPT_SHEET_EARLY)
        late_rows += inject_dirty_exempt_rows(late_rows, college_plans, rng,
                                              EXEMPT_SHEET_LATE)
    for idx, row in enumerate(early_rows):
        row.seq = idx + 1
    for idx, row in enumerate(late_rows):
        row.seq = idx + 1

    exempt_path = out_dir / 'exempt.xlsx'
    write_exempt_workbook(exempt_path, {
        EXEMPT_SHEET_EARLY: early_rows,
        EXEMPT_SHEET_LATE: late_rows,
    }, with_room=with_room)

    return {
        'scale': scale,
        'dirty': dirty,
        'seed': seed,
        'with_room': with_room,
        'detail_path': str(detail_path),
        'exempt_path': str(exempt_path),
        'rooms_per_sheet': spec['rooms'],
        'detail_rows': len(early.rows) + len(late.rows) + len(pingfeng.rows),
        'exempt_rows': len(early_rows) + len(late_rows),
        'colleges': [p.name for p in college_plans],
    }


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description='生成三早一晚免勤统计系统的合成测试数据（纯虚构）')
    parser.add_argument('--out', default=str(DEFAULT_OUT_DIR),
                        help='输出目录（默认 <项目根>/_synthetic）')
    parser.add_argument('--scale', default='small', choices=sorted(SCALES),
                        help='数据规模：small / medium / large')
    parser.add_argument('--dirty', action='store_true',
                        help='注入第 3 类脏数据（重复、集体申请、越界、无法解析等）')
    parser.add_argument('--seed', type=int, default=20261025, help='随机种子')
    parser.add_argument('--with-room', action='store_true',
                        help='免勤表额外带「教室号」列（走精确匹配路径）')
    args = parser.parse_args(argv)

    info = generate(Path(args.out), scale=args.scale, dirty=args.dirty,
                    seed=args.seed, with_room=args.with_room)
    # 控制台按 GBK 输出，只打印 ASCII，避免中文乱码影响阅读
    print('[gen_synthetic] scale=%s dirty=%s seed=%s with_room=%s'
          % (info['scale'], info['dirty'], info['seed'], info['with_room']))
    print('[gen_synthetic] detail.xlsx  sheets=%d rows=%d rooms/sheet=%d'
          % (len(DETAIL_SHEETS), info['detail_rows'], info['rooms_per_sheet']))
    print('[gen_synthetic] exempt.xlsx  sheets=2 rows=%d'
          % info['exempt_rows'])
    print('[gen_synthetic] output_dir=%s' % Path(args.out))
    return 0


if __name__ == '__main__':
    sys.exit(main())
