# -*- coding: utf-8 -*-
"""数据模型：明细表结构与统计结果的内存表示。"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class RoomRow:
    """明细表里的一间教室（一行）。"""
    sheet: str
    row: int                       # Excel 行号（1 基）
    seq: Optional[int]
    college_name: str
    college_total: Optional[int]   # 学院名括号内标注的总人数
    room: str
    classes_text: str
    expected: Optional[int]
    actuals: Dict[str, Optional[int]] = field(default_factory=dict)   # 星期 -> 实到
    actual_raw: Dict[str, str] = field(default_factory=dict)          # 星期 -> 原始文本
    actual_cols: Dict[str, int] = field(default_factory=dict)         # 星期 -> 列号
    exempt_cols: Dict[str, int] = field(default_factory=dict)         # 星期 -> 免勤列号

    @property
    def is_yanshan(self) -> bool:
        return self.sheet.startswith('雁山')


@dataclass
class CollegeBlock:
    """明细表里一个学院的连续行块（用于出勤率纵向合并）。"""
    sheet: str
    name: str
    total: Optional[int]
    rows: List[int] = field(default_factory=list)
    rate_cols: Dict[str, int] = field(default_factory=dict)   # '周一'/'平均' -> 列号
    data_rows: List[int] = field(default_factory=list)        # 块内**教室数据行**

    @property
    def merge_rows(self) -> List[int]:
        """合并区覆盖的行（裁定 R7）。

        R7 ① 教室行 = 块内**含有效教室号且有应到数值**的行；纯表头行剔除。
        R7 ② 合并区起点 = 该学院**首个教室数据行**——第 3 行同时是
        「出勤率子表头行」与「地球科学学院 6504 的数据行」，它**必须**落在
        合并区内。旧实现把第 3 行硬编码排除（`r != 3`），导致 6504 游离于
        所有出勤率块之外，等于该教室永远没有出勤率。
        """
        if self.data_rows:
            return sorted(self.data_rows)
        return sorted(self.rows)


@dataclass
class DetailModel:
    """整份明细表的结构化表示（只含雁山两张表，屏风原样透传不参与计算）。"""
    path: str
    sheet_names: List[str] = field(default_factory=list)
    rooms: List[RoomRow] = field(default_factory=list)
    blocks: List[CollegeBlock] = field(default_factory=list)
    rate_cols_by_sheet: Dict[str, Dict[str, int]] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    exempt_cols_by_sheet: Dict[str, Dict[str, int]] = field(default_factory=dict)
    rate_label_rows: Dict[str, int] = field(default_factory=dict)
    college_exempt_layout: bool = False
    title_weeks: Dict[str, int] = field(default_factory=dict)
    protected_cols: Dict[str, set[int]] = field(default_factory=dict)

    def rooms_of(self, sheet: str) -> List[RoomRow]:
        return [r for r in self.rooms if r.sheet == sheet]

    def blocks_of(self, sheet: str) -> List[CollegeBlock]:
        return [b for b in self.blocks if b.sheet == sheet]


@dataclass
class ExemptRecord:
    """免勤表里的一条申请。"""
    sid: str
    name: str
    college: str
    class_text: str
    class_key: str
    weeks: set
    weekdays: set
    session: Optional[str]
    campus: str
    room: str
    reason: str
    group_size: Optional[int]
    source: str
    row: int
    sid_ok: bool = True          # 学号是否恰好 13 位
    sid_kind: str = ''           # 学号原始形态：整型 13 位 / 空 / 非数字学号 …
    placeholder: bool = False    # 是否是「全班」这类占位学号（不是真人）
    sid_padded: bool = False     # 整型学号是否已左补零到 13 位（裁定 R11 一层）
    whole_term: bool = False     # 周次写作「全学期」（裁定 R4：覆盖统计窗口周）
    pairs: frozenset = frozenset()   # {(周次:int, 星期:str)} 段内配对结果
    group_like: bool = False     # 写法像整班申请（但可能没写人数）
    source_ref: str = ''
    campus_defaulted: bool = False
    reason_category: str = '未分类'

    def days_for(self, week_no: int) -> set:
        """该记录在指定周次窗口内命中的星期集合。

        用「周次×星期」**事件对**，而不是独立的周次集/星期集——否则
        '第6周周一、第7周周三' 这类写法会把周一的免勤算进第 7 周（串周）。
        """
        if self.whole_term and week_no > 0:
            if self.pairs:
                days = {d for (w, d) in self.pairs if w == week_no}
                if days:
                    return days
            return set(self.weekdays)
        if self.pairs:
            return {d for (w, d) in self.pairs if w == week_no}
        return set(self.weekdays) if week_no in self.weeks else set()

    @property
    def is_group(self) -> bool:
        return self.group_size is not None

    @property
    def is_personal(self) -> bool:
        return bool(self.sid) and self.group_size is None

    @property
    def is_incomplete(self) -> bool:
        """没有有效人数的集体申请（含「全班」占位学号）。

        既不能按 1 人计，也不能覆盖同班个人申请，只能挂起等人补人数。
        """
        if self.group_size is not None and self.group_size > 0:
            return False
        if self.sid and not self.placeholder:
            return False
        return self.placeholder or self.group_like


@dataclass
class CollegeRate:
    """一个学院在某天某次自习的出勤率计算结果。"""
    sheet: str
    session: str
    college: str
    day: str
    expected: Optional[int]      # 学院总应到（取学院名括号人数）
    room_expected_sum: int       # 该学院各教室应到合计（仅用于交叉核对）
    actual: Optional[int]        # 学院总实到（任一教室缺失则为 None）
    exempt: int                  # 当次自习学院总免勤（唯一学号去重）
    rate: Optional[float]
    status: str
    note: str = ''
    anomalies: List[str] = field(default_factory=list)


@dataclass
class CollegeWeeklyRate:
    sheet: str
    session: str
    college: str
    calculated_days: int
    scheduled_days: int
    rate: Optional[float]
    needs_review: bool
