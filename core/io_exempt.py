# -*- coding: utf-8 -*-
"""免勤表读取、去重与聚合。

免勤表结构（真实样本）：
    第2行表头  序号 | 学号 | 学院 | 姓名 | 专业班级 | 免勤周次 | 免勤时间 |
              免勤自习 | 所在校区 | 申请免勤理由
    早自习免勤人员统计表 179 行；晚自习免勤人员统计表 3197 行

关键规则：
  * 去重键 = (学号, 星期, 早晚自习)；跨日期分别计，早晚自习各计一次。
  * 集体申请（无学号、含 'XX班全体学生（共N人）'）覆盖同班同日同自习的个人申请。
  * 屏风校区整体排除；'雁山校园'/'燕山校区' 归一为雁山。
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import replace
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook

from .model import ExemptRecord
from .parsing import (extract_group_size, is_group_application, normalize_class,
                      normalize_campus, normalize_room, normalize_text,
                      parse_int, parse_session, parse_sessions, parse_weekdays, parse_weeks,
                      resolve_pairs, strip_spaces)

_HEADER_HINTS = ('学号', '免勤周次')
_ROOM_HEADER_KEYS = ('教室号', '所在教室', '教室')

# 占位学号：这类记录没有真实学号、也没有「共N人」，不能当成一个人计入
PLACEHOLDER_SIDS = {'全班', '全体', '全部', '集体', '无', '待定', '学生', '同学', '/'}


def normalize_sid(value) -> tuple:
    """把学号单元格规整成字符串。

    Excel 里 13 位学号常被存成数字（float），直接 str() 会得到 '32620908172173.0'，
    既位数不对，也会与真正的字符串学号形成两个不同的去重键。

    返回 (学号字符串, 是否是 13 位标准学号, 原始形态描述)
    """
    if value is None:
        return '', False, '空'
    if isinstance(value, bool):
        return '', False, '布尔值'
    if isinstance(value, int):
        text = str(value)
        if len(text) < 13:
            # 裁定 R11 一层：整型学号先**左补零至 13 位**再进键空间。
            # 实测证据：本免勤表 13 位学号号段全为 '32xx'（3262/3252/3242），
            # 唯一那条 11 位整型补零后落在 '0032' —— 仍是表外号段，
            # 补零**不会与任何真实学号碰撞**（计数与「照常计入标待核」等价），
            # 但必须逐条留痕，便于人工判断它究竟是「Excel 丢前导零」还是「位数录错」。
            return text.zfill(13), True, ('整型 %d 位，已左补零至 13 位' % len(text))
        return text, len(text) == 13, ('整型 %d 位' % len(text))
    if isinstance(value, float):
        if value != value or value in (float('inf'), float('-inf')):
            return '', False, '非法数值'
        if float(value).is_integer():
            text = str(int(value))
            if len(text) < 13:
                return text.zfill(13), True, ('整型 %d 位，已左补零至 13 位' % len(text))
            return text, len(text) == 13, ('整型 %d 位' % len(text))
        return str(value), False, '小数型学号'
    text = strip_spaces(value).strip("'\"‘’“”`")
    if not text:
        return '', False, '空'
    if text.isdigit():
        return text, len(text) == 13, ('整型 %d 位' % len(text))
    return text, False, '非数字学号'


def is_placeholder_sid(sid: str) -> bool:
    return bool(sid) and sid in PLACEHOLDER_SIDS


def _one_digit_apart(a: str, b: str) -> bool:
    """两个学号是否只差「插入或删除一位」——用来识别疑似同一人重复登记。"""
    if not a or not b:
        return False
    if abs(len(a) - len(b)) != 1:
        return False
    short, long_ = (a, b) if len(a) < len(b) else (b, a)
    for i in range(len(long_)):
        if long_[:i] + long_[i + 1:] == short:
            return True
    return False


def normalize_student_name(name):
    return strip_spaces(name).replace('・', '·').replace('•', '·')


def _detect_header_row(ws, max_probe: int = 10) -> int:
    for row in range(1, min(max_probe, ws.max_row) + 1):
        if _has_exemption_headers(_detect_columns(ws, row)):
            return row
    return 2


def _detect_columns(ws, header_row: int) -> Dict[str, int]:
    cols: Dict[str, int] = {}
    room_cols: List[int] = []
    for col in range(1, ws.max_column + 1):
        head = strip_spaces(ws.cell(header_row, col).value)
        if not head:
            continue
        if head == '学号':
            cols['sid'] = col
        elif head == '姓名':
            cols['name'] = col
        elif head in ('学院', '学院名称', '所属学院'):
            cols['college'] = col
        elif '专业班级' in head or head == '班级':
            cols['class'] = col
        elif head in ('免勤周次', '请假周次') or '周次' in head:
            cols['week'] = col
        elif head in ('免勤时间', '免勤星期', '请假时间', '请假星期'):
            cols['time'] = col
        elif head in ('免勤自习', '请假自习') or head.startswith(('免勤自习', '请假自习')):
            cols['session'] = col
        elif head in ('所在校区', '校区') or head.startswith('所在校区'):
            cols['campus'] = col
        elif '理由' in head:
            cols['reason'] = col
        elif any(k in head for k in _ROOM_HEADER_KEYS):
            room_cols.append(col)
    if room_cols:
        cols['room'] = room_cols[0]
        if len(room_cols) > 1:
            cols['room2'] = room_cols[1]
    return cols


def _has_exemption_headers(cols):
    return ('college' in cols and any(key in cols for key in ('sid', 'name', 'class'))
            and any(key in cols for key in ('week', 'time')))


def load_exemptions(path: str, source: str = '未分类',
                    week_no: Optional[int] = None,
                    dictionary=None, *, require_schema=False, source_number=None,
                    keep_unassigned=False) -> List[ExemptRecord]:
    """读取免勤表，返回全部记录（尚未按周过滤与去重）。"""
    records: List[ExemptRecord] = []
    campus_map = getattr(dictionary, 'campus_alias', None) if dictionary else None
    class_map = getattr(dictionary, 'class_alias', None) if dictionary else None
    with open(path, 'rb') as f:
        wb = load_workbook(f, data_only=True)
        try:
            recognized = False
            for sheet_number, sheet_name in enumerate(wb.sheetnames, 1):
                ws = wb[sheet_name]
                header_row = _detect_header_row(ws)
                cols = _detect_columns(ws, header_row)
                if require_schema and not _has_exemption_headers(cols):
                    continue
                recognized = True
                sheet_session = parse_session(sheet_name)

                for row in range(header_row + 1, ws.max_row + 1):
                    def cell(key):
                        idx = cols.get(key)
                        return ws.cell(row, idx).value if idx else None

                    sid_raw, sid_ok, sid_kind = normalize_sid(cell('sid'))
                    name = normalize_text(cell('name'))
                    college = normalize_text(cell('college'))
                    class_text = normalize_text(cell('class'))
                    reason = normalize_text(cell('reason'))
                    week_text = normalize_text(cell('week'))
                    time_text = normalize_text(cell('time'))
                    raw_campus = cell('campus')
                    campus = normalize_campus(raw_campus, custom_map=campus_map)
                    if campus == '待确认' and not strip_spaces(raw_campus) and (sid_raw or college or class_text):
                        campus = '雁山'
                    sessions = parse_sessions(cell('session'), fallback_session=sheet_session, time_hint=time_text)
                    room1 = normalize_room(cell('room'))
                    room2 = normalize_room(cell('room2'))

                    if not any([sid_raw, name, college, class_text]):
                        continue
                    # 跳过模板顶部的张三/李四示例空行（无学号且无学院、无班级、无周次）
                    if not sid_raw and not college and not class_text and not week_text:
                        continue
                    if not keep_unassigned and not college and not class_text:
                        continue
                    if not college and class_text:
                        if any(k in class_text for k in ('服设', '服装', '数媒', '数字媒体', '动画', '产品', '环艺', '环设', '首饰')):
                            college = '艺术学院'

                    group_size = extract_group_size(name) or extract_group_size(class_text)
                    # 没有「共N人」的集体申请不构成集体申请：它既不覆盖同班个人，
                    # 也不能按 1 人计——按「免勤不完整（待补人数）」挂起。
                    whole_term = bool(re.search(r'(?:全|本|整|一)学期', f'{week_text} {time_text}'))
                    weeks = parse_weeks(week_text, week_no=week_no)
                    weekdays = parse_weekdays(time_text)
                    pairs_set, _warns = resolve_pairs(week_text, time_text,
                                                      window_week=week_no if week_no is not None else 7)
                    if not weeks and pairs_set:
                        weeks = {w for (w, _d) in pairs_set}
                    if not weekdays and pairs_set:
                        weekdays = {d for (_w, d) in pairs_set}

                    # 支持单格内用逗号/顿号填写的多名无学号学生（如 '梁烨霞，余文霞，...'）自动拆分为独立个人记录
                    if (not sid_raw and not group_size and not is_group_application(name, class_text)
                            and name and any(sep in name for sep in (',', '、'))):
                        sub_names = [n.strip() for n in re.split(r'[,、]+', name) if n.strip()]
                        if len(sub_names) >= 2 and all(2 <= len(n) <= 5 for n in sub_names):
                            name_entries = [(f'NOSID:{college}:{normalize_class(class_text)}:{n}', n, True, '姓名班级合成键') for n in sub_names]
                        else:
                            name_entries = [(sid_raw, name, sid_ok, sid_kind)]
                    else:
                        name_entries = [(sid_raw, name, sid_ok, sid_kind)]

                    for cur_sid, cur_name, cur_sid_ok, cur_sid_kind in name_entries:
                        for session in sessions:
                            if session == '晚自习' and room2:
                                room = room2
                            else:
                                room = room1 or room2
                            records.append(ExemptRecord(
                                sid=cur_sid,
                                name=cur_name,
                                college=college,
                                class_text=class_text,
                                class_key=normalize_class(class_text or cur_name, custom_map=class_map),
                                weeks=weeks,
                                weekdays=weekdays,
                                session=session,
                                campus=campus,
                                room=room,
                                reason=reason,
                                group_size=group_size,
                                source=source,
                                row=row,
                                sid_ok=cur_sid_ok,
                                sid_kind=cur_sid_kind,
                                placeholder=is_placeholder_sid(cur_sid),
                                sid_padded='补零' in (cur_sid_kind or ''),
                                whole_term=whole_term,
                                pairs=frozenset(pairs_set),
                                group_like=is_group_application(cur_name, class_text),
                                source_ref=(f'免勤表{source_number} / 工作表{sheet_number} / 第{row}行'
                                            if source_number else ''),
                            ))
        finally:
            wb.close()
    if require_schema and not recognized:
        raise ValueError('找不到免勤申请表头，请检查学院、人员或班级以及免勤周次/时间列')
    return records


class Aggregate:
    """按 (学院, 星期, 自习) 聚合后的免勤结果。"""

    def __init__(self):
        self.personal: Dict[Tuple[str, str, str], Dict[str, str]] = defaultdict(dict)
        self.groups: Dict[Tuple[str, str, str], Dict[str, int]] = defaultdict(dict)
        self.room_level: Dict[Tuple[str, str, str, str], int] = defaultdict(int)
        self.student_map: Dict[Tuple[str, str, str], List[dict]] = defaultdict(list)
        self.group_details: Dict[Tuple[str, str, str], List[dict]] = defaultdict(list)
        self.covered_personal = 0
        self.duplicate_events = 0
        self.total_records = 0
        self.invalid_records = 0
        self.week_filtered = 0
        self.campus_filtered = 0
        self.group_applications = 0
        self.unique_sids = 0
        self.unmatched_week = 0
        self.unmatched_day = 0
        self.room_matched = 0
        self.roomless = 0
        self.odd_sid_records = 0
        self.incomplete_records = 0
        self.incomplete_scopes = set()
        self.same_person_pairs = 0
        self.padded_sids = 0
        self.cross_college_conflicts = 0
        self.conflicts: List[dict] = []
        self.anomalies: List[dict] = []
        self.outside_attendance_records = 0
        self.unknown_college_records = 0
        self.identity_conflicts = 0
        self.held_identity_events = 0
        self.name_matched_records = 0
        self.identity_review_scopes = set()
        self.unknown_college_scopes = set()

    def exempt_breakdown(self, college: str, day: str, session: str) -> tuple[int, int]:
        """返回去重且剔除同班集体覆盖后的个人数，以及集体免勤人数。

        = Σ 有集体申请的班级人数 + 未被集体申请覆盖的个人去重人数
        集体申请只覆盖**同班**的个人申请，同学院其它班级的人照常计入。
        """
        key = (college, day, session)
        group_by_class = self.groups.get(key, {})
        group_count = sum(group_by_class.values())
        personal = self.personal.get(key, {})
        if isinstance(personal, dict):
            # 结构：{学号: 班级}，集体申请只覆盖同班
            free = sum(1 for _sid, cls in personal.items() if cls not in group_by_class)
        else:
            # 兼容只有学号集合的旧结构
            free = len(personal)
        return free, group_count

    def exempt_count(self, college: str, day: str, session: str) -> int:
        return sum(self.exempt_breakdown(college, day, session))

    def get_students(self, college: str, day: str, session: str) -> List[dict]:
        """返回该学院在指定自习时段去重后的有效免勤学生明细列表。"""
        return self.student_map.get((college, day, session), [])

    def scope_keys(self):
        return set(self.personal) | set(self.groups)


def _audit_personal_identities(records, days_by_record, agg):
    candidates = defaultdict(set)
    for rec in records:
        name = normalize_student_name(rec.name)
        if rec.sid_ok and rec.sid.isdigit() and len(rec.sid) == 13 and name and rec.class_key:
            candidates[(rec.college, name, rec.class_key)].add(rec.sid)

    resolved = []
    by_sid = defaultdict(list)
    for rec in records:
        days = days_by_record[id(rec)]
        name = normalize_student_name(rec.name)
        matches = candidates.get((rec.college, name, rec.class_key), set()) if name and rec.class_key else set()
        if not rec.sid_ok or not rec.sid:
            if len(matches) == 1:
                candidate = next(iter(matches))
                if not rec.sid or _one_digit_apart(rec.sid, candidate):
                    rec = replace(rec, sid=candidate, sid_ok=True)
                    agg.name_matched_records += 1
                    agg.same_person_pairs += 1
                else:
                    matches = {candidate, rec.sid}
            elif (not rec.sid and len(matches) == 0 and name and rec.class_key
                  and not rec.placeholder and not rec.group_like
                  and not name.endswith(('专业', '班'))
                  and name not in ('整班', '整个班级', '新能源、高分子、金属')):
                rec = replace(rec, sid=f'NOSID:{rec.college}:{rec.class_key}:{name}', sid_ok=True)
            if not rec.sid or len(matches) > 1:
                agg.invalid_records += 1
                agg.identity_review_scopes.update((rec.college, day, rec.session) for day in days)
                agg.anomalies.append({
                    'type': '学号无法核对', 'level': 'warn',
                    'scope': f'{rec.college} / {rec.source_ref or rec.row}',
                    'review_ref': rec.source_ref,
                    'note': '缺少学号或有多个可能身份，暂不计入。请核对姓名、班级和学号。',
                })
                continue
            if not rec.sid_ok:
                agg.identity_review_scopes.update((rec.college, day, rec.session) for day in days)
        if not name:
            agg.identity_review_scopes.update((rec.college, day, rec.session) for day in days)
            agg.anomalies.append({
                'type': '姓名缺失', 'level': 'warn',
                'scope': f'{rec.college} / {rec.source_ref or rec.row}',
                'review_ref': rec.source_ref,
                'note': '按学号计入，姓名待核对。',
            })
        by_sid[rec.sid].append((rec, name, days))
        resolved.append((rec, days))

    held = set()
    for sid, entries in by_sid.items():
        colleges = {rec.college for rec, _, _ in entries}
        names = {name for _, name, _ in entries if name}
        if len(colleges) <= 1 and len(names) <= 1:
            continue
        agg.identity_conflicts += 1
        name_by_event = defaultdict(set)
        for rec, name, days in entries:
            for day in days:
                agg.identity_review_scopes.add((rec.college, day, rec.session))
                event = (sid, day, rec.session)
                if name:
                    name_by_event[event].add(name)
                if len(colleges) > 1:
                    held.add(event)
        held.update(event for event, values in name_by_event.items() if len(values) > 1)
        refs = list(dict.fromkeys(rec.source_ref or f'第{rec.row}行' for rec, _, _ in entries))
        agg.anomalies.append({
            'type': '学号对应学院或姓名冲突', 'level': 'warn',
            'scope': '、'.join(sorted(colleges)) + ' / ' + '；'.join(refs[:8]),
            'review_ref': '；'.join(refs[:8]),
            'note': '同学号跨学院的记录、同一时段姓名冲突的记录暂不计入。仅跨时段姓名有异的记录按学号计入并标为待核对。',
        })
    agg.held_identity_events = len(held)
    return resolved, held


def aggregate_exemptions(records: List[ExemptRecord], week_no: int, campus: str = '雁山',
                         *, identity_check=False, attendance_scopes=None, allowed_colleges=None) -> Aggregate:
    """按周次、校区过滤并去重聚合。

    去重键 (学号, 星期, 自习)：同一学生同一周内同一天同一次自习只算一次。
    集体申请覆盖同班同日同自习的个人申请。
    """
    agg = Aggregate()
    agg.total_records = len(records)

    # 先收集集体申请（它们决定哪些个人申请被覆盖）
    group_records: List[ExemptRecord] = []
    personal_records: List[ExemptRecord] = []
    days_by_record = {}

    for rec in records:
        if rec.campus != campus:
            agg.campus_filtered += 1
            continue
        if not rec.weeks and not rec.whole_term:
            # 周次无法识别：无法证明该申请适用于本周，不计入，列入异常清单
            agg.unmatched_week += 1
            agg.anomalies.append({
                'type': '免勤周次无法识别',
                'level': 'warn',
                'scope': '%s / %s（源行 %d）' % (rec.college, rec.class_text, rec.row),
                'note': '该记录的免勤周次无法解析，未计入本次统计，请人工确认',
            })
            continue
        if week_no not in rec.weeks and not rec.whole_term:
            agg.week_filtered += 1
            continue
        if not rec.weekdays:
            agg.unmatched_day += 1
            agg.anomalies.append({
                'type': '免勤星期无法识别',
                'level': 'warn',
                'scope': '%s / %s（源行 %d）' % (rec.college, rec.class_text, rec.row),
                'note': '该记录的免勤时间无法解析为具体星期，未计入本次统计，请人工确认',
            })
            continue
        active_days = rec.days_for(week_no)
        if not active_days:
            agg.week_filtered += 1
            continue
        if identity_check and not rec.session:
            agg.invalid_records += 1
            agg.anomalies.append({
                'type': '自习时段无法识别', 'level': 'warn',
                'scope': f'{rec.college} / {rec.source_ref or rec.row}',
                'review_ref': rec.source_ref,
                'note': '请确认该申请适用于早自习还是晚自习。',
            })
            continue
        if attendance_scopes is not None:
            active_days = {day for day in active_days if (day, rec.session) in attendance_scopes}
            if not active_days:
                agg.outside_attendance_records += 1
                continue
        if allowed_colleges is not None and rec.college not in allowed_colleges:
            agg.unknown_college_records += 1
            agg.unknown_college_scopes.update((rec.college, day, rec.session) for day in active_days)
            agg.anomalies.append({
                'type': '学院无法匹配', 'level': 'warn',
                'scope': f'{rec.college or "学院空白"} / {rec.source_ref or rec.row}',
                'review_ref': rec.source_ref,
                'note': '学院未在考勤表中找到，请核对原表或在词典中补充学院别名。',
            })
            continue
        days_by_record[id(rec)] = active_days
        size = rec.group_size or 0
        if size > 0:
            group_records.append(rec)
        elif (rec.sid and not rec.placeholder) or (identity_check and rec.name and not rec.is_incomplete):
            personal_records.append(rec)
            if rec.sid_padded:
                agg.padded_sids += 1
                agg.anomalies.append({
                    'type': '整型学号已左补零',
                    'level': 'warn',
                    'scope': '%s / %s（源行 %d）' % (rec.college, rec.class_text, rec.row),
                    'review_ref': rec.source_ref,
                    'note': '整数格式学号已补前导零，请与原始学号核对。',
                })
            if rec.sid and not rec.sid_ok:
                # 学号位数异常：可能是手工录入笔误，也可能与另一条记录指向同一人。
                # 不擅自丢弃（会漏算真人），照常计入但必须标记人工核对。
                agg.odd_sid_records += 1
                agg.anomalies.append({
                    'type': '学号位数异常',
                    'level': 'warn',
                    'scope': '%s / %s（源行 %d）' % (rec.college, rec.class_text, rec.row),
                    'review_ref': rec.source_ref,
                    'note': '学号不是 13 位。请核对原表及该人员的其他申报。',
                })
        else:
            # 没有有效人数：占位学号（「全班」这类）或整班申请但没写「共N人」。
            # 既不能按 1 人计，也不能覆盖同班个人申请——挂起等人补人数。
            if rec.placeholder or rec.group_like:
                agg.incomplete_records += 1
                for day in active_days:
                    agg.incomplete_scopes.add((rec.college, day, rec.session or ''))
                agg.anomalies.append({
                    'type': '免勤不完整（待补人数）',
                    'level': 'warn',
                    'scope': '%s / %s（源行 %d）%s' % (
                        rec.college, rec.class_text, rec.row,
                        ' / ' + rec.source_ref if identity_check and rec.source_ref else ''),
                    'review_ref': rec.source_ref,
                    'note': '该记录没有有效人数（占位学号或未写「共N人」），无法确定人数，'
                            '既未计入也未覆盖同班个人申请，请人工补齐人数',
                })
            else:
                agg.invalid_records += 1
                agg.anomalies.append({
                    'type': '免勤记录无法定位',
                    'level': 'warn',
                    'scope': '%s / %s（源行 %d）' % (rec.college, rec.class_text, rec.row),
                    'note': '该记录既无学号也不是集体申请，无法计入任何免勤人数',
                })

    # 集体申请：同 (学院, 班级key, 星期, 自习) 覆盖个人
    group_index: Dict[Tuple[str, str, str, str], ExemptRecord] = {}
    for rec in group_records:
        days = days_by_record[id(rec)]
        for day in days:
            key = (rec.college, rec.class_key, day, rec.session or '')
            prev = group_index.get(key)
            if prev is None:
                group_index[key] = rec
            else:
                # 同一班级同日同自习多条集体申请：取较大声明人数，并记异常
                agg.duplicate_events += 1
                agg.anomalies.append({
                    'type': '集体免勤重复申报',
                    'level': 'warn',
                    'scope': f'{rec.college} / {rec.class_text} / {day} / {rec.session}',
                    'note': '同一班级同一天同一次自习存在多条集体免勤声明，已取较大值',
                })
                if (rec.group_size or 0) > (prev.group_size or 0):
                    group_index[key] = rec

    for (college, class_key, day, session), rec in group_index.items():
        size = rec.group_size or 0
        agg.groups[(college, day, session)][class_key] = size
        agg.group_applications += 1
        agg.group_details[(college, day, session)].append({
            'college': college,
            'class_text': rec.class_text,
            'class_key': class_key,
            'group_size': size,
            'room': rec.room,
            'reason': rec.reason,
            'row': rec.row,
        })
        if rec.room:
            agg.room_level[(college, day, session, rec.room)] += size

    # 个人申请：按「天」判断是否被集体申请覆盖。
    # 注意：集体申请只覆盖**同班同一天同一次自习**，不能整条丢弃——
    # 一个学生若同时申请了周一和周三，周三被集体覆盖后，周一依然要计入。
    # 同时保留「班级归属」，以便集体申请只覆盖同班、其它班级的人照常计入。
    seen: Dict[Tuple[str, str, str], Dict[str, str]] = defaultdict(dict)
    rec_of: Dict[Tuple[str, str, str], Dict[str, ExemptRecord]] = defaultdict(dict)
    abnormal: Dict[Tuple[str, str, str], set] = defaultdict(set)
    row_of: Dict[Tuple[str, str, str], Dict[str, int]] = defaultdict(dict)
    room_of: Dict[Tuple[str, str, str], Dict[str, str]] = defaultdict(dict)
    global_seen: Dict[Tuple[str, str, str], str] = {}
    if identity_check:
        person_days, held = _audit_personal_identities(personal_records, days_by_record, agg)
    else:
        person_days = [(rec, days_by_record[id(rec)]) for rec in personal_records]
        held = set()
    for rec, days in sorted(person_days, key=lambda item: (item[0].college, item[0].sid, item[0].row)):
        session = rec.session or ''
        for day in days:
            if (rec.sid, day, session) in held:
                continue
            gkey = (rec.college, rec.class_key, day, session)
            if gkey in group_index:
                agg.covered_personal += 1
                continue
            key = (rec.college, day, session)
            gscope = (rec.sid, day, session)
            if rec.sid in seen[key]:
                if rec.room and not room_of[key].get(rec.sid):
                    room_of[key][rec.sid] = rec.room
                agg.duplicate_events += 1
                continue
            owner = global_seen.get(gscope)
            if owner is not None:
                # 同一学号出现在两个学院：按「免勤按唯一学号去重」，只计一次
                agg.cross_college_conflicts += 1
                agg.anomalies.append({
                    'type': '同一学号出现在多个学院',
                    'level': 'warn',
                    'scope': '%s / %s / %s（另一学院：%s）' % (rec.college, day, session, owner),
                    'note': '该学号在本次统计范围里已计入 %s，不再重复计入本学院，请人工核对归属' % owner,
                })
                continue
            global_seen[gscope] = rec.college
            seen[key][rec.sid] = rec.class_key
            rec_of[key][rec.sid] = rec
            row_of[key][rec.sid] = rec.row
            if rec.room:
                room_of[key][rec.sid] = rec.room
            if not rec.sid_ok and not identity_check:
                abnormal[key].add(rec.sid)

    # 疑似同一人：同一 (学院, 星期, 自习) 内，异常学号与 13 位学号只差「插入/删除一位」
    # 时判为同一人重复登记，该格按 1 人计（同一份免勤不重复相加），两条原文交人工裁定。
    for key, bad_sids in abnormal.items():
        sids = seen.get(key)
        if not sids:
            continue
        normal = [s for s in sids if s not in bad_sids]
        for bad in sorted(bad_sids):
            for good in sorted(normal):
                if _one_digit_apart(bad, good):
                    sids.pop(bad, None)
                    rec_of[key].pop(bad, None)
                    agg.same_person_pairs += 1
                    rows_txt = ','.join(str(row_of.get(key, {}).get(s, 0))
                                          for s in (bad, good))
                    agg.anomalies.append({
                        'type': '疑似同一人重复登记',
                        'level': 'warn',
                        'scope': '%s / %s / %s（源行 %s）' % (key[0], key[1], key[2], rows_txt),
                        'note': '该范围内有两个学号仅相差一位（疑似同一人重复登记），'
                                '已按 1 人计，请人工裁定',
                    })
                    break

    for key, sids in seen.items():
        group_by_class = agg.groups.get(key, {})
        for sid, cls in sids.items():
            if cls in group_by_class:
                continue
            rm = room_of[key].get(sid, '')
            if rm:
                agg.room_level[(key[0], key[1], key[2], rm)] += 1
            rec = rec_of[key].get(sid)
            if rec:
                agg.student_map[key].append({
                    'sid': rec.sid,
                    'name': rec.name,
                    'college': rec.college,
                    'class_text': rec.class_text,
                    'class_key': rec.class_key,
                    'reason': rec.reason,
                    'room': rm,
                    'row': rec.row,
                    'campus': rec.campus,
                    'source_ref': rec.source_ref,
                })

    agg.personal = seen
    agg.unique_sids = len({sid for records_by_sid in seen.values() for sid in records_by_sid}) if identity_check else len({r.sid for r in personal_records})
    agg.room_matched = sum(1 for r in records if r.room)
    agg.roomless = sum(1 for r in records if not r.room)
    return agg
