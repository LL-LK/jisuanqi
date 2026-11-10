# -*- coding: utf-8 -*-
"""统计计算层：教室级免勤分摊 + 学院出勤率。

口径（以用户最终裁定为准，与对话口径一致）：
    教室出勤率 = 实到 ÷ (应到 − 免勤)
    学院出勤率 = 学院总实到 ÷ (学院总应到 − 当次自习学院总免勤)
    学院总应到 = 明细表学院名括号内标注人数；教室应到合计仅作交叉核对。
    学院总免勤 = 唯一学号去重后的计数（不把教室分摊值重复累加）。
"""

from __future__ import annotations

from collections import defaultdict
from typing import Dict, List, Optional, Tuple

from .io_exempt import Aggregate
from .model import CollegeBlock, CollegeRate, DetailModel, RoomRow
from .parsing import match_class_to_room_classes, parse_session

DAYS = ('周一', '周三', '周四', '周五')

# R17 固定口径：估算标记的色值（导出层与界面共用同一常量，禁止各自写死一份）
ESTIMATE_RGB = 'FFFFF3CD'


def estimate_source(estimated_cells, key) -> str:
    """**估算判定的唯一实现**（导出层与界面共用，禁止第二处判定）。

    key = (工作表, 行号, 星期)。命中 ⇒ 该格免勤数是「按比例分摊」的估算值 ⇒ '估算'；
    否则 ⇒ '精确'（免勤表里带教室号、逐人核对过的值）。
    注意：值为 0 的格同样属估算（0 也是分摊结果，不是核对过的名单）。
    """
    if not estimated_cells:
        return '精确'
    return '估算' if key in estimated_cells else '精确'


def allocate(total: int, weights: List[int], order=None) -> List[int]:
    """按权重把整数 total 分配到各项，最大余数法保证各项之和恒等于 total。

    order：并列名次的落位顺序（默认按数组下标）。业务上应传入**教室号数值升序**，
    这是源表里唯一稳定且可复现的业务键，保证同一份数据每次分摊结果一致。

    >>> allocate(5, [3, 7])
    [2, 3]
    >>> sum(allocate(7, [2, 0, 5]))
    7
    """
    total = int(total)
    if total <= 0 or not weights:
        return [0] * len(weights)
    if order is None:
        order = list(range(len(weights)))
    weight_sum = sum(w for w in weights if w and w > 0)
    if weight_sum <= 0:
        # 没有有效权重：全部落到第一项，保证和相等
        result = [0] * len(weights)
        result[0] = total
        return result

    exact = [total * (w if w and w > 0 else 0) / weight_sum for w in weights]
    base = [int(x) for x in exact]
    # 排序规则：余数降序 → 权重升序 → 教室号升序（order）。
    # 后两条是并列名次的落位依据，保证同一份数据每次分摊结果完全一致。
    ranked = sorted(range(len(weights)),
                    key=lambda i: (-(exact[i] - base[i]), (weights[i] or 0), order[i]))
    missing = total - sum(base)
    for pos in range(missing):
        base[ranked[pos % len(ranked)]] += 1
    return base


def room_sort_key(room: str):
    """教室号排序键：优先按数值升序，非数字的排在后面按字符串排。"""
    text = ''.join(ch for ch in str(room) if ch.isdigit())
    if text:
        return (0, int(text), '')
    return (1, 0, str(room))


def _college_key(name: str) -> str:
    return ''.join(str(name).split())


class Calculator:
    """把明细表结构与免勤聚合结果合成教室级免勤数与学院级出勤率。"""

    def __init__(self, detail: DetailModel, agg: Aggregate,
                 alias: Optional[Dict[str, str]] = None,
                 room_level: Optional[Dict[Tuple[str, str, str, str], int]] = None,
                 complete_mode: bool = False,
                 match_classes: bool = True):
        self.detail = detail
        self.agg = agg
        self.alias = alias or {}
        if room_level is not None:
            self.room_level = room_level
        else:
            self.room_level = getattr(agg, 'room_level', None) or {}
        self.complete_mode = complete_mode
        self.match_classes = match_classes
        self.overflow: List[Tuple[str, str, str, str, int]] = []
        # 估算标记：按比例分摊出来的教室级免勤属于估算值，导出时要在格子上打标
        self.estimated_cells: set = set()
        # 穿透映射：记录每个教室单元格对应的学生明细列表 (sheet, row, day) -> list[dict]
        self.room_students_map: Dict[Tuple[str, int, str], List[dict]] = defaultdict(list)

        # 明细表学院名 → 免勤表学院名
        self._name_map: Dict[str, str] = {}
        exempt_colleges = {k[0] for k in agg.scope_keys()}
        exempt_colleges |= {k[0] for k in getattr(agg, 'incomplete_scopes', set())}
        detail_colleges = {b.name for b in detail.blocks}
        for name in detail_colleges:
            self._name_map[name] = name
        for name in detail_colleges:
            if name in exempt_colleges:
                continue
            target = self.alias.get(name)
            if target and target in exempt_colleges:
                self._name_map[name] = target
                continue
            rev_hit = next((k for k, v in self.alias.items()
                            if v == name and k in exempt_colleges), None)
            if rev_hit:
                self._name_map[name] = rev_hit
                continue
            compact = _college_key(name)
            for cand in sorted(exempt_colleges):
                if _college_key(cand) == compact:
                    self._name_map[name] = cand
                    break
            else:
                if target:
                    self._name_map[name] = target

    # ------------------------------------------------------------ 免勤分摊
    def exempt_college(self, college: str) -> str:
        """返回明细表学院在免勤聚合中的名称。"""
        return self._name_map.get(college, college)

    def exempt_for(self, college: str, day: str, session: str) -> int:
        return self.agg.exempt_count(self.exempt_college(college), day, session)

    def room_exemptions(self) -> Dict[Tuple[str, int, str], int]:
        """返回 {(工作表, 行号, 星期): 免勤人数}。

        * 星期范围取该表**所有**带免勤人数列的星期——即使某个星期不在出勤率区
          （例如晚自习表的周五），也照常填数，绝不把原表已有数值覆盖成 0。
        * 同一工作表下同名学院若分散在多个块，合并后统一分摊一次，杜绝重复分摊。
        * 教室人数按应到人数比例分摊，并以该教室应到人数封顶；超出部分
          记入超额异常，不做二次分摊（学院总免勤仍按唯一学号去重值计算）。
        """
        if getattr(self, '_room_exemptions_cache', None) is not None:
            return self._room_exemptions_cache
        out: Dict[Tuple[str, int, str], int] = {}
        self.overflow = []
        self.estimated_cells = set()
        self.room_students_map = defaultdict(list)
        room_counts = defaultdict(dict)
        for (college, day, session, room), count in self.room_level.items():
            room_counts[(college, day, session)][room] = count
        for sheet in self.detail.rate_cols_by_sheet:
            session = parse_session(sheet) or ''
            blocks = self.detail.blocks_of(sheet)
            by_name: Dict[str, List[CollegeBlock]] = defaultdict(list)
            for block in blocks:
                by_name[block.name].append(block)

            for college_name, blk_list in by_name.items():
                real = self.exempt_college(college_name)
                rooms = [r for r in self.detail.rooms_of(sheet)
                         if any(r.row in b.rows for b in blk_list) and r.room]
                if not rooms:
                    continue
                days = []
                for room in rooms:
                    for day in room.exempt_cols:
                        if day not in days:
                            days.append(day)
                for day in days:
                    day_rooms = [room for room in rooms if day in room.exempt_cols]
                    total = self.exempt_for(college_name, day, session)
                    if total <= 0:
                        for room in day_rooms:
                            out[(sheet, room.row, day)] = 0
                            # 免勤表无人数或为0，标记为估算
                            self.estimated_cells.add((sheet, room.row, day))
                        continue

                    # 1. 深度识别学院去重后的免勤学生名单，建立班级/教室两级精确落位
                    students = self.agg.get_students(real, day, session)
                    room_by_num = {}
                    rows_by_compact_number = defaultdict(list)
                    for room in day_rooms:
                        room_by_num.setdefault(room.room, room.row)
                        rows_by_compact_number[room.room.lstrip('0')].append(room.row)

                    def resolve_room(number):
                        if not number:
                            return None
                        if number in room_by_num:
                            return room_by_num[number]
                        rows = rows_by_compact_number.get(number.lstrip('0'), [])
                        return rows[0] if len(rows) == 1 else None

                    room_students = defaultdict(list)
                    exact_hits = defaultdict(int)
                    inferred_rows = set()
                    for number, count in room_counts[(real, day, session)].items():
                        matched = resolve_room(number)
                        if matched is not None:
                            exact_hits[matched] += count

                    # 对去重后的免勤学生进行落位分类
                    for stu in students:
                        matched = resolve_room(stu.get('room') or '')
                        if matched is None and self.match_classes:
                            # 未命中教室号时，仅在班级唯一匹配时精确落位。
                            stu_cls = stu.get('class_text', '')
                            cands = [r.row for r in day_rooms
                                     if match_class_to_room_classes(stu_cls, r.classes_text)]
                            if len(cands) == 1:
                                matched = cands[0]
                                exact_hits[matched] += 1
                                inferred_rows.add(matched)
                        if matched is not None:
                            room_students[matched].append(stu)

                    assigned_exact = sum(exact_hits.values())
                    leftover = max(total - assigned_exact, 0)
                    share = [0] * len(day_rooms)
                    if leftover:
                        # 存在未能匹配抽查教室的学生余量（或集体免勤）：按剩余容量进行最大余数法分摊
                        weights = [max((r.expected or 0) - exact_hits[r.row], 0) for r in day_rooms]
                        if not any(weights):
                            weights = [max(r.expected or 0, 0) for r in day_rooms]
                        share = allocate(leftover, weights, [room_sort_key(r.room) for r in day_rooms])

                    # 2. 按教室应到人数上限严格封顶（C1 口径），杜绝穿透
                    for r, add_val in zip(day_rooms, share):
                        key = (sheet, r.row, day)
                        base_val = exact_hits[r.row]
                        out[key] = base_val + add_val
                        self.room_students_map[key].extend(room_students[r.row])
                        if r.row in inferred_rows or (leftover and (add_val > 0 or base_val == 0)):
                            self.estimated_cells.add(key)
                        cap = r.expected
                        if cap is not None and out[key] > cap:
                            self.overflow.append(
                                (sheet, college_name, day, r.room, out[key] - cap))
                            out[key] = cap
        self._room_exemptions_cache = out
        return out

    # ------------------------------------------------------------ 学院出勤率
    def college_rates(self) -> List[CollegeRate]:
        results: List[CollegeRate] = []
        for sheet, rate_cols in self.detail.rate_cols_by_sheet.items():
            session = parse_session(sheet) or ''
            blocks = self.detail.blocks_of(sheet)
            # 按学院名聚合（同名学院可能分散在多个块）
            by_name: Dict[str, List[CollegeBlock]] = defaultdict(list)
            for block in blocks:
                by_name[block.name].append(block)

            for college, blk_list in by_name.items():
                rooms = [r for r in self.detail.rooms_of(sheet)
                         if any(r.row in b.rows for b in blk_list)]
                total_label = next((b.total for b in blk_list if b.total is not None), None)
                room_expected_sum = sum(max(r.expected or 0, 0) for r in rooms)
                expected = total_label
                if expected is None and rooms and all(r.expected is not None for r in rooms):
                    expected = room_expected_sum

                day_rates: List[float] = []
                for day in rate_cols:
                    if day == '平均':
                        continue
                    day_rooms = [r for r in rooms if day in r.actuals]
                    actual_values = [r.actuals.get(day) for r in day_rooms]

                    exempt = self.exempt_for(college, day, session)
                    denominator = (expected - exempt) if expected is not None else None

                    anomalies: List[str] = []
                    if total_label is not None and room_expected_sum != total_label:
                        anomalies.append(
                            '学院标注总人数 %d 与各教室应到合计 %d 相差 %d 人'
                            % (total_label, room_expected_sum, room_expected_sum - total_label))

                    if not self.complete_mode:
                        # 严格契约模式（向后兼容）：实到缺失留空，分母非正不出率
                        missing = [r for r, v in zip(day_rooms, actual_values) if v is None]
                        actual_sum = None if missing or not day_rooms else sum(actual_values)
                        if not day_rooms:
                            status = '实到数据未录入'
                            note = '该学院当次自习没有可用的教室实到栏，按约定留空。'
                            rate = None
                        elif missing:
                            status = '实到数据未完整填写'
                            exempt_marked = [r.room for r in missing
                                             if any(k in (r.actual_raw.get(day) or '') for k in ('免勤', '请假', '未上', '无'))]
                            if exempt_marked:
                                note = ('该学院该次自习共 %d 间教室，其中 %d 间未填实到数值'
                                        '（含实到栏填“免勤”教室：%s），按约定留空且不生成出勤率。'
                                        % (len(day_rooms), len(missing), '、'.join(exempt_marked)))
                            else:
                                note = ('该学院该次自习共 %d 间教室，其中 %d 间实到人数缺失，'
                                        '按约定留空且不生成出勤率。'
                                        % (len(day_rooms), len(missing)))
                            rate = None
                        elif denominator is None:
                            status = '应到人数缺失，无法计算'
                            note = '学院未标注应到总人数，且有教室应到人数缺失。'
                            rate = None
                        elif denominator <= 0:
                            status = '分母非正数，无法计算'
                            note = ('学院总应到 %s − 当次总免勤 %d = %s，分母非正，不生成出勤率。'
                                    % (expected, exempt, denominator))
                            rate = None
                        else:
                            rate = actual_sum / denominator
                            day_rates.append(rate)
                            real_college = self._name_map.get(college, college)
                            inc_scopes = getattr(self.agg, 'incomplete_scopes', set())
                            incomplete = ((college, day, session) in inc_scopes
                                          or (real_college, day, session) in inc_scopes)
                            if incomplete:
                                status = '免勤不完整（待补人数）'
                                note = ('该学院该次存在无人数的占位免勤申请，免勤人数不完整，'
                                        '以下数值仅供参考、不得以完整名义上报；'
                                        '按公式算得 %.2f%%。' % (rate * 100))
                                anomalies.append('免勤人数不完整，待补人数')
                            elif rate > 1:
                                status = '已计算（超过100%，需复核）'
                                note = ('按公式算得 %.2f%%，实到合计 %d 大于考勤基数 %d，'
                                        '需核对免勤与实到的统计口径。'
                                        % (rate * 100, actual_sum, denominator))
                                anomalies.append('出勤率超过 100%')
                            else:
                                status = '完整'
                                note = ('学院出勤率 = 学院总实到 %d ÷ (学院总应到 %s − 当次总免勤 %d)。'
                                        % (actual_sum, expected, exempt))
                    else:
                        # 容错完整计算模式（新生产特性）：实到免勤折算为0、部分教室缺失按有效教室计算、分母<=0按满勤100%兜底
                        effective_values = []
                        exempt_marked = []
                        for r in day_rooms:
                            raw_txt = r.actual_raw.get(day) or ''
                            if any(k in raw_txt for k in ('免勤', '请假', '未上', '无', '放假', '公假')):
                                effective_values.append(0)
                                exempt_marked.append(r.room)
                            else:
                                effective_values.append(r.actuals.get(day))

                        valid_pairs = [(r, v) for r, v in zip(day_rooms, effective_values) if v is not None]
                        missing_rooms = [r for r, v in zip(day_rooms, effective_values) if v is None]

                        if not valid_pairs:
                            status = '实到数据未录入'
                            note = ('该学院该次自习共 %d 间教室均未录入实到人数，按约定留空。'
                                    % len(day_rooms))
                            rate = None
                            actual_sum = None
                        else:
                            actual_sum = sum(v for _, v in valid_pairs)
                            if missing_rooms:
                                valid_expected = None
                                if all(r.expected is not None for r, _ in valid_pairs):
                                    valid_expected = sum(max(r.expected, 0) for r, _ in valid_pairs)
                                room_exempts = getattr(self, '_room_exemptions_cache', None)
                                if room_exempts is None:
                                    room_exempts = self.room_exemptions()
                                valid_exempt = sum(room_exempts.get((sheet, r.row, day), 0) for r, _ in valid_pairs)
                                eff_denominator = None if valid_expected is None else valid_expected - valid_exempt
                                eff_expected = valid_expected
                                eff_exempt = valid_exempt
                            else:
                                eff_denominator = denominator
                                eff_expected = expected
                                eff_exempt = exempt

                            if (denominator is not None and denominator <= 0) or (eff_denominator is not None and eff_denominator <= 0):
                                rate = 1.0
                                day_rates.append(rate)
                                status = '全员免勤（满勤）'
                                note = ('学院总应到 %s − 当次总免勤 %d = %s ≤ 0，按 100.00%% 满勤出勤计入。'
                                        % (expected if (denominator is not None and denominator <= 0) else eff_expected,
                                           exempt if (denominator is not None and denominator <= 0) else eff_exempt,
                                           denominator if (denominator is not None and denominator <= 0) else eff_denominator))
                                anomalies.append('全员免勤（免勤人数达到或超出应到人数）')
                            elif eff_denominator is None:
                                status = '无法确定应到人数'
                                note = '学院未标注总人数且无有效教室应到合计，无法确定考勤分母。'
                                rate = None
                            else:
                                rate = actual_sum / eff_denominator
                                day_rates.append(rate)
                                real_college = self._name_map.get(college, college)
                                inc_scopes = getattr(self.agg, 'incomplete_scopes', set())
                                incomplete = ((college, day, session) in inc_scopes
                                              or (real_college, day, session) in inc_scopes)
                                extra_notes = []
                                if exempt_marked:
                                    extra_notes.append('含实到栏填“免勤”教室：%s，已折算为0实到计入' % '、'.join(exempt_marked))
                                if missing_rooms:
                                    miss_names = '、'.join(r.room for r in missing_rooms)
                                    extra_notes.append('其中 %d 间教室（%s）未录入实到，已按有效教室计算' % (len(missing_rooms), miss_names))
                                    anomalies.append('%d 间教室实到未录入，已按有效教室折算' % len(missing_rooms))

                                note_prefix = ('（%s）' % '；'.join(extra_notes)) if extra_notes else ''
                                if incomplete:
                                    status = '免勤不完整（待补人数）'
                                    note = ('该学院该次存在无人数的占位免勤申请，免勤人数不完整，'
                                            '以下数值仅供参考、不得以完整名义上报；'
                                            '按公式算得 %.2f%%%s。' % (rate * 100, note_prefix))
                                    anomalies.append('免勤人数不完整，待补人数')
                                elif missing_rooms:
                                    status = '部分教室缺失（已按有效实到计算）'
                                    note = ('该学院该次共 %d 间教室，其中 %d 间未录入实到%s；'
                                            '按有效教室实到合计 %d ÷ (有效教室应到 %d − 有效教室免勤 %d) 算得出勤率 %.2f%%。'
                                            % (len(day_rooms), len(missing_rooms), note_prefix,
                                               actual_sum, valid_expected, valid_exempt, rate * 100))
                                elif rate > 1:
                                    status = '已计算（超过100%，需复核）'
                                    note = ('按公式算得 %.2f%%%s，实到合计 %d 大于考勤基数 %d，'
                                            '需核对免勤与实到的统计口径。'
                                            % (rate * 100, note_prefix, actual_sum, eff_denominator))
                                    anomalies.append('出勤率超过 100%')
                                else:
                                    status = '完整'
                                    note = ('学院出勤率 = 学院总实到 %d ÷ (学院总应到 %s − 当次总免勤 %d) = %.2f%%%s。'
                                            % (actual_sum, expected, exempt, rate * 100, note_prefix))

                    results.append(CollegeRate(
                        sheet=sheet, session=session, college=college, day=day,
                        expected=expected, room_expected_sum=room_expected_sum,
                        actual=actual_sum, exempt=exempt, rate=rate,
                        status=status, note=note, anomalies=anomalies,
                    ))
        return results

    def average_rates(self, rates: List[CollegeRate]) -> Dict[Tuple[str, str], float]:
        """(工作表, 学院) → 各天出勤率的算术平均。

        口径（对话口径，**与上一轮产物一致**）：该学院该表的**所有应统计天**
        都算出出勤率时才给平均值；只要有一天缺失（实到未填、分母非正等），
        平均值就留空，避免用部分天的平均冒充整周平均。

        （另一种做法是「留空天不参与、只对已算出的天取平均」——那是口径修正，
        会与上一轮产物不一致，需要先确认再改。此处保留对话口径。）
        """
        buckets: Dict[Tuple[str, str], List[float]] = defaultdict(list)
        expected_days: Dict[Tuple[str, str], int] = {}
        for sheet, rate_cols in self.detail.rate_cols_by_sheet.items():
            days = [d for d in rate_cols if d != '平均']
            for block in self.detail.blocks_of(sheet):
                expected_days[(sheet, block.name)] = len(days)
        for r in rates:
            key = (r.sheet, r.college)
            if r.rate is not None:
                buckets[key].append(r.rate)
        out = {}
        for key, values in buckets.items():
            if len(values) == expected_days.get(key, len(values)):
                out[key] = sum(values) / len(values)
        return out
