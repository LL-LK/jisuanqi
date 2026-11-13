# -*- coding: utf-8 -*-
"""编排层：把导入 → 解析 → 去重 → 分摊 → 计算 → 导出串成一条流水线。

桌面界面、命令行、测试与合成数据自测入口都只调用 run()，保证行为一致。
"""

from __future__ import annotations

import os
import time
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .calc import Calculator, estimate_source
from .attendance_summary import export_attendance_summary, load_summary_template
from .dictionary import Dictionary
from .exempt_report import append_exempt_report, build_exempt_rows
from .export import (ALLOWED_SHEETS, TAMPER_META_SHEET, ExportResult, _same_path, check_output_path,
                     export, unique_workbook_paths, verify_workbook_integrity)
from .io_detail import load_detail
from .io_exempt import aggregate_exemptions, load_exemptions
from .model import CollegeRate, CollegeWeeklyRate
from .parsing import WEEKDAY_NAMES, normalize_campus, parse_session
from .summary_template import attendance_period, generate_summary_file

DEFAULT_OUTPUT_DIR = '考勤统计输出'


@dataclass
class RunConfig:
    detail_path: str
    exempt_path: Optional[str] = None
    output_path: Optional[str] = None
    week_no: int = 7
    exempt_source: str = '未分类'       # 全免勤大表 / 在线存档 / 未分类
    campus: str = '雁山'
    dictionary_path: Optional[str] = None
    complete_mode: Optional[bool] = None  # None: 按词典参数决定; True: 强制容错完整计算; False: 强制严格契约
    match_classes: bool = False         # 未带教室号默认按已验收的比例分摊口径处理
    exempt_paths: tuple[str, ...] = ()
    college_exempt: Optional[bool] = None  # 未指定时按模板合并结构判断
    include_exempt_summary: bool = False
    db_path: Optional[str] = None          # [M1 新增] 指定数据库路径 (None 表示采用默认路径)
    enable_storage: bool = True            # [M1 新增] 历史记录入库开关 (默认开启)
    summary_output_path: Optional[str] = None  # 未指定时由主结果文件名派生终稿路径
    # 注：本口径「免勤不封顶」，因此不提供任何封顶开关——
    # 留一个可切到被否口径的旋钮，比没有旋钮更危险。

    @property
    def exemption_files(self):
        return unique_workbook_paths((self.exempt_path, *self.exempt_paths))


@dataclass
class RunResult:
    ok: bool = False
    error: str = ''
    output_path: Optional[str] = None
    summary_output_path: Optional[str] = None
    summary: Dict = field(default_factory=dict)
    college_rates: List[CollegeRate] = field(default_factory=list)
    room_rows: List[Dict] = field(default_factory=list)
    room_students: Dict = field(default_factory=dict)       # (sheet, row, day) -> list[dict]
    student_details: Dict = field(default_factory=dict)     # (college, day, session) -> list[dict]
    exempt_overflow: List[Dict] = field(default_factory=list)
    anomalies: List[Dict] = field(default_factory=list)
    exempt_stats: Dict = field(default_factory=dict)
    sheets: List[str] = field(default_factory=list)
    exempt_rows: List[Dict] = field(default_factory=list)
    input_paths: tuple[str, ...] = ()
    weekly_rates: List[CollegeWeeklyRate] = field(default_factory=list)
    college_exempt: bool = False
    summary_template_rates: List[Dict] = field(default_factory=list)
    college_names: Dict[str, str] = field(default_factory=dict)
    storage_run_id: Optional[int] = None
    phase_durations: Dict[str, int] = field(default_factory=dict)
    duration_ms: int = 0


def _anomaly(kind: str, level: str, scope: str, note: str) -> Dict:
    return {'type': kind, 'level': level, 'scope': scope, 'note': note}


def output_path_for(config: RunConfig) -> str:
    if config.output_path:
        return config.output_path
    base = os.path.splitext(os.path.basename(config.detail_path))[0]
    out_dir = os.path.join(os.path.dirname(config.detail_path) or '.', DEFAULT_OUTPUT_DIR)
    return os.path.join(out_dir, '%s_第%d周计算结果.xlsx' % (base, config.week_no))


def summary_path_for(config: RunConfig) -> str:
    return config.summary_output_path or os.path.splitext(output_path_for(config))[0] + '_汇总终稿.xlsx'


def run(config: RunConfig) -> RunResult:
    """执行一次完整统计。"""
    start_time = time.perf_counter()
    result = RunResult()
    try:
        if not config.detail_path or not os.path.exists(config.detail_path):
            raise FileNotFoundError('找不到自习执行情况明细表：%s' % config.detail_path)

        campus = normalize_campus(config.campus)
        if campus != '雁山':
            raise ValueError('当前考勤模板只计算雁山校区，不能将其他校区免勤扣入雁山应到人数')

        exempt_paths = config.exemption_files
        result.input_paths = tuple(os.path.abspath(path) for path in
                                   (config.detail_path, *exempt_paths, config.dictionary_path)
                                   if path)
        output_path = output_path_for(config)
        if _same_path(config.detail_path, output_path):
            raise ValueError('输出路径不能与输入明细表相同，以免覆盖原始文件')
        if any(_same_path(path, output_path) for path in exempt_paths):
            raise ValueError('输出路径不能与任一输入免勤表相同，以免覆盖原始文件')
        if config.dictionary_path and _same_path(config.dictionary_path, output_path):
            raise ValueError('输出路径不能与免勤词典相同，以免覆盖词典')
        check_output_path(output_path, ())

        detail = load_detail(config.detail_path)
        if TAMPER_META_SHEET in detail.sheet_names:
            valid, violations = verify_workbook_integrity(config.detail_path)
            if not valid:
                raise ValueError('输入考勤表完整性校验失败：' + '、'.join(violations[:5]))
        t_preflight_end = time.perf_counter()
        preflight_ms = int((t_preflight_end - start_time) * 1000)

        t_parse_start = t_preflight_end
        attendance = None
        if not any(detail.blocks_of(sheet) for sheet in ALLOWED_SHEETS[:2]):
            attendance = load_summary_template(config.detail_path)
            if not attendance:
                raise ValueError('未找到可统计的雁山明细表或学院出勤汇总表')
            if attendance.week and attendance.week != config.week_no:
                raise ValueError(f'汇总表标注第 {attendance.week} 周，请核对统计周次')
        result.college_exempt = bool(attendance or detail.college_exempt_layout or config.college_exempt)
        if result.college_exempt and not exempt_paths:
            raise ValueError('请导入至少一份免勤表。确无免勤时，可导入只保留表头的免勤表')
        if not attendance and not any(detail.rate_cols_by_sheet.get(sheet) for sheet in ALLOWED_SHEETS[:2]):
            raise ValueError('明细表内没有可识别的出勤率列，请核对表头')
        summary_target = None if attendance else summary_path_for(config)
        if summary_target:
            check_output_path(summary_target, (*result.input_paths, output_path))
        result.sheets = list(detail.sheet_names)
        for warn in detail.warnings:
            result.anomalies.append(_anomaly('明细表结构提示', 'warn', '明细表', warn))
        for sheet, title_week in detail.title_weeks.items():
            if title_week != config.week_no:
                result.anomalies.append(_anomaly(
                    '模板标题周次不一致', 'warn', sheet,
                    f'表内标题为第 {title_week} 周，本次按所选第 {config.week_no} 周统计。'))

        dictionary = Dictionary.load(config.dictionary_path)
        identity_check = bool(result.college_exempt or exempt_paths)
        target_names = ([college for _, college, _ in attendance.rows] if attendance else
                        list(dict.fromkeys(block.name for block in detail.blocks if block.name)))
        result.college_names = {name: dictionary.exact_college(name) for name in target_names}
        if attendance and len(set(result.college_names.values())) != len(attendance.rows):
            raise ValueError('汇总表有重复学院或学院别名指向同一学院，请核对后重试')

        records = []
        for index, path in enumerate(exempt_paths, 1):
            if not os.path.isfile(path):
                raise FileNotFoundError('找不到免勤人员统计表：%s' % path)
            if _same_path(path, config.detail_path):
                raise ValueError('明细表与免勤表不能是同一个文件')
            try:
                records.extend(load_exemptions(path, source=config.exempt_source,
                                               week_no=config.week_no, dictionary=dictionary,
                                               require_schema=True, source_number=index,
                                               keep_unassigned=identity_check))
            except ValueError as exc:
                raise ValueError(f'{os.path.basename(path)}：{exc}') from exc
        if identity_check:
            for record in records:
                record.college = dictionary.exact_college(record.college)
        agg = aggregate_exemptions(
            records, config.week_no, campus=campus, identity_check=identity_check,
            attendance_scopes=set(attendance.scopes) if attendance else None,
            allowed_colleges=set(result.college_names.values()) if identity_check else None)

        result.anomalies.extend(agg.anomalies)

        t_parse_end = time.perf_counter()
        parse_ms = int((t_parse_end - t_parse_start) * 1000)

        t_calc_start = t_parse_end
        complete_mode = getattr(config, 'complete_mode', None)
        if complete_mode is None:
            complete_mode = dictionary.params.get('出勤率完整计算模式', '0') == '1'
        calc = Calculator(detail, agg, alias=dictionary.college_alias,
                          room_level=getattr(agg, 'room_level', None),
                          complete_mode=complete_mode, match_classes=config.match_classes)
        if not identity_check:
            result.college_names = {name: calc.exempt_college(name) for name in target_names}
        room_exempt = {} if result.college_exempt else calc.room_exemptions()
        college_exempt = _college_writes(detail, calc) if result.college_exempt and not attendance else None
        # C1=甲：结构化超额头（学院/星期/自习/教室/超额人数），供断言与行级归因使用
        result.exempt_overflow = [
            {'sheet': s, 'college': c, 'day': d, 'room': rm, 'excess': e}
            for (s, c, d, rm, e) in getattr(calc, 'overflow', [])]
        rates = calc.college_rates()
        for rate in rates:
            scope = (calc.exempt_college(rate.college), rate.day, rate.session)
            if scope in agg.rule_review_scopes:
                if rate.status == '完整':
                    rate.status = '免勤规则待核对'
                rate.note += ' 该时段有校区、集体人数或考勤时段待核对。'
            if scope in agg.identity_review_scopes:
                if rate.status == '完整':
                    rate.status = '已计算（身份争议按1人计）'
                rate.note += (' 该时段有同号异名/跨学院争议，已按同一学号·同一天·同一自习去重计 1 人；'
                              '出勤率照常计算，请核对「免勤核对」来源行。')
        averages = calc.average_rates(rates)
        result.weekly_rates = _weekly_rates(rates, averages)
        summary_scopes = ({(college, day, session) for college in result.college_names.values()
                           for day, session in attendance.scopes} if attendance else None)
        result.exempt_rows = build_exempt_rows(calc, room_exempt, rates, bool(exempt_paths),
                                               college_exempt, summary_scopes=summary_scopes)

        t_calc_end = time.perf_counter()
        calc_ms = int((t_calc_end - t_calc_start) * 1000)

        t_export_start = t_calc_end
        mode = ('学院汇总' if result.college_exempt else
                '班级推定后分摊' if config.match_classes else '教室号匹配后比例分摊')
        report_info = {'统计周次': config.week_no, '免勤回填口径': mode,
                       '导入免勤表数': len(exempt_paths),
                       '空白校区默认值': dictionary.params.get('空白校区默认值', '雁山'),
                       '按默认校区参与筛选的记录': agg.campus_defaulted_records}
        def append_report(book):
            append_exempt_report(book, result.exempt_rows, result.college_exempt, report_info,
                                 agg.anomalies if identity_check else ())
        if attendance:
            report_info['考勤时段'] = attendance.period
            totals = _summary_totals(attendance, agg, result.college_names)
            result.summary_template_rates = [
                {'学院': college, '原表学生出勤率': rate, '免勤合计': sum(totals[college][:2])}
                for _, college, rate in attendance.rows]
            export_attendance_summary(config.detail_path, output_path, attendance, totals, append_report)
            export_result = ExportResult()
            export_result.output_path = output_path
            export_result.written_exempt_cells = 3 * len(totals)
        else:
            export_result = export(
                config.detail_path, output_path, detail, room_exempt, rates, averages,
                estimated_cells=calc.estimated_cells, college_exempt=college_exempt,
                append_report=append_report if config.include_exempt_summary else None)
        result.output_path = export_result.output_path

        for name in export_result.removed_sheets:
            result.anomalies.append(_anomaly('已删除多余工作表', 'info', name,
                                             '按规则只保留三张原始工作表'))

        # ---- 异常清单 ----
        for sheet, college, day, room, excess in getattr(calc, 'overflow', []):
            result.anomalies.append(_anomaly(
                '教室免勤人数超过该教室应到人数', 'warn',
                '%s / %s / %s / 教室 %s' % (sheet, college, day, room),
                '该教室分摊值按应到人数封顶，超出 %d 人未再分摊；'
                '学院总免勤仍按唯一学号去重值计算，出勤率分母为负时不产出出勤率' % excess))
        for college, day, session in sorted(getattr(agg, 'incomplete_scopes', ())):
            result.anomalies.append(_anomaly(
                '免勤不完整（待补人数）', 'warn',
                '%s / %s / %s' % (college, day, session),
                '该学院当次存在「全班」这类无人数的占位申请，免勤人数不完整，'
                '当前数值不得宣称完整，请人工补齐人数后重算'))
        # 有免勤列、但出勤率区没有这一列的（如晚自习表的周五）：
        # 免勤人数照常写入表格，但不能静默丢弃，必须在异常清单里说明。
        for sheet, rate_cols in detail.rate_cols_by_sheet.items():
            exempt_days = set()
            for room in detail.rooms_of(sheet):
                exempt_days.update(room.exempt_cols)
            for day in sorted(exempt_days - set(rate_cols)):
                result.anomalies.append(_anomaly(
                    '该天免勤已填表但不参与出勤率计算', 'info',
                    '%s / %s' % (sheet, day),
                    '该工作表的出勤率区没有「%s」列，这一天的免勤人数已写入「%s免勤人数」列，'
                    '但不参与出勤率计算，也不计入学院汇总' % (day, day)))

        # 有免勤记录、但表里根本没有这一列的：无处落表，只能进异常清单。
        days_by_session = defaultdict(set)
        for college, day, session in agg.scope_keys():
            days_by_session[session or ''].add(day)
        for sheet in detail.rate_cols_by_sheet:
            session = parse_session(sheet) or ''
            have = set()
            for room in detail.rooms_of(sheet):
                have.update(room.exempt_cols)
            for day in sorted(days_by_session.get(session, set()) - have):
                result.anomalies.append(_anomaly(
                    '该天免勤无处落表', 'warn', '%s / %s' % (sheet, day),
                    '免勤表里有这一天的申请，但明细表没有「%s免勤人数」列，无法落表；'
                    '这些免勤保留在学院免勤统计中，请人工确认' % day))
        for rec in rates:
            scope = '%s / %s / %s' % (rec.sheet, rec.college, rec.day)
            if rec.status != '完整':
                result.anomalies.append(_anomaly(rec.status,
                                                 'warn' if rec.rate is None else 'info',
                                                 scope, rec.note))
            for extra in rec.anomalies:
                result.anomalies.append(_anomaly(extra, 'warn', scope, rec.note))
        for row in result.exempt_rows:
            if row['回填差额'] != 0:
                result.anomalies.append(_anomaly(
                    '免勤回填差额', 'warn',
                    ' / '.join(row[key] for key in ('学院', '星期', '自习')),
                    '%s。免勤合计 %d − 已回填 %d − 容量超额 %d = %d 人次。' % (
                        row['回填状态'], row['免勤合计'], row['已回填'],
                        row['容量超额'], row['回填差额'])))

        if summary_target:
            try:
                # 容错模式的估算结果仅留在明细，终稿始终采用严格口径。
                final_rates = rates if not complete_mode else Calculator(
                    detail, agg, alias=dictionary.college_alias).college_rates()
                _write_summary(config, summary_target, detail, calc, final_rates)
                result.summary_output_path = summary_target
            except (OSError, ValueError) as exc:
                result.anomalies.append(_anomaly('终稿导出失败', 'warn', '汇总终稿', str(exc)))
        t_export_end = time.perf_counter()
        export_ms = int((t_export_end - t_export_start) * 1000)

        # ---- 汇总 ----
        computed = [r for r in rates if r.rate is not None]
        over = [r for r in computed if r.rate > 1]
        result.college_rates = rates
        result.exempt_stats = {
            '总记录数': agg.total_records,
            '有效记录数': agg.total_records - (
                agg.invalid_records + agg.week_filtered + agg.campus_filtered
                + agg.unmatched_week + agg.unmatched_day + agg.incomplete_records
                + agg.outside_attendance_records + agg.unknown_college_records),
            '因周次不符过滤': agg.week_filtered,
            '因校区过滤': agg.campus_filtered,
            '空白校区按词典默认值计入筛选的记录': agg.campus_defaulted_records,
            '唯一学号数': agg.unique_sids,
            '重复申请事件': agg.duplicate_events,
            '集体免勤申请': agg.group_applications,
            '被集体申请覆盖的个人记录': agg.covered_personal,
            '带教室号记录': agg.room_matched,
            '无教室号记录': agg.roomless,
            '周次无法识别': agg.unmatched_week,
            '星期无法识别': agg.unmatched_day,
            '无法定位的记录': agg.invalid_records,
            '学号位数异常记录': agg.odd_sid_records,
            '占位学号待补人数记录': agg.incomplete_records,
            '疑似同一人重复登记对': agg.same_person_pairs,
            '跨学院重复申请事件': agg.cross_college_conflicts,
            '不在汇总考勤时段的记录': agg.outside_attendance_records,
            '学院待核对记录': agg.unknown_college_records,
            '按班级核对学院记录': agg.college_matched_by_class,
            '按姓名学院班级核对学号': agg.name_matched_records,
            '同学号身份冲突': agg.identity_conflicts,
            '身份冲突事件（已按1人计）': getattr(agg, 'identity_conflict_events', 0),
            # 兼容旧报告字段：新口径不再暂缓计入，固定为 0。
            '身份冲突暂缓计入人次': agg.held_identity_events,
        }
        result.summary = {
            '工作表': ' / '.join(result.sheets),
            '教室数': len(detail.rooms),
            '学院数': len(target_names),
            **report_info,
            '免勤记录总数': agg.total_records,
            '集体免勤申请': agg.group_applications,
            '重复申请事件': agg.duplicate_events,
            '可计算出勤率的学院×天组合': '%d / %d' % (len(computed), len(rates)),
            '超过 100% 的组合': len(over),
            '异常条目': len(result.anomalies),
            '写出免勤单元格': export_result.written_exempt_cells,
            '写出出勤率单元格': export_result.written_rate_cells,
            '合并区归一化次数': export_result.merged_normalized,
        }
        result.room_rows = _room_rows(detail, room_exempt,
                                       calc.estimated_cells, result.college_exempt)
        result.room_students = getattr(calc, 'room_students_map', {})
        result.student_details = getattr(agg, 'student_map', {})
        result.ok = True

        duration_ms = int((time.perf_counter() - start_time) * 1000)
        result.duration_ms = duration_ms
        result.phase_durations = {
            'preflight_ms': preflight_ms,
            'parse_ms': parse_ms,
            'calc_ms': calc_ms,
            'export_ms': export_ms,
        }

        # ---- M1 数据持久化入库 Hook ----
        if getattr(config, 'enable_storage', True):
            try:
                from .storage import AttendanceStorage
                storage = AttendanceStorage(db_path=getattr(config, 'db_path', None))
                run_id = storage.record_run(config, result, duration_ms=duration_ms)
                result.storage_run_id = run_id
                result.summary['run_id'] = run_id
            except Exception as exc:
                result.anomalies.append(_anomaly(
                    '计算记录存储异常', 'warn', '本地数据库',
                    f'自动保存至本地历史数据库失败: {exc}（结果文件已正常生成，不受影响）'
                ))
    except Exception as exc:                      # noqa: BLE001 - 界面需要展示原因
        result.ok = False
        result.error = '%s: %s' % (type(exc).__name__, exc)
        result.duration_ms = int((time.perf_counter() - start_time) * 1000)
    return result


def _write_summary(config, path, detail, calc, rates):
    scopes = tuple(
        (day, session) for day in WEEKDAY_NAMES for sheet in ALLOWED_SHEETS[:2]
        if day in detail.rate_cols_by_sheet.get(sheet, {})
        if (session := parse_session(sheet)))
    by_college = defaultdict(dict)
    for rate in rates:
        college = calc.exempt_college(rate.college)
        by_college[college][(rate.day, rate.session)] = rate

    results, notes = {}, {}
    for college, by_scope in by_college.items():
        daily = [by_scope.get(scope) for scope in scopes]
        if not all(rate is not None and rate.rate is not None for rate in daily):
            notes[college] = '待复核：应统计时段缺实到或有效应到人数'
            continue
        identity_pending = any(
            (college, day, session) in calc.agg.identity_review_scopes
            for day, session in scopes)
        review_pending = any(
            (college, day, session) in calc.agg.incomplete_scopes
            or (college, day, session) in calc.agg.rule_review_scopes
            for day, session in scopes)
        denominator = sum(rate.expected - rate.exempt for rate in daily)
        results[college] = sum(rate.actual for rate in daily) / denominator
        flags = []
        if identity_pending:
            flags.append('身份争议已按1人计')
        if review_pending:
            flags.append('免勤人数/适用范围待确认')
        if any(rate.rate > 1 for rate in daily):
            flags.append('有时段出勤率超过100%')
        if flags:
            notes[college] = '；'.join(flags) + '；请核对「免勤核对」来源行'
        else:
            notes[college] = '待填写巡查与执行情况'

    term, label, period = attendance_period(config.detail_path, scopes)
    generate_summary_file(path, results, college_anomalies=notes,
                          week_label=f'第{config.week_no}周',
                          attendance_dates=period, period_label=label, academic_term=term)


def _college_writes(detail, calc):
    values = {}
    seen = set()
    for block in detail.blocks:
        sheet, college = block.sheet, block.name
        scope = (sheet, calc.exempt_college(college))
        if sheet not in ALLOWED_SHEETS[:2] or not block.rows or not college or scope in seen:
            continue
        seen.add(scope)
        for day in detail.exempt_cols_by_sheet[sheet]:
            values[(sheet, college, day)] = calc.exempt_for(college, day, parse_session(sheet) or '')
    return values


def _summary_totals(template, agg, college_names):
    totals = {}
    for _, name, _ in template.rows:
        college = college_names[name]
        early = late = 0
        needs_review = False
        for day, session in template.scopes:
            scope = (college, day, session)
            count = agg.exempt_count(*scope)
            if session == '早自习':
                early += count
            else:
                late += count
            needs_review |= (scope in agg.identity_review_scopes or scope in agg.incomplete_scopes
                             or scope in agg.rule_review_scopes)
        totals[name] = early, late, needs_review
    return totals


def _weekly_rates(rates, averages):
    groups = defaultdict(list)
    for rate in rates:
        groups[(rate.sheet, rate.college)].append(rate)
    return [
        CollegeWeeklyRate(sheet, daily[0].session, college,
                          sum(rate.rate is not None for rate in daily), len(daily),
                          averages.get((sheet, college)),
                          any(rate.status != '完整' for rate in daily))
        for (sheet, college), daily in groups.items()
    ]


def _room_rows(detail, room_exempt, estimated_cells=None, college_exempt=False) -> List[Dict]:
    rows = []
    for room in detail.rooms:
        item = {
            '工作表': room.sheet,
            '行号': room.row,
            '学院': room.college_name,
            '教室': room.room,
            '应到人数': room.expected,
            '分配专业班级': room.classes_text,
            # 行级便利字段：该行只要有任一天是估算即为「估算」
            '免勤来源': '精确',
        }
        if college_exempt:
            item.update({f'{day}实到': actual for day, actual in room.actuals.items()})
            del item['免勤来源']
            rows.append(item)
            continue
        for day, col in sorted(room.exempt_cols.items()):
            item['%s实到' % day] = room.actuals.get(day)
            item['%s免勤' % day] = room_exempt.get((room.sheet, room.row, day), 0)
            # 估算标记位（唯一判定源：core.calc.estimate_source）——只做标记出口，
            # 不参与任何计算，也不改变产出数字
            item['%s免勤来源' % day] = estimate_source(
                estimated_cells, (room.sheet, room.row, day))
            if item['%s免勤来源' % day] == '估算':
                item['免勤来源'] = '估算'
        rows.append(item)
    return rows
