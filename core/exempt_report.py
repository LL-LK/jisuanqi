"""免勤分项统计及独立的汇总工作簿。"""

from collections import Counter

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from .export import ALLOWED_SHEETS, atomic_output, check_output_path
from .parsing import WEEKDAY_NAMES, parse_session


REPORT_COLUMNS = (
    '学院', '星期', '自习', '个人免勤', '集体免勤', '免勤合计',
    '已回填', '容量超额', '回填差额', '免勤状态', '模板出勤率列', '回填状态',
)
COLLEGE_REPORT_COLUMNS = tuple(column for column in REPORT_COLUMNS if column != '容量超额')
REPORT_WIDTHS = dict(zip(REPORT_COLUMNS, (28, 11, 14, 14, 14, 14, 14, 14, 14, 21, 20, 25)))


def build_exempt_rows(calc, room_exempt, rates, has_exempt_file, college_exempt=None,
                      *, summary_scopes=None):
    """统计已解析的申请以及明细表中可回填或可算率的范围。"""
    agg = calc.agg
    input_scopes = (agg.scope_keys() | agg.incomplete_scopes
                    | agg.identity_review_scopes | agg.unknown_college_scopes)
    rate_scopes = {
        (calc.exempt_college(rate.college), rate.day, rate.session)
        for rate in rates if rate.sheet in ALLOWED_SHEETS
    }
    written = Counter()
    if summary_scopes is not None:
        for scope in summary_scopes:
            written[scope] = agg.exempt_count(*scope)
    elif college_exempt is not None:
        for (sheet, college, day), value in college_exempt.items():
            written[(calc.exempt_college(college), day, parse_session(sheet) or '')] += value
    else:
        for room in calc.detail.rooms:
            if room.sheet not in ALLOWED_SHEETS or not room.college_name:
                continue
            session = parse_session(room.sheet) or ''
            college = calc.exempt_college(room.college_name)
            for day, column in room.exempt_cols.items():
                if column is not None:
                    written[(college, day, session)] += room_exempt.get((room.sheet, room.row, day), 0)
    overflow = Counter()
    for sheet, college, day, _room, excess in calc.overflow:
        if sheet in ALLOWED_SHEETS:
            overflow[(calc.exempt_college(college), day, parse_session(sheet) or '')] += excess

    def scope_order(scope):
        college, day, session = scope
        return (college, WEEKDAY_NAMES.index(day) if day in WEEKDAY_NAMES else len(WEEKDAY_NAMES),
                session)

    rows = []
    for scope in sorted(input_scopes | rate_scopes | written.keys(), key=scope_order):
        college, day, session = scope
        personal, group = agg.exempt_breakdown(*scope)
        total = personal + group
        difference = total - written[scope] - overflow[scope]
        if not has_exempt_file:
            status = '未导入免勤表'
        elif scope in agg.unknown_college_scopes:
            status = '学院待核对'
        elif scope in agg.identity_review_scopes:
            status = '身份待核对'
        elif scope in agg.incomplete_scopes:
            status = '待补人数'
        else:
            status = '已统计'

        if not session:
            write_status = '自习时段未识别'
        elif difference < 0:
            write_status = '回填超计，需复核'
        elif scope not in written:
            write_status = ('模板无对应免勤格' if college_exempt is not None or summary_scopes is not None
                            else '无可回填教室')
        elif difference > 0:
            write_status = '部分未落表，需复核'
        elif overflow[scope]:
            write_status = '有容量超额'
        else:
            write_status = '已回填'

        rows.append(dict(zip(REPORT_COLUMNS, (
            college or '未识别学院', day, session or '未识别', personal, group, total,
            written[scope], overflow[scope], difference, status,
            '有' if scope in rate_scopes else '无', write_status,
        ))))
    return rows


def export_exempt_report(path, result, rows=None, extra_protected_paths=()):
    """另存免勤汇总，导出范围默认为本次计算的全部行。"""
    if not result.ok:
        raise ValueError('请先完成计算')
    check_output_path(path, (*result.input_paths, result.output_path, result.summary_output_path,
                             *extra_protected_paths))
    rows = result.exempt_rows if rows is None else rows
    if not rows:
        raise ValueError('当前没有可导出的免勤统计行')

    book = Workbook()
    book.remove(book.active)
    try:
        append_exempt_report(book, rows, result.college_exempt, result.summary)
        with atomic_output(path) as temporary:
            book.save(temporary)
    finally:
        book.close()


def append_exempt_report(book, rows, college_exempt, summary, issues=()):
    report_name = '学院免勤统计' if college_exempt else '免勤统计'
    sheet = book.create_sheet(report_name)
    columns = COLLEGE_REPORT_COLUMNS if college_exempt else REPORT_COLUMNS
    sheet.append(columns)
    for row in rows:
        sheet.append(['未匹配学院' if column == '学院' and row['免勤状态'] == '学院待核对'
                      else row[column] for column in columns])
        for cell in sheet[sheet.max_row]:
            # Excel 将以等号开头的字符串当作公式，导出的来源标签应始终是文本。
            if cell.data_type == 'f':
                cell.data_type = 's'
        if row['免勤状态'] != '已统计' or row['回填差额']:
            for cell in sheet[sheet.max_row]:
                cell.fill = PatternFill('solid', fgColor='FFFFF3CD')
    sheet.freeze_panes = 'D2'
    sheet.auto_filter.ref = sheet.dimensions
    for index, column in enumerate(columns, 1):
        sheet.column_dimensions[get_column_letter(index)].width = REPORT_WIDTHS[column]

    notes = book.create_sheet('汇总与口径')
    notes.append(['项目', '数值或说明'])
    notes.append(['统计周次', summary['统计周次']])
    notes.append(['考勤时段', summary.get('考勤时段', '按明细表的早晚自习列')])
    notes.append(['免勤回填口径', summary.get('免勤回填口径', '')])
    notes.append(['导入免勤表数', summary.get('导入免勤表数', 0)])
    notes.append(['导出组合数', len(rows)])
    for column in columns[3:columns.index('回填差额')]:
        notes.append([column + '（人次）', sum(row[column] for row in rows)])
    notes.append(['未落表（人次）', sum(max(row['回填差额'], 0) for row in rows)])
    notes.append(['回填超计（人次）', sum(max(-row['回填差额'], 0) for row in rows)])
    notes.append(['待补人数（组合）', sum(row['免勤状态'] == '待补人数' for row in rows)])
    notes.append(['身份或学院待核对（组合）',
                  sum(row['免勤状态'] in ('身份待核对', '学院待核对') for row in rows)])
    notes.append(['统计范围', '每行对应一个学院、星期和自习。跨时段求和为人次。'])
    notes.append(['个人免勤', '学号去重后，剔除同班、同日、同自习的集体申请覆盖。'])
    notes.append(['集体免勤', '同班、同日、同自习取有效申报人数。缺人数的集体申请标记待补。'])
    notes.append(['身份核对', '同名但学号不同分别计入。同一时段姓名冲突或同学号跨学院时，相关人次暂缓计入。'])
    calculation = '免勤合计减去已回填' if college_exempt else '免勤合计减去已回填和容量超额'
    notes.append(['回填差额', calculation + '。正数为未落表，负数为回填超计。'])
    notes.append(['模板出勤率列', '显示该组合是否对应模板出勤率列。实际出勤率仍取决于实到数据等条件。'])
    notes.freeze_panes = 'A2'
    notes.column_dimensions['A'].width = 24
    notes.column_dimensions['B'].width = 92

    formatted = [sheet, notes]
    review_issues = [item for item in issues if item.get('review_ref')]
    if review_issues:
        review = book.create_sheet('免勤核对')
        review.append(['问题', '来源', '处理说明'])
        for item in review_issues:
            review.append([item['type'], item['review_ref'], item['note']])
            for cell in review[review.max_row]:
                if cell.data_type == 'f':
                    cell.data_type = 's'
        for column, width in zip('ABC', (26, 65, 82)):
            review.column_dimensions[column].width = width
        review.freeze_panes = 'A2'
        review.auto_filter.ref = review.dimensions
        formatted.append(review)
    for ws in formatted:
        ws.row_dimensions[1].height = 24
        for cell in ws[1]:
            cell.font = Font(bold=True, color='FFFFFFFF')
            cell.fill = PatternFill('solid', fgColor='FF244E9D')
