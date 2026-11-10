"""识别已有学院出勤率的汇总模板，并补充对应时段的免勤。"""

import re
from contextlib import closing
from dataclasses import dataclass
from datetime import date

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .export import atomic_output, generate_tamper_metadata
from .io_detail import week_from_label
from .parsing import WEEKDAY_NAMES, normalize_text, parse_college_label, strip_spaces


@dataclass
class SummaryTemplate:
    sheet: str
    header: int
    rows: list
    week: int | None
    scopes: tuple
    period: str


def _attendance_period(sheet):
    texts = [normalize_text(line) for row in sheet.iter_rows(max_col=min(sheet.max_column, 12))
             for cell in row if isinstance(cell.value, str) for line in cell.value.splitlines()]
    note = next((match.group(1).strip() for value in texts
                 if (match := re.search(r'考勤时间\s*[:：]\s*(.*)', value))), None)
    if note is None:
        weekly = next((match.group(1).strip() for value in texts
                       if (match := re.search(r'考勤时段\s*[:：]\s*(.*)', value))), None)
        if weekly is not None:
            scopes = {}
            for part in re.split(r'[,、]+', weekly):
                day, session = part.strip()[:2], part.strip()[2:]
                if day not in WEEKDAY_NAMES or session not in ('早自习', '晚自习'):
                    raise ValueError('无法识别汇总表的考勤时段，请按「周一早自习」格式填写')
                scopes[(day, session)] = None
            return tuple(scopes), weekly
        return tuple((day, session) for day in WEEKDAY_NAMES
                     for session in ('早自习', '晚自习')), '所选周次全部早晚自习'

    year = re.search(r'(20\d{2})\s*年', normalize_text(sheet['A1'].value))
    scopes = {}
    periods = []
    calendar_weeks = set()
    for segment in re.split(r'[,，、;；]+', note):
        match = re.fullmatch(r'\s*(\d{1,2})月\s*(\d{1,2})日\s*(早晚|早|晚)(?:自习)?\s*[。.]?\s*', segment)
        if not match or not year:
            raise ValueError('无法识别汇总表的考勤时间，请按「10月26日早晚自习」格式填写，并在标题注明年份')
        month, day = map(int, match.groups()[:2])
        day_date = date(int(year[1]), month, day)
        calendar_weeks.add(day_date.isocalendar()[:2])
        if len(calendar_weeks) > 1:
            raise ValueError('汇总表考勤时间跨周，请按周分别统计')
        weekday = WEEKDAY_NAMES[day_date.weekday()]
        for period in match[3]:
            scope = (weekday, period + '自习')
            label = f'{month}月{day}日{period}自习'
            if scope in scopes and scopes[scope] != label:
                raise ValueError('汇总表考勤时间跨周，请按周分别统计')
            if scope not in scopes:
                scopes[scope] = label
                periods.append(label)
    return tuple(scopes), '、'.join(periods)


def load_summary_template(path):
    with closing(load_workbook(path, data_only=True)) as book:
        for sheet in book:
            for cells in sheet.iter_rows(max_row=min(sheet.max_row, 8),
                                         max_col=min(sheet.max_column, 40)):
                headings = {strip_spaces(cell.value): cell.column for cell in cells if cell.value}
                college_col = headings.get('学院') or headings.get('学院名称')
                rate_col = next((column for title, column in headings.items()
                                 if title.startswith(('学生出勤率', '学院出勤率'))), None)
                if not college_col or not rate_col or '教室' in headings:
                    continue
                header = cells[0].row
                sequence_col = headings.get('序号')
                rows = []
                for row in range(header + 1, sheet.max_row + 1):
                    college, _ = parse_college_label(sheet.cell(row, college_col).value)
                    if not college or college.startswith(('备注', '合计', '总计')):
                        continue
                    if sequence_col and not str(sheet.cell(row, sequence_col).value).strip().isdigit():
                        continue
                    rows.append((row, college, sheet.cell(row, rate_col).value))
                if rows:
                    scopes, period = _attendance_period(sheet)
                    return SummaryTemplate(sheet.title, header, rows,
                                           week_from_label(sheet['A1'].value), scopes, period)
    return None


def export_attendance_summary(input_path, output_path, template, totals, append_report):
    with closing(load_workbook(input_path)) as book:
        sheet = book[template.sheet]
        labels = ('早自习免勤人次', '晚自习免勤人次', '合计免勤人次', '免勤核对状态')
        old = {strip_spaces(cell.value): cell.column for cell in sheet[template.header]
               if cell.value is not None}
        # 从表头右侧的空列追加，保留人工字段、公式和零散备注。
        next_column = max(old.values(), default=0) + 1
        columns = []
        for label in labels:
            column = old.get(label)
            if not column:
                while any(sheet.cell(row, next_column).value is not None
                          for row in range(template.header, sheet.max_row + 1)):
                    next_column += 1
                column = next_column
                next_column += 1
            columns.append(column)
            cell = sheet.cell(template.header, column, label)
            cell.font = Font(bold=True, color='FFFFFFFF')
            cell.fill = PatternFill('solid', fgColor='FF244E9D')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            sheet.column_dimensions[get_column_letter(column)].width = 18
        for row, college, _ in template.rows:
            early, late, pending = totals[college]
            for column, value in zip(columns, (early, late, early + late, '待核对' if pending else '已核对')):
                cell = sheet.cell(row, column)
                cell.value = value
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.fill = PatternFill('solid', fgColor='FFFFF3CD') if pending else PatternFill()
        append_report(book)
        generate_tamper_metadata(book, sheet_names=(template.sheet,))
        with atomic_output(output_path) as temporary:
            book.save(temporary)
