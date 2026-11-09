# -*- coding: utf-8 -*-
"""
内置标准空白汇总表（早晚自习执行情况终稿）模板生成与填充引擎。
1:1 复刻桂林理工大学官方 15 学院排版、字体、列宽、三线边框与百分比格式。
"""
from __future__ import annotations

import os
from typing import Dict, List, Optional, Tuple, Any
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# 官方标准 15 学院顺序与名称
STANDARD_COLLEGES: Tuple[str, ...] = (
    "地球科学学院",
    "环境科学与工程学院",
    "材料科学与工程学院",
    "化学与生物工程学院",
    "土木工程学院",
    "测绘地理信息学院",
    "计算机科学与工程学院",
    "机械与控制工程学院",
    "物理与电子信息工程学院",
    "商学院",
    "旅游与风景园林学院",
    "艺术学院",
    "外国语学院",
    "数学与统计学院",
    "亚太信息技术学院",
)


def _safe_font(name: str, size: float, bold: bool = False, color: str = "000000") -> Font:
    return Font(name=name, size=size, bold=bold, color=color)


def create_blank_summary_workbook(
    week_label: str = "第九周",
    attendance_dates: str = "",
) -> openpyxl.Workbook:
    """生成 1:1 官方格式的空白汇总表工作簿。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet2"

    # 视图设置：默认显示网格线
    ws.views.sheetView[0].showGridLines = True

    # 列宽设置
    ws.column_dimensions["A"].width = 8.0
    ws.column_dimensions["B"].width = 30.0
    ws.column_dimensions["C"].width = 18.0
    ws.column_dimensions["D"].width = 24.0
    ws.column_dimensions["E"].width = 58.0

    # 边框定义
    thin_side = Side(style="thin", color="000000")
    cell_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # 对齐
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left_top = Alignment(horizontal="left", vertical="top", wrap_text=True)

    # 字体定义
    font_title = _safe_font("方正小标宋简体", 20.0, bold=True)
    font_header = _safe_font("方正仿宋_GB2312", 14.0, bold=False)
    font_data = _safe_font("方正仿宋_GB2312", 14.0, bold=False)
    font_note = _safe_font("仿宋", 14.0, bold=False)

    # 行高
    ws.row_dimensions[1].height = 42.0
    ws.row_dimensions[2].height = 28.0
    for r in range(3, 18):
        ws.row_dimensions[r].height = 26.0
    ws.row_dimensions[18].height = 90.0

    # 1. 标题行 A1:E1
    ws.merge_cells("A1:E1")
    cell_a1 = ws["A1"]
    cell_a1.value = f"桂林理工大学2026年秋季学期{week_label}早晚自习执行情况终稿"
    cell_a1.font = font_title
    cell_a1.alignment = align_center

    # 2. 表头行
    headers = ["序号 ", "学院", "学生出勤率", "值班辅导员巡查率", "执行效果"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = font_header
        cell.alignment = align_center
        cell.border = cell_border

    # 3. 15 学院数据行
    for idx, college_name in enumerate(STANDARD_COLLEGES, 1):
        row_num = idx + 2
        # 序号
        c_num = ws.cell(row=row_num, column=1, value=idx)
        c_num.font = font_data
        c_num.alignment = align_center
        c_num.border = cell_border

        # 学院
        c_col = ws.cell(row=row_num, column=2, value=college_name)
        c_col.font = font_data
        c_col.alignment = align_center
        c_col.border = cell_border

        # 学生出勤率 (默认空白或占位)
        c_rate = ws.cell(row=row_num, column=3)
        c_rate.font = font_data
        c_rate.alignment = align_center
        c_rate.border = cell_border
        c_rate.number_format = "0.00%"

        # 值班辅导员巡查率
        c_patrol = ws.cell(row=row_num, column=4)
        c_patrol.font = font_data
        c_patrol.alignment = align_center
        c_patrol.border = cell_border
        c_patrol.number_format = "0.00%"

        # 执行效果
        c_eff = ws.cell(row=row_num, column=5, value="正常")
        c_eff.font = font_data
        c_eff.alignment = align_center
        c_eff.border = cell_border

    # 4. 备注行 A18:E18
    ws.merge_cells("A18:E18")
    cell_note = ws["A18"]
    date_part = f"\n    考勤时间：{attendance_dates}" if attendance_dates else ""
    note_text = (
        f"备注：{date_part}\n"
        "    学生出勤率=学院实际出勤人次/学院应出勤人次\n"
        "    值班辅导员巡查率=学院值班辅导员实际签到次数/学院值班辅导员应签到次数"
    )
    cell_note.value = note_text
    cell_note.font = font_note
    cell_note.alignment = align_left_top

    # 创建辅助空表 Sheet3（对齐原版）
    wb.create_sheet("Sheet3")
    return wb


def fill_summary_workbook(
    wb: openpyxl.Workbook,
    college_rates: Dict[str, float],
    college_patrol_rates: Optional[Dict[str, float]] = None,
    college_anomalies: Optional[Dict[str, str]] = None,
    week_label: str = "第九周",
    attendance_dates: str = "",
) -> openpyxl.Workbook:
    """将计算好的出勤率、巡查率与异常说明填入汇总工作簿。"""
    ws = wb["Sheet2"] if "Sheet2" in wb.sheetnames else wb.active

    # 更新标题
    ws["A1"].value = f"桂林理工大学2026年秋季学期{week_label}早晚自习执行情况终稿"

    # 如果有新的考勤时间说明
    if attendance_dates:
        cell_note = ws["A18"]
        cell_note.value = (
            f"备注：\n    考勤时间：{attendance_dates}\n"
            "    学生出勤率=学院实际出勤人次/学院应出勤人次\n"
            "    值班辅导员巡查率=学院值班辅导员实际签到次数/学院值班辅导员应签到次数"
        )

    # 遍历填入
    for r in range(3, 18):
        c_name = ws.cell(r, 2).value
        if not c_name:
            continue
        c_name = str(c_name).strip()

        # 1. 学生出勤率
        rate = college_rates.get(c_name)
        if rate is not None:
            cell_rate = ws.cell(r, 3, value=round(rate, 4))
            cell_rate.number_format = "0.00%"

        # 2. 辅导员巡查率
        if college_patrol_rates and c_name in college_patrol_rates:
            p_rate = college_patrol_rates[c_name]
            cell_patrol = ws.cell(r, 4, value=round(p_rate, 4))
            cell_patrol.number_format = "0.00%"
        elif ws.cell(r, 4).value is None:
            # 默认 100%
            ws.cell(r, 4, value=1.0).number_format = "0.00%"

        # 3. 执行效果
        if college_anomalies and c_name in college_anomalies and college_anomalies[c_name]:
            ws.cell(r, 5, value=college_anomalies[c_name])
        elif ws.cell(r, 5).value is None or ws.cell(r, 5).value == "":
            ws.cell(r, 5, value="正常")

    return wb


def generate_summary_file(
    output_path: str,
    college_rates: Dict[str, float],
    college_patrol_rates: Optional[Dict[str, float]] = None,
    college_anomalies: Optional[Dict[str, str]] = None,
    week_label: str = "第九周",
    attendance_dates: str = "",
) -> str:
    """直接生成并保存规范终稿汇总表 Excel 文件。"""
    wb = create_blank_summary_workbook(week_label=week_label, attendance_dates=attendance_dates)
    wb = fill_summary_workbook(
        wb,
        college_rates=college_rates,
        college_patrol_rates=college_patrol_rates,
        college_anomalies=college_anomalies,
        week_label=week_label,
        attendance_dates=attendance_dates,
    )
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    wb.close()
    return output_path
