# -*- coding: utf-8 -*-
"""免勤词典：可编辑的归一化规则表。

词典以 Excel 形式存放（免勤词典.xlsx），界面里可直接编辑并保存。
表结构：
    学院别名      正式名称 | 别名 | 说明
    班级别名      别名 | 规范名称 | 说明
    校区别名      原始写法 | 规范校区 | 说明
    免勤理由分类  理由关键词 | 分类 | 说明
    统计参数      参数 | 值 | 说明
"""

from __future__ import annotations

import os
from typing import Dict, List, Optional, Union

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font

from .parsing import normalize_campus, normalize_class

SHEET_NAMES = ('学院别名', '班级别名', '校区别名', '免勤理由分类', '统计参数')
HEADERS = {
    '学院别名': ('正式名称', '别名', '说明'),
    '班级别名': ('别名', '规范名称', '说明'),
    '校区别名': ('原始写法', '规范校区', '说明'),
    '免勤理由分类': ('理由关键词', '分类', '说明'),
    '统计参数': ('参数', '值', '说明'),
}

DEFAULT_CAMPUS_ALIASES = [
    ('雁山校区', '雁山', '标准写法'),
    ('雁山校园', '雁山', '手写变体'),
    ('雁山区', '雁山', '区划简称'),
    ('燕山校区', '雁山', '同音错字'),
    ('雁三校区', '雁山', '形似错字'),
    ('雁门校区', '雁门', '单独标识，默认不计入雁山'),
    ('雁山县区', '雁山', '错别字'),
    ('雁山校區', '雁山', '繁体字'),
    ('㕍山校区', '雁山', '异体字变体'),
    ('雁山校庆', '雁山', '录入笔误'),
    ('屏风校区', '屏风', '本项目整体排除'),
    ('屏风', '屏风', '本项目整体排除'),
    ('屏峰校区', '屏风', '同音错字'),
]

DEFAULT_COLLEGE_ALIASES = [
    # 官方全称, 免勤表常见别名/简称/错别字, 说明
    ('化学与生物工程学院', '化生学院', '标准简称'),
    ('化学与生物工程学院', '化生院', '口语简称'),
    ('化学与生物工程学院', '化生', '口语极简简称'),
    ('化学与生物工程学院', '化工学院', '历史简称'),
    ('化学与生物工程学院', '化学与工程学院', '常见填报错字'),
    ('化学与生物工程学院', '化学与生物学院', '常见填报漏字'),
    ('化学与生物工程学院', '化学与生物工程', '末尾漏学院二字'),
    ('化学与生物工程学院', '化学工程与工艺学院', '专业误填为学院'),
    ('化学与生物工程学院', '化工与生物工程学院', '别名变体'),
    ('土木工程学院', '土木学院', '标准简称'),
    ('土木工程学院', '土建学院', '历史简称'),
    ('土木工程学院', '土木与建筑工程学院', '历史全称'),
    ('地球科学学院', '地科学院', '标准简称'),
    ('地球科学学院', '地学学院', '简称变体'),
    ('地球科学学院', '地球学院', '常见填报漏字'),
    ('地球科学学院', '地院', '口语极简简称'),
    ('机械与控制工程学院', '机控学院', '标准简称'),
    ('机械与控制工程学院', '机控院', '口语简称'),
    ('机械与控制工程学院', '机电学院', '历史简称'),
    ('机械与控制工程学院', '机械学院', '历史简称'),
    ('机械与控制工程学院', '体现与控制工程学院', '五笔/拼音填报错字'),
    ('机械与控制工程学院', '机械控制与工程学院', '语序颠倒填报错字'),
    ('计算机科学与工程学院', '计算机学院', '标准简称'),
    ('计算机科学与工程学院', '计科院', '标准简称'),
    ('计算机科学与工程学院', '计网学院', '常见简称'),
    ('计算机科学与工程学院', '计信学院', '常见简称'),
    ('计算机科学与工程学院', '软件学院', '历史别名'),
    ('材料科学与工程学院', '材料学院', '标准简称'),
    ('材料科学与工程学院', '材工学院', '标准简称'),
    ('材料科学与工程学院', '材料科学与工程', '末尾漏学院二字'),
    ('材料科学与工程学院', '材料科学与工程x', '历史填报错字'),
    ('材料科学与工程学院', '材料学院与工程学院', '历史填报错字'),
    ('测绘地理信息学院', '测绘学院', '标准简称'),
    ('测绘地理信息学院', '测地学院', '简称变体'),
    ('物理与电子信息工程学院', '物电学院', '标准简称'),
    ('物理与电子信息工程学院', '物理学院', '常见简称'),
    ('物理与电子信息工程学院', '物电工程学院', '别名变体'),
    ('环境科学与工程学院', '环境学院', '标准简称'),
    ('环境科学与工程学院', '环工学院', '标准简称'),
    ('环境科学与工程学院', '环境科学与工程', '末尾漏学院二字'),
    ('环境科学与工程学院', '环境科学院工程学院', '历史填报错字'),
    ('旅游与风景园林学院', '旅风学院', '标准简称'),
    ('旅游与风景园林学院', '旅游学院', '常见简称'),
    ('旅游与风景园林学院', '风景园林学院', '常见简称'),
    ('旅游与风景园林学院', '旅院', '口语极简简称'),
    ('数学与统计学院', '数统学院', '标准简称'),
    ('数学与统计学院', '数学院', '常见简称'),
    ('数学与统计学院', '数统', '口语极简简称'),
    ('外国语学院', '外语学院', '常见简称'),
    ('外国语学院', '外语', '口语极简简称'),
    ('外国语学院', '外院', '常见简称'),
    ('商学院', '经管学院', '历史别名'),
    ('商学院', '经济管理学院', '历史别名'),
    ('艺术学院', '艺院', '口语极简简称'),
    ('艺术学院', '艺术院校', '常见填报错字'),
    ('亚太信息技术学院', '亚太学院', '常见简称'),
    ('马克思主义学院', '马院', '标准简称'),
    ('公共管理与传媒学院', '公管学院', '标准简称'),
    ('公共管理与传媒学院', '传媒学院', '常见简称'),
]

DEFAULT_REASON_CATEGORIES = [
    ['有选修课', '课务学业', '通用课程/事务类别，不含个人指向'],
    ['军事理论课冲突', '课务学业', '全校通用课程冲突'],
    ['生病', '健康医疗', '通用健康情形词，不含具体诊断或症状'],
    ['就医', '健康医疗', '通用健康情形词，不含具体诊断或症状'],
    ['住院', '健康医疗', '通用健康情形词，不含具体诊断或症状'],
    ['治疗', '健康医疗', '通用健康情形词，不含具体诊断或症状'],
]


def default_tables() -> Dict[str, List[List[str]]]:
    return {
        '学院别名': [list(row) for row in DEFAULT_COLLEGE_ALIASES],
        '班级别名': [],
        '校区别名': [list(row) for row in DEFAULT_CAMPUS_ALIASES],
        '免勤理由分类': [list(row) for row in DEFAULT_REASON_CATEGORIES],
        '统计参数': [
            ['统计周次', '7', '当前统计的周次'],
            ['计入校区', '雁山', '屏风整体排除'],
            ['估算标记固定口径', 'FFFFF3CD', '固定口径（与代码常量 core/export.py:ESTIMATE_FILL 一致，修改不生效）；斜体；值为 0 的免勤格同样带标记'],
        ],
    }


def _atomic_save_wb(wb: Workbook, path: str) -> None:
    """Rule 12 防线 2：采用临时文件 + os.replace 原子落盘，防词典损坏。"""
    parent = os.path.dirname(os.path.abspath(path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    tmp_path = '%s.%d.tmp.xlsx' % (path, os.getpid())
    try:
        wb.save(tmp_path)
        wb.close()
        os.replace(tmp_path, path)
    finally:
        wb.close()
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def create_dictionary(path: str) -> str:
    """生成一份默认词典文件。"""
    tables = default_tables()
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEET_NAMES:
        ws = wb.create_sheet(name)
        head = HEADERS[name]
        ws.append(list(head))
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        for row in tables.get(name, []):
            ws.append(list(row))
        for idx, width in enumerate(HEADERS[name], start=1):
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = 26 if idx < 3 else 34
    _atomic_save_wb(wb, path)
    return path


def load_tables(path: Optional[str]) -> Dict[str, List[List[str]]]:
    """读取词典；文件不存在或损坏时回退到默认词典。"""
    tables = default_tables()
    if not path or not os.path.exists(path):
        return tables
    try:
        with open(path, 'rb') as f:
            wb = load_workbook(f, data_only=True)
    except Exception:
        return tables
    try:
        for name in SHEET_NAMES:
            if name not in wb.sheetnames:
                continue
            ws = wb[name]
            rows: List[List[str]] = []
            for r in range(2, ws.max_row + 1):
                values = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                if all(v is None or str(v).strip() == '' for v in values):
                    continue
                rows.append(['' if v is None else str(v).strip() for v in values])
            tables[name] = rows
    finally:
        wb.close()
    return tables


def save_tables(path: str, tables: Dict[str, List[List[str]]]) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEET_NAMES:
        ws = wb.create_sheet(name)
        ws.append(list(HEADERS[name]))
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        for row in tables.get(name, []):
            ws.append(list(row))
        for idx, width in enumerate(HEADERS[name], start=1):
            ws.column_dimensions[ws.cell(1, idx).column_letter].width = 26 if idx < 3 else 34
    _atomic_save_wb(wb, path)


class Dictionary:
    """词典查询接口。"""

    def __init__(self, tables: Union[str, Dict[str, List[List[str]]]]):
        if isinstance(tables, str):
            tables = load_tables(tables)
        self.tables = tables
        self.college_alias: Dict[str, str] = {}
        for row in tables.get('学院别名', []):
            if len(row) >= 2 and row[0] and row[1]:
                formal = ''.join(str(row[0]).split())
                alias_name = ''.join(str(row[1]).split())
                if formal and alias_name:
                    self.college_alias[alias_name] = formal
                    self.college_alias[formal] = formal
        self.class_alias: Dict[str, str] = {}
        for row in tables.get('班级别名', []):
            if len(row) >= 2 and row[0] and row[1]:
                src = ''.join(str(row[0]).split())
                dst = ''.join(str(row[1]).split())
                if src and dst:
                    self.class_alias[src] = dst
                    norm_src = normalize_class(src)
                    if norm_src:
                        self.class_alias[norm_src] = dst
        self.campus_alias: Dict[str, str] = {}
        for row in tables.get('校区别名', []):
            if len(row) >= 2 and row[0] and row[1]:
                self.campus_alias[''.join(str(row[0]).split())] = ''.join(str(row[1]).split())
        self.reason_class: Dict[str, str] = {}
        for row in tables.get('免勤理由分类', []):
            if len(row) >= 2 and row[0] and row[1]:
                self.reason_class[''.join(str(row[0]).split())] = ''.join(str(row[1]).split())
        self.params: Dict[str, str] = {}
        for row in tables.get('统计参数', []):
            if len(row) >= 2 and row[0]:
                self.params[str(row[0]).strip()] = str(row[1]).strip()

    @classmethod
    def load(cls, path: Optional[str]) -> 'Dictionary':
        return cls(load_tables(path))

    def campus(self, text: str) -> str:
        return normalize_campus(text, custom_map=self.campus_alias)

    def class_key(self, text: str) -> str:
        return normalize_class(text, custom_map=self.class_alias)

    def college(self, text: str) -> str:
        key = ''.join(str(text or '').split())
        if not key:
            return ''
        # 1. 精确别名映射（含正名自我映射）
        if key in self.college_alias:
            return self.college_alias[key]
        # 2. 遍历已知官方全称集合进行包含检查（如输入含“化学与生物工程学院”完整字样）
        official_colleges = set(self.college_alias.values())
        for off in sorted(official_colleges, key=len, reverse=True):
            if off in key or (len(key) >= 4 and key in off):
                return off
        # 3. 核心关键字智能匹配兜底（按关键词长度倒序最长匹配）
        keywords = [
            ('化学与生物', '化学与生物工程学院'), ('化生', '化学与生物工程学院'), ('化工', '化学与生物工程学院'),
            ('土木工程', '土木工程学院'), ('土建', '土木工程学院'), ('土木', '土木工程学院'),
            ('地球科学', '地球科学学院'), ('地科', '地球科学学院'), ('地学', '地球科学学院'), ('地球', '地球科学学院'),
            ('机械与控制', '机械与控制工程学院'), ('机控', '机械与控制工程学院'), ('机电', '机械与控制工程学院'), ('机械', '机械与控制工程学院'),
            ('计算机科学', '计算机科学与工程学院'), ('计算机', '计算机科学与工程学院'), ('计科', '计算机科学与工程学院'), ('计网', '计算机科学与工程学院'), ('计信', '计算机科学与工程学院'), ('软件', '计算机科学与工程学院'),
            ('材料科学', '材料科学与工程学院'), ('材料', '材料科学与工程学院'), ('材工', '材料科学与工程学院'),
            ('测绘地理', '测绘地理信息学院'), ('测绘', '测绘地理信息学院'), ('测地', '测绘地理信息学院'),
            ('物理与电子', '物理与电子信息工程学院'), ('物电', '物理与电子信息工程学院'), ('物理', '物理与电子信息工程学院'),
            ('环境科学', '环境科学与工程学院'), ('环境', '环境科学与工程学院'), ('环工', '环境科学与工程学院'),
            ('旅游与风景', '旅游与风景园林学院'), ('风景园林', '旅游与风景园林学院'), ('旅风', '旅游与风景园林学院'), ('旅游', '旅游与风景园林学院'), ('旅院', '旅游与风景园林学院'),
            ('数学与统计', '数学与统计学院'), ('数统', '数学与统计学院'), ('数学', '数学与统计学院'),
            ('外国语', '外国语学院'), ('外语', '外国语学院'),
            ('商学', '商学院'), ('经管', '商学院'), ('经济管理', '商学院'),
            ('艺术', '艺术学院'),
            ('亚太', '亚太信息技术学院'),
            ('马克思', '马克思主义学院'), ('马院', '马克思主义学院'),
            ('传媒', '公共管理与传媒学院'), ('公管', '公共管理与传媒学院'),
        ]
        for kw, off in keywords:
            if kw in key:
                return off
        return text

    def reason_class_of(self, text: str) -> str:
        key = ''.join(str(text or '').split())
        # 1. Excel 词典显式配置
        if key in self.reason_class:
            return self.reason_class[key]
        for kw, cat in self.reason_class.items():
            if kw and kw in key:
                return cat
        # 2. 内置通用分类规则（不随包落盘到 xlsx，纯代码逻辑，安全合规）
        builtin_rules = [
            (('选修', '课程冲突', '课', '实验', '重修', '补考', '答辩', '竞赛', '学业'), '课务学业'),
            (('训练', '校运会', '排练', '比赛', '运动会', '文体'), '文体活动'),
            (('党团', '入党', '入团', '党课', '团课', '积极分子'), '党团事务'),
            (('学生会', '值班', '会议', '志愿', '体测', '公假', '迎新'), '公共事务'),
            (('病', '医', '院', '就医', '住院', '治疗', '拔牙', '体检', '伤', '痛'), '健康医疗'),
            (('事假', '家中有事', '返校', '车票', '休学', '请假', '回家'), '因事请假'),
        ]
        for keywords, cat in builtin_rules:
            if any(k in key for k in keywords):
                return cat
        return '未分类'
