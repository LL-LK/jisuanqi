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
import shutil
from datetime import datetime
from zipfile import BadZipFile
from typing import Dict, List, Optional, Union

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils.exceptions import InvalidFileException

from .export import atomic_output
from .parsing import normalize_campus, normalize_class, parse_college_label, strip_spaces

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
            ['统计周次', '7', '参考值，实际周次由输入表或用户选择决定'],
            ['计入校区', '雁山', '默认统计范围，实际范围由运行设置决定'],
            ['空白校区默认值', '雁山', '旧表默认雁山；不确定时填「待确认」以暂停计入'],
            ['出勤率完整计算模式', '0', '0 为严格；1 为命令行未另设时的容错计算'],
            ['估算标记固定口径', 'FFFFF3CD', '固定口径（与代码常量 core/export.py:ESTIMATE_FILL 一致，修改不生效）；斜体；值为 0 的免勤格同样带标记'],
        ],
    }


def create_dictionary(path: str) -> str:
    """生成一份默认词典文件。"""
    save_tables(path, default_tables())
    return path


def load_tables(path: Optional[str], *, strict=False) -> Dict[str, List[List[str]]]:
    """读取词典。编辑预览可回退，正式计算须明确报告指定词典的读入错误。"""
    tables = default_tables()
    if not path:
        return tables
    try:
        with open(path, 'rb') as source:
            wb = load_workbook(source, data_only=False)
    except (OSError, ValueError, KeyError, BadZipFile, InvalidFileException) as exc:
        if strict:
            raise ValueError(f'无法读取免勤词典「{os.path.basename(path)}」') from exc
        return tables
    try:
        present = set(SHEET_NAMES) & set(wb.sheetnames)
        if strict and not present:
            raise ValueError('免勤词典没有可识别的规则工作表')
        for name in SHEET_NAMES:
            if name not in present:
                continue
            ws = wb[name]
            headings = tuple(strip_spaces(ws.cell(1, column).value) for column in (1, 2))
            if strict and headings != HEADERS[name][:2]:
                raise ValueError(f'词典「{name}」的前两列表头应为「{" / ".join(HEADERS[name][:2])}」')
            rows = []
            for cells in ws.iter_rows(min_row=2, max_col=3):
                values = [str(cell.value).strip() if cell.value is not None else '' for cell in cells]
                if not any(values):
                    continue
                if strict and any(cell.data_type == 'f' for cell in cells[:2]):
                    raise ValueError(f'词典「{name}」第 {cells[0].row} 行的规则不能使用公式')
                rows.append(values)
            tables[name] = rows
    finally:
        wb.close()
    return tables


def backup_dictionary(path: str) -> Optional[str]:
    """在保存词典前创建带时间戳的备份文件。

    返回备份文件完整路径，失败或原文件不存在时返回 None。
    """
    if not path or not os.path.exists(path) or not os.path.isfile(path):
        return None
    try:
        dirname = os.path.dirname(os.path.abspath(path))
        basename = os.path.basename(path)
        name, ext = os.path.splitext(basename)
        ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_path = os.path.join(dirname, f'{name}_backup_{ts}{ext}')
        shutil.copy2(path, backup_path)
        return backup_path
    except Exception:
        return None


def save_tables(path: str, tables: Dict[str, List[List[str]]], backup: bool = False) -> Optional[str]:
    """保存词典。当 backup=True 且原文件存在时自动备份。返回备份路径（若有）。"""
    backup_path = None
    if backup and os.path.isfile(path):
        backup_path = backup_dictionary(path)

    Dictionary(tables)
    wb = Workbook()
    wb.remove(wb.active)
    try:
        for name in SHEET_NAMES:
            ws = wb.create_sheet(name)
            ws.append(list(HEADERS[name]))
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                ws.column_dimensions[cell.column_letter].width = 26 if cell.column < 3 else 34
            for row in tables.get(name, []):
                ws.append(list(row))
                for cell in ws[ws.max_row]:
                    if cell.data_type == 'f':
                        cell.data_type = 's'
        with atomic_output(path) as temporary:
            wb.save(temporary)
    finally:
        wb.close()
    return backup_path


def college_key(value):
    name, _ = parse_college_label(value)
    return strip_spaces(name)


def _put_alias(mapping, key, target, sheet, row):
    if not key or not target:
        raise ValueError(f'词典「{sheet}」第 {row} 行的别名和目标均须填写')
    if key in mapping and mapping[key] != target:
        raise ValueError(f'词典「{sheet}」第 {row} 行的规则「{key}」与其他行冲突')
    mapping[key] = target


class Dictionary:
    """词典查询接口。"""

    def __init__(self, tables: Union[str, Dict[str, List[List[str]]]]):
        if isinstance(tables, (str, os.PathLike)):
            tables = load_tables(tables, strict=True)
        self.tables = tables
        self.college_alias: Dict[str, str] = {}
        self.class_alias: Dict[str, str] = {}
        self.campus_alias: Dict[str, str] = {}
        self.reason_class: Dict[str, str] = {}
        self.params: Dict[str, str] = {}
        rules = (
            ('学院别名', self.college_alias, college_key, True),
            ('班级别名', self.class_alias, normalize_class, False),
            ('校区别名', self.campus_alias, strip_spaces, False),
            ('免勤理由分类', self.reason_class, strip_spaces, False),
            ('统计参数', self.params, strip_spaces, False),
        )
        for sheet, mapping, normalize, formal_first in rules:
            for row_number, row in enumerate(tables.get(sheet, ()), 2):
                if not any(strip_spaces(value) for value in row):
                    continue
                first, second = (list(row) + ['', ''])[:2]
                key, target = normalize(first), normalize(second)
                if formal_first:
                    key, target = target, key
                    # 只填正式名称时，作为正式名称登记。
                    key = key or target
                    _put_alias(mapping, target, target, sheet, row_number)
                _put_alias(mapping, key, target, sheet, row_number)
        for formal, alias, _ in DEFAULT_COLLEGE_ALIASES:
            target = self.college_alias.setdefault(college_key(formal), college_key(formal))
            self.college_alias.setdefault(college_key(alias), target)
        for alias, target, _ in DEFAULT_CAMPUS_ALIASES:
            self.campus_alias.setdefault(strip_spaces(alias), target)
        for sheet, mapping in (('班级别名', self.class_alias), ('校区别名', self.campus_alias)):
            for key, target in mapping.items():
                if mapping.get(target, target) != target:
                    raise ValueError(f'词典「{sheet}」的规则「{key}」须直接指向规范名称')
        if self.params.get('出勤率完整计算模式', '0') not in ('0', '1'):
            raise ValueError('词典「出勤率完整计算模式」只接受 0 或 1')
        blank_campus = self.params.get('空白校区默认值', '雁山')
        if blank_campus not in ('雁山', '屏风', '雁门', '待确认') and blank_campus not in self.campus_alias.values():
            raise ValueError('词典「空白校区默认值」须填写规范校区或「待确认」')

    @classmethod
    def load(cls, path: Optional[str]) -> 'Dictionary':
        return cls(load_tables(path, strict=True))

    def campus(self, text: str) -> str:
        return normalize_campus(text, custom_map=self.campus_alias)

    def class_key(self, text: str) -> str:
        return normalize_class(text, custom_map=self.class_alias)

    def exact_college(self, text: str) -> str:
        key = college_key(text)
        return self.college_alias.get(key, key)

    def college(self, text: str) -> str:
        key = college_key(text)
        if not key:
            return ''
        # 1. 精确别名映射（含正名自我映射）
        if key in self.college_alias:
            return self.college_alias[key]
        # 2. 遍历已知官方全称集合进行包含检查（如输入含“化学与生物工程学院”完整字样）
        official_colleges = set(self.college_alias.values())
        candidates = {off for off in official_colleges if off in key or (len(key) >= 4 and key in off)}
        if candidates:
            return next(iter(candidates)) if len(candidates) == 1 else text
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
        candidates = {off for kw, off in keywords if kw in key}
        return next(iter(candidates)) if len(candidates) == 1 else text

    def reason_class_of(self, text: str) -> str:
        key = strip_spaces(text)
        # 1. Excel 词典显式配置
        if key in self.reason_class:
            return self.reason_class[key]
        for kw in sorted(self.reason_class, key=lambda value: (-len(value), value)):
            if kw in key:
                return self.reason_class[kw]
        # 2. 内置通用分类规则（不随包落盘到 xlsx，纯代码逻辑，安全合规）
        builtin_rules = [
            (('党团', '入党', '入团', '党课', '团课', '积极分子'), '党团事务'),
            (('选修', '课程冲突', '课', '实验', '重修', '补考', '答辩', '竞赛', '学业'), '课务学业'),
            (('训练', '校运会', '排练', '比赛', '运动会', '文体'), '文体活动'),
            (('学生会', '值班', '会议', '志愿', '体测', '公假', '迎新'), '公共事务'),
            (('病', '医院', '就医', '住院', '治疗', '拔牙', '体检', '伤', '痛'), '健康医疗'),
            (('事假', '家中有事', '返校', '车票', '休学', '请假', '回家'), '因事请假'),
        ]
        for keywords, cat in builtin_rules:
            if any(k in key for k in keywords):
                return cat
        return '未分类'
