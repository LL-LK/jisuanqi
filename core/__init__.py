# -*- coding: utf-8 -*-
"""三早一晚免勤自动统计系统 · 计算引擎。

对外唯一入口：``from core import RunConfig, run``。
"""

from .model import CollegeBlock, CollegeRate, DetailModel, ExemptRecord, RoomRow
from .pipeline import RunConfig, RunResult, run
from .dictionary import Dictionary, create_dictionary, load_tables, save_tables
from .excel_safety import (
    DiagnosticMasker,
    FileCheckResult,
    classify_excel_file,
    generate_diagnostic_log,
    verify_excel_health,
)

__version__ = '1.0.8'

__all__ = [
    'RunConfig', 'RunResult', 'run',
    'Dictionary', 'create_dictionary', 'load_tables', 'save_tables',
    'DetailModel', 'RoomRow', 'CollegeBlock', 'CollegeRate', 'ExemptRecord',
    'FileCheckResult', 'verify_excel_health', 'classify_excel_file',
    'DiagnosticMasker', 'generate_diagnostic_log',
    '__version__',
]
