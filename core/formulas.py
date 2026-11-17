# -*- coding: utf-8 -*-
"""Excel 公式安全求值与全工作簿零公式展平引擎。

核心能力 [VERIFIED_ALGO]：
1. 纯 Python 递归下降公式求值器（禁止 eval），支持加减乘除、幂、比较、字符串连接、
   单元格与跨表区域引用，以及 SUM/AVERAGE/COUNT/COUNTA/MIN/MAX/ROUND/ROUNDUP/
   ROUNDDOWN/ABS/INT/IF/IFERROR/AND/OR/NOT/SUMIF/COUNTIF/SUBTOTAL 等常用函数。
2. 语义感知的数值类型与格式归一化：
   - 人数/人次类单元格（如免勤、实到、应到、合计）固化为整数 int（含非零小数时保留 float），
     数字格式维持或设为 '0'（如 39、105）。
   - 比率/百分比类单元格（如出勤率、周平均、巡查率）固化为浮点比例 float，
     数字格式强制维持或设为 '0.00%'（保留两位小数带 % 显示）。
3. 全工作簿零公式展平（flatten_workbook_formulas / export_flattened_workbook），
   确保计算前可清除公式只留数字，且最终导出的 Excel 工作簿无任何公式单元格残留。
"""

from __future__ import annotations

import math
import os
import re
import zipfile
from contextlib import closing, contextmanager
from dataclasses import dataclass, field
from decimal import Decimal, DecimalException, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP
from io import BytesIO
from tempfile import NamedTemporaryFile
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

from .parsing import strip_illegal

EXCEL_ERROR_VALUES = frozenset({
    '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA',
})

PERCENT_HEADER_KEYWORDS = ('出勤率', '巡查率', '周平均', '占比', '百分比')
PERCENT_EXCLUDE_KEYWORDS = ('说明', '备注', '原因', '事由', '姓名', '学号', '班级', '学院', '教室', '状态')

_CELL_REF_RE = re.compile(
    r"^(?:(?:'([^']+)'|([A-Za-z0-9_\u4e00-\u9fff（）()]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$"
)
_SHEET_PREFIX_RE = re.compile(
    r"(?:'((?:[^']|'')+)'|([A-Za-z0-9_\u4e00-\u9fff（）()]+))!"
)
_CELL_COORD_RE = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)")


class FormulaEvalError(ValueError):
    """公式无法安全求值时抛出的内部异常。"""


@dataclass
class FlattenStats:
    """全工作簿公式展平统计信息。"""

    total_formula_cells: int = 0
    frozen_from_cache: int = 0
    evaluated_formulas: int = 0
    cleared_unresolved: int = 0
    residual_formulas: int = 0
    strings_sanitized: int = 0
    unresolved_cells: List[str] = field(default_factory=list)

    @property
    def formulas_frozen(self) -> int:
        """成功固化为数值/结果的公式总数（含缓存命中与实时求值）。"""
        return self.frozen_from_cache + self.evaluated_formulas

    def to_dict(self) -> Dict[str, Any]:
        return {
            'total_formula_cells': self.total_formula_cells,
            'formulas_frozen': self.formulas_frozen,
            'frozen_from_cache': self.frozen_from_cache,
            'evaluated_formulas': self.evaluated_formulas,
            'cleared_unresolved': self.cleared_unresolved,
            'residual_formulas': self.residual_formulas,
            'strings_sanitized': self.strings_sanitized,
            'unresolved_cells': list(self.unresolved_cells),
        }


def is_formula_cell(cell) -> bool:
    """判断 openpyxl 单元格是否为公式格（含字符串公式与 ArrayFormula）。"""
    if cell is None:
        return False
    if cell.data_type == 'f':
        return True
    val = cell.value
    if isinstance(val, str) and val.lstrip().startswith('='):
        return True
    text = getattr(val, 'text', None)
    if isinstance(text, str) and text.lstrip().startswith('='):
        return True
    return False


def extract_formula_text(cell) -> Optional[str]:
    """提取公式单元格的原始公式表达式文本。"""
    if cell is None:
        return None
    val = cell.value
    if isinstance(val, str):
        stripped = val.strip()
        if stripped.startswith('='):
            return stripped
        return ('=' + stripped) if cell.data_type == 'f' else None
    text = getattr(val, 'text', None)
    if isinstance(text, str):
        stripped = text.strip()
        return stripped if stripped.startswith('=') else f'={stripped}'
    return None


def xlsx_has_formulas(raw_bytes: bytes) -> bool:
    """快速探测 .xlsx 字节流中是否包含 XML 公式标签 <f>。"""
    try:
        with zipfile.ZipFile(BytesIO(raw_bytes)) as zf:
            for name in zf.namelist():
                if name.startswith('xl/worksheets/') and name.endswith('.xml'):
                    content = zf.read(name)
                    if b'<f>' in content or b'<f ' in content:
                        return True
    except (zipfile.BadZipFile, OSError):
        return True
    return False


def count_workbook_formulas(wb: Workbook) -> int:
    """统计工作簿全部工作表中仍残留的公式单元格数量。"""
    count = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if is_formula_cell(cell):
                    count += 1
    return count


def _merged_header_value(ws, merged_lookup: Dict[Tuple[int, int], Any], row: int, col: int) -> Any:
    if (row, col) in merged_lookup:
        return merged_lookup[(row, col)]
    return ws.cell(row, col).value


def _build_column_headers(ws, max_header_rows: int = 6) -> Dict[int, str]:
    """聚合前几行（含跨列合并表头）作为列语义文本。"""
    merged_lookup: Dict[Tuple[int, int], Any] = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= max_header_rows:
            top_left = ws.cell(rng.min_row, rng.min_col).value
            if top_left is not None:
                for r in range(rng.min_row, min(rng.max_row, max_header_rows) + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        merged_lookup[(r, c)] = top_left

    headers: Dict[int, str] = {}
    scan_rows = min(ws.max_row or 0, max_header_rows)
    for col in range(1, (ws.max_column or 0) + 1):
        parts: List[str] = []
        for row in range(1, scan_rows + 1):
            val = _merged_header_value(ws, merged_lookup, row, col)
            if val is not None:
                text = re.sub(r'\s+', '', str(val))
                if text and text not in parts:
                    parts.append(text)
        headers[col] = '|'.join(parts)
    return headers


def _is_percent_semantic(cell, col_header: str = '') -> bool:
    fmt = str(getattr(cell, 'number_format', '') or '')
    if '%' in fmt:
        return True
    if not col_header:
        return False
    tokens = [tok for tok in col_header.split('|') if tok]
    for tok in reversed(tokens):
        if any(kw in tok for kw in PERCENT_HEADER_KEYWORDS) or tok.endswith('率'):
            if not any(ex in tok for ex in ('说明', '备注', '原因', '事由', '姓名', '学号', '状态')):
                return True
        if any(ex in tok for ex in ('姓名', '学号', '班级', '教室', '说明', '备注', '原因', '事由', '状态')):
            return False
    return False


def normalize_output_cell_value(
    cell,
    value: Any,
    *,
    col_header: str = '',
    sanitize_fn: Optional[Callable[[str], str]] = None,
) -> Any:
    """按业务语义规范化导出单元格数值与 Excel 显示格式。

    - 出勤率/巡查率等百分比语义格：保持 float 并确保 '0.00%'；
    - 免勤/实到/应到/计数类整数格：转换为 int 并确保 '0'（有自定义浮点格式或非整数小数则保留 float）。
    """
    if isinstance(value, bool) or value is None:
        return value

    cleaner = sanitize_fn or strip_illegal
    is_pct = _is_percent_semantic(cell, col_header)
    fmt = str(getattr(cell, 'number_format', '') or '')

    if isinstance(value, (int, float)):
        fval = float(value)
        if math.isnan(fval) or math.isinf(fval):
            return None
        if is_pct:
            cell.number_format = '0.00%'
            return fval
        if fval.is_integer():
            cell.number_format = '0'
            return int(round(fval))
        return fval

    if isinstance(value, str):
        return cleaner(value)
    return value


# ---------------------------------------------------------------------------
# 纯 Python Excel 公式词法分析与递归下降求值器
# ---------------------------------------------------------------------------

class _RangeRef:
    __slots__ = ('sheet', 'min_row', 'min_col', 'max_row', 'max_col')

    def __init__(self, sheet: str, min_row: int, min_col: int, max_row: int, max_col: int):
        self.sheet = sheet
        self.min_row = min(min_row, max_row)
        self.min_col = min(min_col, max_col)
        self.max_row = max(min_row, max_row)
        self.max_col = max(min_col, max_col)


class _CellRef:
    __slots__ = ('sheet', 'row', 'col')

    def __init__(self, sheet: str, row: int, col: int):
        self.sheet = sheet
        self.row = row
        self.col = col


class _Token:
    __slots__ = ('kind', 'value')

    def __init__(self, kind: str, value: Any):
        self.kind = kind
        self.value = value


def _tokenize_formula(expr: str) -> List[_Token]:
    """将去掉前导 '=' 的公式字符串拆分为 Token 列表。"""
    tokens: List[_Token] = []
    i = 0
    n = len(expr)

    while i < n:
        ch = expr[i]
        if ch.isspace():
            i += 1
            continue

        # 字符串字面量 "..."
        if ch == '"':
            i += 1
            buf: List[str] = []
            while i < n:
                if expr[i] == '"':
                    if i + 1 < n and expr[i + 1] == '"':
                        buf.append('"')
                        i += 2
                        continue
                    i += 1
                    break
                buf.append(expr[i])
                i += 1
            tokens.append(_Token('STR', ''.join(buf)))
            continue

        # 双字符比较符
        if expr[i:i + 2] in ('<=', '>=', '<>', '!='):
            tokens.append(_Token('OP', expr[i:i + 2]))
            i += 2
            continue

        # 单字符运算符与分隔符
        if ch in '+-*/^&=<>(),;':
            kind = 'SEP' if ch in '(),;' else 'OP'
            tokens.append(_Token(kind, ch))
            i += 1
            continue

        # 动态数组隐式交集前缀 @
        if ch == '@':
            i += 1
            continue

        # 跨表引用或引号包裹的工作表名: 'Sheet'!A1 或 Sheet!A1:B2
        if ch == "'":
            m_sheet = _SHEET_PREFIX_RE.match(expr, i)
            if not m_sheet:
                raise FormulaEvalError(f'无法解析工作表前缀: {expr[i:]}')
            sheet_name = (m_sheet.group(1) or '').replace("''", "'")
            i = m_sheet.end()
            m_cell = _CELL_COORD_RE.match(expr, i)
            if not m_cell:
                raise FormulaEvalError(f'工作表引用后缺少单元格坐标: {expr[i:]}')
            col1 = column_index_from_string(m_cell.group(1).upper())
            row1 = int(m_cell.group(2))
            i = m_cell.end()
            if i < n and expr[i] == ':':
                i += 1
                m_cell2 = _CELL_COORD_RE.match(expr, i)
                if not m_cell2:
                    raise FormulaEvalError(f'区域结束坐标无效: {expr[i:]}')
                col2 = column_index_from_string(m_cell2.group(1).upper())
                row2 = int(m_cell2.group(2))
                i = m_cell2.end()
                tokens.append(_Token('RANGE', (sheet_name, row1, col1, row2, col2)))
            else:
                tokens.append(_Token('CELL', (sheet_name, row1, col1)))
            continue

        # 数字（支持科学计数法与百分比后缀）
        if ch.isdigit() or (ch == '.' and i + 1 < n and expr[i + 1].isdigit()):
            m_num = re.match(r'(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?', expr[i:])
            if m_num:
                raw_num = m_num.group(0)
                i += len(raw_num)
                if raw_num.endswith('%'):
                    tokens.append(_Token('NUM', float(raw_num[:-1]) / 100.0))
                elif '.' in raw_num or 'e' in raw_num.lower():
                    tokens.append(_Token('NUM', float(raw_num)))
                else:
                    tokens.append(_Token('NUM', int(raw_num)))
                continue

        # 不带引号的 Sheet!Ref、区域引用、函数名、布尔常量或单元格引用
        m_sheet_plain = _SHEET_PREFIX_RE.match(expr, i)
        if m_sheet_plain:
            sheet_name = m_sheet_plain.group(2)
            pos = m_sheet_plain.end()
            m_cell = _CELL_COORD_RE.match(expr, pos)
            if m_cell:
                col1 = column_index_from_string(m_cell.group(1).upper())
                row1 = int(m_cell.group(2))
                pos = m_cell.end()
                if pos < n and expr[pos] == ':':
                    pos += 1
                    m_cell2 = _CELL_COORD_RE.match(expr, pos)
                    if not m_cell2:
                        raise FormulaEvalError(f'区域结束坐标无效: {expr[pos:]}')
                    col2 = column_index_from_string(m_cell2.group(1).upper())
                    row2 = int(m_cell2.group(2))
                    pos = m_cell2.end()
                    tokens.append(_Token('RANGE', (sheet_name, row1, col1, row2, col2)))
                else:
                    tokens.append(_Token('CELL', (sheet_name, row1, col1)))
                i = pos
                continue

        # 同表区域 $A$1:$B$2 或单元格 $A$1
        m_range = re.match(r'\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?([A-Za-z]{1,3})\$?(\d+)', expr[i:])
        if m_range:
            c1 = column_index_from_string(m_range.group(1).upper())
            r1 = int(m_range.group(2))
            c2 = column_index_from_string(m_range.group(3).upper())
            r2 = int(m_range.group(4))
            i += len(m_range.group(0))
            tokens.append(_Token('RANGE', (None, r1, c1, r2, c2)))
            continue

        m_coord = _CELL_COORD_RE.match(expr, i)
        if m_coord and (m_coord.end() == n or expr[m_coord.end()] != '('):
            c1 = column_index_from_string(m_coord.group(1).upper())
            r1 = int(m_coord.group(2))
            # 列字母不能超过 Excel 最大 3 字母上限 XFD (16384)
            if c1 <= 16384 and r1 >= 1:
                i = m_coord.end()
                tokens.append(_Token('CELL', (None, r1, c1)))
                continue

        # 标识符 / 函数名（兼容 _xlfn. 前缀）
        m_ident = re.match(r'[A-Za-z_][A-Za-z0-9_.]*', expr[i:])
        if m_ident:
            ident = m_ident.group(0)
            i += len(ident)
            if ident.lower().startswith('_xlfn.'):
                ident = ident[6:]
            upper = ident.upper()
            if i < n and expr[i] == '(':
                tokens.append(_Token('FUNC', upper))
            elif upper == 'TRUE':
                tokens.append(_Token('BOOL', True))
            elif upper == 'FALSE':
                tokens.append(_Token('BOOL', False))
            else:
                raise FormulaEvalError(f'不支持的命名标识符: {ident}')
            continue

        raise FormulaEvalError(f'公式含不支持的字符: {expr[i:]}')

    return tokens


def _to_number(val: Any) -> float:
    if val is None or val == '':
        return 0.0
    if isinstance(val, bool):
        return 1.0 if val else 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        text = val.strip()
        if text.endswith('%'):
            return float(text[:-1]) / 100.0
        return float(text)
    raise FormulaEvalError(f'无法转换为数值: {val!r}')


def _to_bool(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    if isinstance(val, str):
        up = val.strip().upper()
        if up == 'TRUE':
            return True
        if up == 'FALSE':
            return False
    raise FormulaEvalError(f'无法转换为布尔值: {val!r}')


def _round_excel(number: float, digits: int, mode: str = 'HALF_UP') -> float:
    if math.isnan(number) or math.isinf(number):
        raise FormulaEvalError(f'非法数值无法舍入: {number!r}')
    try:
        d_num = Decimal(str(number))
        d_digits = int(digits)
        rounding = ROUND_HALF_UP if mode == 'HALF_UP' else (ROUND_UP if mode == 'UP' else ROUND_DOWN)
        if d_digits < 0:
            scale = Decimal('10') ** (-d_digits)
            return float((d_num / scale).quantize(Decimal('1'), rounding=rounding) * scale)
        quant = Decimal('10') ** (-d_digits)
        return float(d_num.quantize(quant, rounding=rounding))
    except (DecimalException, OverflowError, ValueError):
        try:
            res = round(float(number), int(digits))
            if math.isnan(res) or math.isinf(res):
                raise FormulaEvalError(f'舍入结果溢出: {number!r}')
            return res
        except (OverflowError, ValueError) as exc:
            raise FormulaEvalError(f'舍入运算溢出: {exc}') from exc


def _match_criteria(cell_val: Any, criteria: Any) -> bool:
    if isinstance(criteria, (int, float)) and not isinstance(criteria, bool):
        try:
            return not isinstance(cell_val, bool) and _to_number(cell_val) == float(criteria)
        except (ValueError, FormulaEvalError):
            return False
    text = '' if criteria is None else str(criteria)
    for op in ('<=', '>=', '<>', '!=', '=', '<', '>'):
        if text.startswith(op):
            rhs_str = text[len(op):].strip()
            try:
                rhs_num = _to_number(rhs_str)
                lhs_num = _to_number(cell_val) if cell_val not in (None, '') else None
                if lhs_num is None:
                    return False if op in ('=', '<', '<=', '>', '>=') else True
                if op == '=':
                    return lhs_num == rhs_num
                if op in ('<>', '!='):
                    return lhs_num != rhs_num
                if op == '<':
                    return lhs_num < rhs_num
                if op == '<=':
                    return lhs_num <= rhs_num
                if op == '>':
                    return lhs_num > rhs_num
                if op == '>=':
                    return lhs_num >= rhs_num
            except (ValueError, FormulaEvalError):
                lhs_str = '' if cell_val is None else str(cell_val)
                if op == '=':
                    return lhs_str.lower() == rhs_str.lower()
                if op in ('<>', '!='):
                    return lhs_str.lower() != rhs_str.lower()
                return False
    lhs_str = '' if cell_val is None else str(cell_val)
    return lhs_str.lower() == text.lower()


class _WorkbookFormulaContext:
    """带记忆化缓存与循环引用保护的跨表公式求值上下文。"""

    def __init__(self, wb: Workbook, wb_values: Optional[Workbook] = None):
        self.wb = wb
        self.wb_values = wb_values
        self.memo: Dict[Tuple[str, int, int], Any] = {}
        self.visiting: Set[Tuple[str, int, int]] = set()

    def get_cell_value(self, sheet_name: str, row: int, col: int) -> Any:
        key = (sheet_name, row, col)
        if key in self.memo:
            return self.memo[key]
        if sheet_name not in self.wb.sheetnames:
            raise FormulaEvalError(f'引用了不存在的工作表: {sheet_name}')

        ws = self.wb[sheet_name]
        cell = ws.cell(row, col)
        if not is_formula_cell(cell):
            val = cell.value
            self.memo[key] = val
            return val

        # 优先取已有缓存值（当缓存存在且不是错误字符串时）
        if self.wb_values is not None and sheet_name in self.wb_values.sheetnames:
            cached = self.wb_values[sheet_name].cell(row, col).value
            if cached is not None and not (isinstance(cached, str) and cached.strip() in EXCEL_ERROR_VALUES):
                self.memo[key] = cached
                return cached

        if key in self.visiting:
            raise FormulaEvalError(f'检测到循环引用: {sheet_name}!{get_column_letter(col)}{row}')

        formula_text = extract_formula_text(cell)
        if not formula_text:
            raise FormulaEvalError(f'空公式: {sheet_name}!{get_column_letter(col)}{row}')

        self.visiting.add(key)
        try:
            result = self.evaluate_formula(formula_text, sheet_name)
            self.memo[key] = result
            return result
        finally:
            self.visiting.discard(key)

    def iter_range_values(self, rng: _RangeRef) -> List[Any]:
        if rng.sheet not in self.wb.sheetnames:
            raise FormulaEvalError(f'引用了不存在的工作表: {rng.sheet}')
        values: List[Any] = []
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                values.append(self.get_cell_value(rng.sheet, r, c))
        return values

    def evaluate_formula(self, formula_text: str, current_sheet: str) -> Any:
        expr = formula_text.lstrip()
        if expr.startswith('='):
            expr = expr[1:]
        tokens = _tokenize_formula(expr)
        if not tokens:
            raise FormulaEvalError('空公式表达式')
        parser = _Parser(tokens, self, current_sheet)
        val = parser.parse_expression(evaluate=True)
        if parser.pos != len(tokens):
            raise FormulaEvalError(f'公式尾部存在未解析内容: {formula_text}')
        if isinstance(val, _CellRef):
            val = self.get_cell_value(val.sheet, val.row, val.col)
        elif isinstance(val, _RangeRef):
            raise FormulaEvalError('单独的区域引用不能作为标量结果')
        return val


class _Parser:
    """递归下降公式解析求值器，支持 IF / IFERROR 分支短路求值。"""

    def __init__(self, tokens: List[_Token], ctx: _WorkbookFormulaContext, current_sheet: str):
        self.tokens = tokens
        self.ctx = ctx
        self.current_sheet = current_sheet
        self.pos = 0

    def _peek(self) -> Optional[_Token]:
        return self.tokens[self.pos] if self.pos < len(self.tokens) else None

    def _consume(self) -> _Token:
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def _resolve_scalar(self, val: Any) -> Any:
        if isinstance(val, _CellRef):
            return self.ctx.get_cell_value(val.sheet, val.row, val.col)
        if isinstance(val, _RangeRef):
            if val.min_row == val.max_row and val.min_col == val.max_col:
                return self.ctx.get_cell_value(val.sheet, val.min_row, val.min_col)
            raise FormulaEvalError('区域不能直接作为标量参与运算')
        return val

    def parse_expression(self, evaluate: bool = True) -> Any:
        return self._parse_comparison(evaluate)

    def _parse_comparison(self, evaluate: bool) -> Any:
        left = self._parse_concat(evaluate)
        while (tok := self._peek()) is not None and tok.kind == 'OP' and tok.value in ('=', '<>', '!=', '<', '<=', '>', '>='):
            op = self._consume().value
            right = self._parse_concat(evaluate)
            if not evaluate:
                continue
            lv = self._resolve_scalar(left)
            rv = self._resolve_scalar(right)
            left = self._compare(lv, rv, op)
        return left

    @staticmethod
    def _compare(lv: Any, rv: Any, op: str) -> bool:
        if lv is None and rv is None:
            return op in ('=', '<=', '>=')
        if isinstance(lv, str) or isinstance(rv, str):
            ls = '' if lv is None else str(lv).lower()
            rs = '' if rv is None else str(rv).lower()
            if op == '=':
                return ls == rs
            if op in ('<>', '!='):
                return ls != rs
            if op == '<':
                return ls < rs
            if op == '<=':
                return ls <= rs
            if op == '>':
                return ls > rs
            return ls >= rs
        ln = _to_number(lv)
        rn = _to_number(rv)
        if op == '=':
            return ln == rn
        if op in ('<>', '!='):
            return ln != rn
        if op == '<':
            return ln < rn
        if op == '<=':
            return ln <= rn
        if op == '>':
            return ln > rn
        return ln >= rn

    def _parse_concat(self, evaluate: bool) -> Any:
        left = self._parse_add_sub(evaluate)
        while (tok := self._peek()) is not None and tok.kind == 'OP' and tok.value == '&':
            self._consume()
            right = self._parse_add_sub(evaluate)
            if evaluate:
                lv = self._resolve_scalar(left)
                rv = self._resolve_scalar(right)
                left = ('' if lv is None else str(lv)) + ('' if rv is None else str(rv))
        return left

    def _parse_add_sub(self, evaluate: bool) -> Any:
        left = self._parse_mul_div(evaluate)
        while (tok := self._peek()) is not None and tok.kind == 'OP' and tok.value in ('+', '-'):
            op = self._consume().value
            right = self._parse_mul_div(evaluate)
            if evaluate:
                ln = _to_number(self._resolve_scalar(left))
                rn = _to_number(self._resolve_scalar(right))
                left = (ln + rn) if op == '+' else (ln - rn)
        return left

    def _parse_mul_div(self, evaluate: bool) -> Any:
        left = self._parse_power(evaluate)
        while (tok := self._peek()) is not None and tok.kind == 'OP' and tok.value in ('*', '/'):
            op = self._consume().value
            right = self._parse_power(evaluate)
            if evaluate:
                ln = _to_number(self._resolve_scalar(left))
                rn = _to_number(self._resolve_scalar(right))
                if op == '*':
                    left = ln * rn
                else:
                    if rn == 0:
                        raise FormulaEvalError('除数为 0 (#DIV/0!)')
                    left = ln / rn
        return left

    def _parse_power(self, evaluate: bool) -> Any:
        left = self._parse_unary(evaluate)
        if (tok := self._peek()) is not None and tok.kind == 'OP' and tok.value == '^':
            self._consume()
            right = self._parse_power(evaluate)
            if evaluate:
                ln = _to_number(self._resolve_scalar(left))
                rn = _to_number(self._resolve_scalar(right))
                try:
                    res = ln ** rn
                    if isinstance(res, complex):
                        raise FormulaEvalError('指数运算产生复数结果 (#NUM!)')
                    if isinstance(res, float) and (math.isnan(res) or math.isinf(res)):
                        raise FormulaEvalError('指数运算结果溢出 (#NUM!)')
                    return res
                except (OverflowError, ArithmeticError, ValueError) as exc:
                    raise FormulaEvalError(f'指数运算溢出: {exc}') from exc
        return left

    def _parse_unary(self, evaluate: bool) -> Any:
        tok = self._peek()
        if tok is not None and tok.kind == 'OP' and tok.value in ('+', '-'):
            op = self._consume().value
            val = self._parse_unary(evaluate)
            if not evaluate:
                return None
            num = _to_number(self._resolve_scalar(val))
            return num if op == '+' else -num
        return self._parse_primary(evaluate)

    def _parse_primary(self, evaluate: bool) -> Any:
        tok = self._peek()
        if tok is None:
            raise FormulaEvalError('公式意外结束')

        if tok.kind in ('NUM', 'STR', 'BOOL'):
            self._consume()
            return tok.value

        if tok.kind == 'CELL':
            self._consume()
            sheet, row, col = tok.value
            return _CellRef(sheet or self.current_sheet, row, col)

        if tok.kind == 'RANGE':
            self._consume()
            sheet, r1, c1, r2, c2 = tok.value
            return _RangeRef(sheet or self.current_sheet, r1, c1, r2, c2)

        if tok.kind == 'SEP' and tok.value == '(':
            self._consume()
            val = self.parse_expression(evaluate=evaluate)
            closing_tok = self._peek()
            if closing_tok is None or closing_tok.kind != 'SEP' or closing_tok.value != ')':
                raise FormulaEvalError('缺少右括号 )')
            self._consume()
            return val

        if tok.kind == 'FUNC':
            return self._parse_function_call(evaluate)

        raise FormulaEvalError(f'无法识别的语法单元: {tok.kind}={tok.value!r}')

    def _consume_open_paren(self) -> None:
        tok = self._peek()
        if tok is None or tok.kind != 'SEP' or tok.value != '(':
            raise FormulaEvalError('函数调用后缺少左括号 (')
        self._consume()

    def _skip_remaining_args(self) -> None:
        while (tok := self._peek()) is not None:
            if tok.kind == 'SEP' and tok.value in (',', ';'):
                self._consume()
                self.parse_expression(evaluate=False)
                continue
            break

    def _parse_function_call(self, evaluate: bool) -> Any:
        func_name = self._consume().value
        self._consume_open_paren()

        # 短路求值函数：IF / IFERROR
        if func_name == 'IF':
            cond_raw = self.parse_expression(evaluate=evaluate)
            cond = _to_bool(self._resolve_scalar(cond_raw)) if evaluate else False
            if (tok := self._peek()) is None or tok.kind != 'SEP' or tok.value not in (',', ';'):
                raise FormulaEvalError('IF 缺少第 2 个参数')
            self._consume()
            true_val = self.parse_expression(evaluate=(evaluate and cond))
            false_val: Any = False
            if (tok := self._peek()) is not None and tok.kind == 'SEP' and tok.value in (',', ';'):
                self._consume()
                false_val = self.parse_expression(evaluate=(evaluate and not cond))
            self._expect_close_paren(func_name)
            return true_val if (evaluate and cond) else (false_val if evaluate else None)

        if func_name == 'IFERROR':
            saved_pos = self.pos
            fallback: Any = None
            val: Any = None
            failed = False
            if evaluate:
                try:
                    val = self.parse_expression(evaluate=True)
                    val = self._resolve_scalar(val)
                    if isinstance(val, str) and val.strip() in EXCEL_ERROR_VALUES:
                        failed = True
                except (FormulaEvalError, ValueError, ArithmeticError, DecimalException):
                    failed = True
                    self.pos = saved_pos
                    self.parse_expression(evaluate=False)
            else:
                self.parse_expression(evaluate=False)
            if (tok := self._peek()) is not None and tok.kind == 'SEP' and tok.value in (',', ';'):
                self._consume()
                fallback = self._resolve_scalar(self.parse_expression(evaluate=(evaluate and failed)))
            self._expect_close_paren(func_name)
            if not evaluate:
                return None
            return fallback if failed else val

        # 常规急切求值函数
        args: List[Any] = []
        if (tok := self._peek()) is not None and not (tok.kind == 'SEP' and tok.value == ')'):
            while True:
                args.append(self.parse_expression(evaluate=evaluate))
                sep = self._peek()
                if sep is not None and sep.kind == 'SEP' and sep.value in (',', ';'):
                    self._consume()
                    continue
                break
        self._expect_close_paren(func_name)
        if not evaluate:
            return None
        return self._exec_function(func_name, args)

    def _expect_close_paren(self, func_name: str) -> None:
        tok = self._peek()
        if tok is None or tok.kind != 'SEP' or tok.value != ')':
            raise FormulaEvalError(f'函数 {func_name} 缺少右括号 )')
        self._consume()

    def _flatten_numeric_args(self, args: Sequence[Any], *, include_bool_scalar: bool = True) -> List[float]:
        nums: List[float] = []
        for arg in args:
            if isinstance(arg, _RangeRef):
                for cell_val in self.ctx.iter_range_values(arg):
                    if isinstance(cell_val, (int, float)) and not isinstance(cell_val, bool):
                        nums.append(float(cell_val))
            elif isinstance(arg, _CellRef):
                cell_val = self.ctx.get_cell_value(arg.sheet, arg.row, arg.col)
                if isinstance(cell_val, (int, float)) and not isinstance(cell_val, bool):
                    nums.append(float(cell_val))
            else:
                if arg is None or arg == '':
                    continue
                if isinstance(arg, bool):
                    if include_bool_scalar:
                        nums.append(1.0 if arg else 0.0)
                    continue
                nums.append(_to_number(arg))
        return nums

    def _flatten_nonempty_args(self, args: Sequence[Any]) -> List[Any]:
        items: List[Any] = []
        for arg in args:
            if isinstance(arg, _RangeRef):
                for cell_val in self.ctx.iter_range_values(arg):
                    if cell_val is not None and cell_val != '':
                        items.append(cell_val)
            else:
                scalar = self._resolve_scalar(arg)
                if scalar is not None and scalar != '':
                    items.append(scalar)
        return items

    def _exec_function(self, name: str, args: List[Any]) -> Any:
        if name == 'SUM':
            return sum(self._flatten_numeric_args(args))
        if name in ('AVERAGE', 'AVG'):
            nums = self._flatten_numeric_args(args)
            if not nums:
                raise FormulaEvalError('AVERAGE 区间无有效数值 (#DIV/0!)')
            return sum(nums) / len(nums)
        if name == 'COUNT':
            return len(self._flatten_numeric_args(args))
        if name == 'COUNTA':
            return len(self._flatten_nonempty_args(args))
        if name == 'MIN':
            nums = self._flatten_numeric_args(args)
            return min(nums) if nums else 0
        if name == 'MAX':
            nums = self._flatten_numeric_args(args)
            return max(nums) if nums else 0
        if name == 'ROUND':
            if len(args) != 2:
                raise FormulaEvalError('ROUND 需要 2 个参数')
            return _round_excel(_to_number(self._resolve_scalar(args[0])),
                                int(_to_number(self._resolve_scalar(args[1]))), 'HALF_UP')
        if name == 'ROUNDUP':
            if len(args) != 2:
                raise FormulaEvalError('ROUNDUP 需要 2 个参数')
            return _round_excel(_to_number(self._resolve_scalar(args[0])),
                                int(_to_number(self._resolve_scalar(args[1]))), 'UP')
        if name == 'ROUNDDOWN':
            if len(args) != 2:
                raise FormulaEvalError('ROUNDDOWN 需要 2 个参数')
            return _round_excel(_to_number(self._resolve_scalar(args[0])),
                                int(_to_number(self._resolve_scalar(args[1]))), 'DOWN')
        if name == 'ABS':
            if len(args) != 1:
                raise FormulaEvalError('ABS 需要 1 个参数')
            return abs(_to_number(self._resolve_scalar(args[0])))
        if name == 'INT':
            if len(args) != 1:
                raise FormulaEvalError('INT 需要 1 个参数')
            return math.floor(_to_number(self._resolve_scalar(args[0])))
        if name == 'AND':
            return all(_to_bool(x) for x in self._flatten_nonempty_args(args))
        if name == 'OR':
            return any(_to_bool(x) for x in self._flatten_nonempty_args(args))
        if name == 'NOT':
            if len(args) != 1:
                raise FormulaEvalError('NOT 需要 1 个参数')
            return not _to_bool(self._resolve_scalar(args[0]))
        if name == 'COUNTIF':
            if len(args) != 2 or not isinstance(args[0], (_RangeRef, _CellRef)):
                raise FormulaEvalError('COUNTIF 需要条件区间和条件')
            vals = (self.ctx.iter_range_values(args[0]) if isinstance(args[0], _RangeRef)
                    else [self.ctx.get_cell_value(args[0].sheet, args[0].row, args[0].col)])
            crit = self._resolve_scalar(args[1])
            return sum(1 for v in vals if _match_criteria(v, crit))
        if name == 'SUMIF':
            if len(args) not in (2, 3) or not isinstance(args[0], _RangeRef):
                raise FormulaEvalError('SUMIF 参数错误')
            crit_vals = self.ctx.iter_range_values(args[0])
            sum_rng = args[2] if len(args) == 3 and isinstance(args[2], _RangeRef) else args[0]
            sum_vals = self.ctx.iter_range_values(sum_rng)
            crit = self._resolve_scalar(args[1])
            total = 0.0
            for cv, sv in zip(crit_vals, sum_vals):
                if _match_criteria(cv, crit) and isinstance(sv, (int, float)) and not isinstance(sv, bool):
                    total += float(sv)
            return total
        if name == 'SUBTOTAL':
            if len(args) < 2:
                raise FormulaEvalError('SUBTOTAL 至少需要 2 个参数')
            fn_code = int(_to_number(self._resolve_scalar(args[0])))
            rest = args[1:]
            if fn_code in (9, 109):
                return sum(self._flatten_numeric_args(rest))
            if fn_code in (1, 101):
                nums = self._flatten_numeric_args(rest)
                if not nums:
                    raise FormulaEvalError('SUBTOTAL AVERAGE 无数值')
                return sum(nums) / len(nums)
            if fn_code in (2, 102):
                return len(self._flatten_numeric_args(rest))
            if fn_code in (3, 103):
                return len(self._flatten_nonempty_args(rest))
            if fn_code in (4, 104):
                nums = self._flatten_numeric_args(rest)
                return max(nums) if nums else 0
            if fn_code in (5, 105):
                nums = self._flatten_numeric_args(rest)
                return min(nums) if nums else 0
            raise FormulaEvalError(f'SUBTOTAL 暂不支持功能码 {fn_code}')

        raise FormulaEvalError(f'暂不支持的 Excel 函数: {name}')


def evaluate_workbook_cell(
    wb: Workbook,
    sheet_name: str,
    row: int,
    col: int,
    *,
    wb_values: Optional[Workbook] = None,
) -> Any:
    """对工作簿中单个单元格执行只读公式求值（优先命中 wb_values 缓存）。"""
    ctx = _WorkbookFormulaContext(wb, wb_values=wb_values)
    return ctx.get_cell_value(sheet_name, row, col)


def flatten_workbook_formulas(
    wb: Workbook,
    *,
    wb_values: Optional[Workbook] = None,
    clear_unresolved: bool = True,
    sanitize_strings: bool = True,
    sanitize_fn: Optional[Callable[[str], str]] = None,
    updated_cells: Optional[Dict[Tuple[str, int, int], Any]] = None,
) -> FlattenStats:
    """将工作簿全部工作表的全部公式格展平为纯数字/结果值。

    - 若单元格位于百分比列或原有百分比格式，输出 float 且保证 number_format='0.00%'；
    - 若单元格为人数/人次/普通整数格，规范化为 int 并保持/设为 '0'；
    - clear_unresolved=True 时，无法求值且无有效缓存的公式同样会被清理为 None，
      保证导出后工作簿 0 个公式残留。
    """
    stats = FlattenStats()
    ctx = _WorkbookFormulaContext(wb, wb_values=wb_values)
    cleaner = sanitize_fn or strip_illegal

    # 先收集全部待展平的公式单元格与其缓存/求值结果，避免前格提前覆盖影响同表后续依赖公式
    resolved: List[Tuple[str, int, int, Any, str, str]] = []  # (sheet, r, c, value, source, col_header)
    for ws in wb.worksheets:
        headers = _build_column_headers(ws)
        for row in ws.iter_rows():
            for cell in row:
                if is_formula_cell(cell):
                    stats.total_formula_cells += 1
                    cached = None
                    if wb_values is not None and ws.title in wb_values.sheetnames:
                        cached = wb_values[ws.title].cell(cell.row, cell.column).value
                    if cached is not None and not (isinstance(cached, str) and cached.strip() in EXCEL_ERROR_VALUES):
                        resolved.append((ws.title, cell.row, cell.column, cached, 'cache', headers.get(cell.column, '')))
                        continue
                    try:
                        val = ctx.get_cell_value(ws.title, cell.row, cell.column)
                        if isinstance(val, str) and val.strip() in EXCEL_ERROR_VALUES:
                            raise FormulaEvalError(f'公式返回错误值 {val}')
                        resolved.append((ws.title, cell.row, cell.column, val, 'eval', headers.get(cell.column, '')))
                    except (FormulaEvalError, ValueError, ArithmeticError, DecimalException):
                        resolved.append((ws.title, cell.row, cell.column, None, 'unresolved', headers.get(cell.column, '')))
                elif sanitize_strings and isinstance(cell.value, str):
                    cleaned = cleaner(cell.value)
                    if cleaned != cell.value:
                        cell.value = cleaned
                        stats.strings_sanitized += 1

    # 统一写回单元格值并归一化格式
    for sheet_name, row_idx, col_idx, val, source, col_header in resolved:
        cell = wb[sheet_name].cell(row_idx, col_idx)
        coord = f'{sheet_name}!{get_column_letter(col_idx)}{row_idx}'
        if source == 'unresolved':
            stats.unresolved_cells.append(coord)
            if clear_unresolved:
                cell.value = None
                cell.data_type = 'n'
                stats.cleared_unresolved += 1
                if updated_cells is not None:
                    updated_cells[(sheet_name, row_idx, col_idx)] = None
            continue

        norm_val = normalize_output_cell_value(cell, val, col_header=col_header, sanitize_fn=cleaner)
        if norm_val is None and val is not None:
            # NaN/Inf 等非法数值按无法解析处理
            stats.unresolved_cells.append(coord)
            if clear_unresolved:
                cell.value = None
                cell.data_type = 'n'
                stats.cleared_unresolved += 1
            continue

        cell.value = norm_val
        if source == 'cache':
            stats.frozen_from_cache += 1
        else:
            stats.evaluated_formulas += 1
        if updated_cells is not None:
            updated_cells[(sheet_name, row_idx, col_idx)] = norm_val

    stats.residual_formulas = count_workbook_formulas(wb)
    return stats


def load_evaluated_workbook(path: str) -> Tuple[Workbook, FlattenStats]:
    """只读载入工作簿并将其中未缓存公式在内存中实时求值展平。"""
    with open(path, 'rb') as fh:
        raw_bytes = fh.read()
    if not xlsx_has_formulas(raw_bytes):
        wb = load_workbook(BytesIO(raw_bytes), data_only=True)
        return wb, FlattenStats()
    wb = load_workbook(BytesIO(raw_bytes), data_only=False)
    try:
        with closing(load_workbook(BytesIO(raw_bytes), data_only=True)) as wb_values:
            stats = flatten_workbook_formulas(wb, wb_values=wb_values, clear_unresolved=True, sanitize_strings=False)
        return wb, stats
    except Exception:
        wb.close()
        raise


@contextmanager
def open_evaluated_workbook(path: str):
    """上下文管理器：打开已求值工作簿并在退出时自动关闭。"""
    wb, stats = load_evaluated_workbook(path)
    try:
        yield wb, stats
    finally:
        wb.close()


def _atomic_save_workbook(wb: Workbook, output_path: str) -> None:
    out_dir = os.path.dirname(os.path.abspath(output_path)) or '.'
    os.makedirs(out_dir, exist_ok=True)
    with NamedTemporaryFile(dir=out_dir, prefix='.flatten-', suffix='.xlsx', delete=False) as handle:
        tmp_path = handle.name
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, output_path)
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass
        raise


def export_flattened_workbook(input_path: str, output_path: str) -> FlattenStats:
    """将含公式的输入工作簿展平并原子保存为纯数值 Excel（绝不保留公式格）。"""
    if not str(output_path).lower().endswith('.xlsx'):
        raise ValueError('输出文件须使用 .xlsx 扩展名')
    with open(input_path, 'rb') as fh:
        raw_bytes = fh.read()

    with closing(load_workbook(BytesIO(raw_bytes), data_only=False)) as wb:
        if xlsx_has_formulas(raw_bytes):
            with closing(load_workbook(BytesIO(raw_bytes), data_only=True)) as wb_values:
                stats = flatten_workbook_formulas(wb, wb_values=wb_values, clear_unresolved=True, sanitize_strings=True)
        else:
            stats = flatten_workbook_formulas(wb, wb_values=None, clear_unresolved=True, sanitize_strings=True)
        if stats.residual_formulas > 0:
            raise RuntimeError(f'展平后仍存在 {stats.residual_formulas} 个公式单元格')
        _atomic_save_workbook(wb, output_path)
    return stats
