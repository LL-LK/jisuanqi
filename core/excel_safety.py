# -*- coding: utf-8 -*-
"""三早一晚免勤自动统计系统 · 健壮性前置熔断、智能分类与脱敏诊断引擎。

核心能力：
1. verify_excel_health: 物理尺寸、独占锁定、旧版 OLE2 .xls、OOXML 魔数、ZIP 完整性与密码加密探测；
2. classify_excel_file: 文件名特征先验 + 工作表与表头前 5 行指纹识别 (明细表 DETAIL / 免勤表 EXEMPT)；
3. DiagnosticMasker: 个人信息 (PII) 零泄露脱敏 (学号、姓名、就医病由泛化、用户主目录路径)；
4. generate_diagnostic_log: 一键生成安全脱敏的 Markdown 故障排查诊断报告。
"""

from __future__ import annotations

import hashlib
import os
import platform
import re
import sys
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Any, Dict, List, Optional


@dataclass
class FileCheckResult:
    """Excel 文件前置健康自检结果。

    Attributes:
        is_valid: 是否健康可用。True 表示可安全交由 openpyxl 处理。
        error_code: 机器可读的标准化状态码。
        user_message: 用户可读的处理提示。
        tech_detail: 详细技术根因，用于技术排查与脱敏诊断日志落盘。
    """
    is_valid: bool
    error_code: str
    user_message: str
    tech_detail: str = ''

    def __bool__(self) -> bool:
        """支持 if result: 语法。"""
        return self.is_valid

    def to_dict(self) -> Dict[str, Any]:
        return {
            'is_valid': self.is_valid,
            'error_code': self.error_code,
            'user_message': self.user_message,
            'tech_detail': self.tech_detail,
        }


def verify_excel_health(path: str) -> FileCheckResult:
    """毫秒级前置 Excel 健康自检与智能熔断。

    检测范围：文件有效性、物理大小、独占锁占用、OLE2 旧版伪装、OOXML 魔数、
    ZIP 完整性、密码加密标记以及核心内容模型。
    """
    if not path or not isinstance(path, str):
        return FileCheckResult(
            is_valid=False,
            error_code='INVALID_PATH',
            user_message='请选择 Excel 文件。',
            tech_detail=f'Invalid path argument: {path!r}',
        )

    # 1. 物理存在性与文件大小
    if not os.path.exists(path):
        return FileCheckResult(
            is_valid=False,
            error_code='NOT_FOUND',
            user_message='找不到文件，请确认路径。',
            tech_detail=f'Path does not exist: {path}',
        )

    if os.path.isdir(path):
        return FileCheckResult(
            is_valid=False,
            error_code='IS_DIRECTORY',
            user_message='请选择 Excel 文件，当前路径是文件夹。',
            tech_detail=f'Path is a directory: {path}',
        )

    try:
        size = os.path.getsize(path)
    except OSError as e:
        return FileCheckResult(
            is_valid=False,
            error_code='UNREADABLE',
            user_message='无法获取文件信息，请检查权限。',
            tech_detail=f'os.path.getsize failed: {e}',
        )

    if size == 0:
        return FileCheckResult(
            is_valid=False,
            error_code='EMPTY_FILE',
            user_message='文件为空，请保存表格后重试。',
            tech_detail='File size is 0 bytes',
        )

    if size < 4:
        return FileCheckResult(
            is_valid=False,
            error_code='TOO_SMALL',
            user_message='文件不完整，无法读取。',
            tech_detail=f'File size {size} bytes is too small for magic header',
        )

    # 2. 魔数嗅探与独占锁检查
    try:
        with open(path, 'rb') as f:
            header = f.read(8)
    except PermissionError as e:
        return FileCheckResult(
            is_valid=False,
            error_code='LOCKED',
            user_message='文件被占用，请保存并关闭后重试。',
            tech_detail=f'PermissionError: {e}',
        )
    except OSError as e:
        return FileCheckResult(
            is_valid=False,
            error_code='UNREADABLE',
            user_message='无法读取文件，请检查文件权限。',
            tech_detail=f'OSError reading file: {e}',
        )

    # OLE2 复合二进制魔数 (Excel 97-2003 .xls: D0 CF 11 E0)
    if header.startswith(b'\xD0\xCF\x11\xE0'):
        return FileCheckResult(
            is_valid=False,
            error_code='OLD_XLS_FORMAT',
            user_message='这是旧版 .xls 文件，请另存为 .xlsx 后重试。',
            tech_detail='Magic header matches OLE2 CFB (Excel 97-2003: D0 CF 11 E0), openpyxl requires OOXML zip format',
        )

    # 标准 OOXML zip 魔数必须以 PK\x03\x04 起始
    if not header.startswith(b'PK\x03\x04'):
        hex_bytes = header[:4].hex().upper() if len(header) >= 4 else header.hex().upper()
        return FileCheckResult(
            is_valid=False,
            error_code='CORRUPTED',
            user_message='无法识别 Excel 格式，请检查扩展名和文件完整性。',
            tech_detail=f'Invalid magic bytes: 0x{hex_bytes} (expected 0x504B0304)',
        )

    # 3. ZIP 完整性与密码加密位检测
    try:
        with zipfile.ZipFile(path, 'r') as zf:
            # 探测加密：flag_bits Bit 0 代表通用加密标志
            for info in zf.infolist():
                if info.flag_bits & 0x1:
                    return FileCheckResult(
                        is_valid=False,
                        error_code='ENCRYPTED',
                        user_message='文件设有打开密码，请另存无密码副本后重试。',
                        tech_detail=f'Zip entry {info.filename!r} has encryption bit set (flag_bits={info.flag_bits:#x})',
                    )

            # 测试 ZIP 结构完整性（校验 CRC32 与解压头）
            bad_file = zf.testzip()
            if bad_file is not None:
                return FileCheckResult(
                    is_valid=False,
                    error_code='CORRUPTED',
                    user_message=f'表格压缩包内「{bad_file}」数据损坏。',
                    tech_detail=f'testzip failed on member: {bad_file}',
                )

            # 校验是否具备 Excel OOXML 核心结构标记
            names = zf.namelist()
            if not any(n == '[Content_Types].xml' or n.startswith('xl/') for n in names):
                return FileCheckResult(
                    is_valid=False,
                    error_code='NOT_EXCEL',
                    user_message='文件缺少 Excel 工作簿结构。',
                    tech_detail='Missing Excel OOXML structures in zip archive',
                )
    except zipfile.BadZipFile as e:
        return FileCheckResult(
            is_valid=False,
            error_code='CORRUPTED',
            user_message='Excel 文件已损坏，无法读取。',
            tech_detail=f'zipfile.BadZipFile: {e}',
        )
    except Exception as e:
        return FileCheckResult(
            is_valid=False,
            error_code='CORRUPTED' if 'Bad CRC' in str(e) else 'UNKNOWN_ZIP_ERR',
            user_message=f'无法检查文件：{e}',
            tech_detail=f'{type(e).__name__}: {e}',
        )

    return FileCheckResult(
        is_valid=True,
        error_code='OK',
        user_message='文件可读取。',
        tech_detail='',
    )


def classify_excel_file(path: str) -> str:
    """快速嗅探 Excel 类别：返回 'DETAIL' | 'SUMMARY' | 'EXEMPT' | 'DICTIONARY' | 'UNKNOWN'。

    采用 Tier 1 文件名先验打分 + Tier 2 工作表名与表头指纹识别的多层机制。
    带严格的 openpyxl.close() 资源释放保障。
    """
    health = verify_excel_health(path)
    if not health.is_valid:
        return 'UNKNOWN'

    base_name = os.path.basename(path).lower()
    detail_score = 0
    exempt_score = 0
    summary_score = 0
    dict_score = 0

    # 1. Tier 1: 文件名关键词先验
    dict_strong_kw = ('词典', 'dictionary', '免勤词典')
    detail_strong_kw = ('明细', '执行情况', '早晚自习', '自习执行')
    detail_medium_kw = ('早自习', '晚自习', '考勤', '考勤明细', '自习')
    exempt_strong_kw = ('免勤', '请假', '免修', '免勤人员', '免勤汇总')
    exempt_medium_kw = ('人员统计', '免勤统计', '假条', '请假汇总')
    summary_strong_kw = ('终稿', '汇总', '考勤汇总', '出勤统计')

    if any(k in base_name for k in dict_strong_kw):
        dict_score += 5

    if any(k in base_name for k in detail_strong_kw):
        detail_score += 3
    elif any(k in base_name for k in detail_medium_kw):
        detail_score += 2

    if any(k in base_name for k in exempt_strong_kw):
        exempt_score += 3
    elif any(k in base_name for k in exempt_medium_kw):
        exempt_score += 2

    if any(k in base_name for k in summary_strong_kw):
        summary_score += 3

    # 2. Tier 2: 工作表名称与前 5 行表头指纹提取
    wb = None
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        sheet_names_lower = [s.lower() for s in sheet_names]

        # 2.1 词典表特征工作表
        dict_sheets = {'学院别名', '班级别名', '校区别名', '免勤理由分类', '统计参数'}
        matched_dict_sheets = dict_sheets & set(sheet_names)
        if matched_dict_sheets:
            dict_score += len(matched_dict_sheets) * 5

        # 2.2 工作表名特征
        if any('雁山' in s or '屏风' in s for s in sheet_names):
            detail_score += 5
        if any('免勤' in s or '请假' in s or '免修' in s for s in sheet_names):
            exempt_score += 4
        if any('sheet2' in s or '终稿' in s or '汇总' in s for s in sheet_names_lower):
            summary_score += 4

        # 2.3 扫描前两个工作表的前 5 行文本特征
        cell_texts = set()
        for sname in sheet_names[:2]:
            ws = wb[sname]
            for row in ws.iter_rows(max_row=5, values_only=True):
                for val in row:
                    if val is not None:
                        s = str(val).strip()
                        if s.startswith("'"):
                            s = s[1:].strip()  # 吸收 WorkBuddy 前导单引号规整模式
                        if s:
                            cell_texts.add(s)

        detail_fingerprints = {'教室', '应到人数', '分配专业班级', '学生出勤率', '值班辅导员', '执行效果', '应到', '实到'}
        exempt_fingerprints = {'学号', '免勤周次', '免勤自习', '免勤原因', '请假类型', '自习室', '免勤时间'}
        summary_fingerprints = {'学生出勤率', '值班辅导员巡查率', '执行效果', '早晚自习执行情况终稿', '辅导员巡查率', '巡查率'}

        # 汇总表独有特征：含出勤率与巡查率，且不含教室
        has_classroom = any('教室' in c for c in cell_texts)
        has_student_rate = any('出勤率' in c for c in cell_texts)
        has_patrol_rate = any('巡查率' in c for c in cell_texts)
        if has_student_rate and has_patrol_rate and not has_classroom:
            summary_score += 10

        detail_hits = sum(1 for k in detail_fingerprints if any(k in c for c in cell_texts))
        exempt_hits = sum(1 for k in exempt_fingerprints if any(k in c for c in cell_texts))
        summary_hits = sum(1 for k in summary_fingerprints if any(k in c for c in cell_texts))

        detail_score += detail_hits * 3
        exempt_score += exempt_hits * 3
        summary_score += summary_hits * 3
    except Exception:
        # 轻量读取失败不中断，沿用 Tier 1 分数
        pass
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass

    scores = [
        (detail_score, 'DETAIL'),
        (exempt_score, 'EXEMPT'),
        (summary_score, 'SUMMARY'),
        (dict_score, 'DICTIONARY'),
    ]
    scores.sort(key=lambda x: x[0], reverse=True)
    best_score, best_type = scores[0]
    if best_score >= 3:
        if len(scores) > 1 and scores[1][0] == best_score:
            return 'UNKNOWN'
        return best_type
    return 'UNKNOWN'


def preflight_excel_file(path: str) -> Dict[str, Any]:
    """对单个 Excel 文件执行全面的前置健康诊断与表型推断。

    整合物理健康校验、11 种标准化异常代码诊断、脱敏技术根因、
    文件尺寸与 SHA256 指纹及表型自动嗅探（DETAIL/SUMMARY/EXEMPT/DICTIONARY/UNKNOWN）。
    """
    health = verify_excel_health(path)
    table_type = 'UNKNOWN'
    sheet_names: List[str] = []
    file_size = 0
    sha_prefix = 'N/A'

    if path and os.path.exists(path) and not os.path.isdir(path):
        try:
            file_size = os.path.getsize(path)
            sha_prefix = _file_sha256_prefix(path)
        except OSError:
            pass

    if health.is_valid:
        table_type = classify_excel_file(path)
        wb = None
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True)
            sheet_names = list(wb.sheetnames)
        except Exception:
            pass
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    return {
        'path': DiagnosticMasker.mask_path(path) if path else '',
        'is_valid': health.is_valid,
        'error_code': health.error_code,
        'user_message': health.user_message,
        'tech_detail': DiagnosticMasker.mask_path(health.tech_detail) if health.tech_detail else '',
        'table_type': table_type,
        'file_size_bytes': file_size,
        'file_size_formatted': _format_size(file_size),
        'sha256_prefix': sha_prefix,
        'sheet_names': sheet_names,
    }


class DiagnosticMasker:
    """个人信息（PII）与敏感环境路径脱敏工具类。"""

    # 预编译正则，匹配本地用户主目录（支持 Windows 各盘符 Users 及 Unix /Users, /home）
    _USER_DIR_RE = re.compile(
        r'(?:[A-Za-z]:[\\/]+(?:Users|Documents and Settings)|/(?:Users|home))[\\/]+[^\\/\r\n\'"]+',
        re.IGNORECASE,
    )

    # 免勤理由归一化词典
    _REASON_CATEGORIES = [
        ('健康医疗', (
            '病', '医', '院', '伤', '痛', '诊', '术', '药', '咳', '烧', '炎',
            '复查', '拔牙', '挂水', '急诊', '就医', '治疗', '住院', '生病', '感染', '骨折', '发热'
        )),
        ('课务学业', (
            '课', '学', '考', '试', '重修', '选修', '实验', '上机', '辩论', '竞赛',
            '实训', '答辩', '四级', '六级', '教资', '冲突', '补考'
        )),
        ('公共事务', (
            '公假', '团', '党', '学工', '志愿', '会议', '值班', '巡查', '排练',
            '演出', '开会', '活动', '组织', '例会', '协会'
        )),
        ('因事请假', (
            '事假', '家中有事', '出行', '回家', '返校', '车票', '红白', '搬宿'
        )),
    ]

    @classmethod
    def mask_sid(cls, sid: Any) -> str:
        """脱敏学号：保留前缀与末尾校验位，中间打码。"""
        if sid is None:
            return '***'
        s = str(sid).strip()
        length = len(s)
        if length < 8:
            return '***'
        if length == 13:
            return s[:4] + '*' * 7 + s[-2:]
        if length >= 10:
            prefix_len = 5
            suffix_len = 2
            masked_len = max(4, length - prefix_len - suffix_len)
            return s[:prefix_len] + '*' * masked_len + s[-suffix_len:]
        if length >= 8:
            return s[:3] + '*' * (length - 5) + s[-2:]
        return '***'

    @classmethod
    def mask_name(cls, name: Any) -> str:
        """脱敏姓名：支持 2 字、3 字及多字姓名（含少数民族间隔点）中间打码。"""
        if name is None:
            return ''
        s = str(name).strip().strip('·•')
        length = len(s)
        if length == 0:
            return ''
        if length == 1:
            return '*'
        if length == 2:
            return s[0] + '*'
        return s[0] + '*' * (length - 2) + s[-1]

    @classmethod
    def mask_reason(cls, reason: Any) -> str:
        """泛化免勤理由：将具体的医疗就诊、病历描述归纳为分类标签。"""
        if reason is None:
            return '[未填写]'
        s = str(reason).strip()
        if not s:
            return '[未填写]'
        for cat_name, keywords in cls._REASON_CATEGORIES:
            if any(k in s for k in keywords):
                return f'[{cat_name}]'
        return '[其他请假]'

    @classmethod
    def mask_path(cls, path: Any) -> str:
        """脱敏路径：将工作区及祖先目录屏蔽为 <WORKSPACE>，用户个人目录统一屏蔽为 %USERPROFILE%，其余盘符路径屏蔽为 <DRIVE>/。"""
        if path is None:
            return ''
        s = str(path)
        if not s:
            return ''

        try:
            p = os.path.abspath(os.getcwd())
            ancestors = [p]
            while p and os.path.dirname(p) != p:
                p = os.path.dirname(p)
                if os.path.splitdrive(p)[1].strip('\\/'):
                    ancestors.append(p)
            for anc in sorted(ancestors, key=len, reverse=True):
                anc_pat = re.escape(anc).replace(r'\\', r'[\\/]+')
                s = re.sub(anc_pat, '<WORKSPACE>', s, flags=re.IGNORECASE)
        except Exception:
            pass

        user_home = os.path.expanduser('~')
        if user_home and user_home in s:
            s = s.replace(user_home, '%USERPROFILE%')

        user_profile = os.environ.get('USERPROFILE')
        if user_profile and user_profile in s:
            s = s.replace(user_profile, '%USERPROFILE%')

        s = cls._USER_DIR_RE.sub('%USERPROFILE%', s)

        username = os.environ.get('USERNAME') or os.environ.get('USER')
        if username and len(username) >= 2:
            s = re.sub(re.escape(username), '<USER>', s, flags=re.IGNORECASE)

        s = re.sub(r'\b[A-Za-z]:[\\/]+', '<DRIVE>/', s)
        return s


def _file_sha256_prefix(path: str, length: int = 12) -> str:
    """安全读取文件前 64KB 计算哈希前缀，用于轻量版本指纹比对。"""
    if not os.path.exists(path) or os.path.isdir(path):
        return 'N/A'
    try:
        h = hashlib.sha256()
        with open(path, 'rb') as f:
            chunk = f.read(65536)
            h.update(chunk)
        return h.hexdigest()[:length] + '...'
    except Exception:
        return 'ERR'


def _format_size(size_bytes: int) -> str:
    """格式化字节大小。"""
    if size_bytes < 1024:
        return f'{size_bytes} B'
    if size_bytes < 1024 * 1024:
        return f'{size_bytes / 1024:.1f} KB'
    return f'{size_bytes / (1024 * 1024):.2f} MB'


def _sanitize_anomaly_text(text: str, extra_names: Optional[set] = None) -> str:
    """全面清洗异常条目文本，抹除所有学号、姓名与系统路径。"""
    if not text:
        return ''
    names_to_mask = extra_names if extra_names is not None else set()
    s = DiagnosticMasker.mask_path(str(text))

    # 1. 匹配并脱敏 姓名(学号) 或 姓名（学号）
    def _mask_name_sid(m: re.Match) -> str:
        name = m.group(1)
        bracket_open = m.group(2)
        sid = m.group(3)
        bracket_close = m.group(4)
        names_to_mask.add(name)
        return f'{DiagnosticMasker.mask_name(name)}{bracket_open}{DiagnosticMasker.mask_sid(sid)}{bracket_close}'

    s = re.sub(
        r'([\u4e00-\u9fa5·•]{2,12})\s*([\(（])(\d{8,13})([\)\)）])',
        _mask_name_sid,
        s,
    )

    # 2. 匹配 学生/同学/老师/辅导员 + 姓名（支持含间隔点的少数民族长姓名，并在常见助词/标点前截断）
    def _mask_student_name(m: re.Match) -> str:
        prefix = m.group(1)
        name = m.group(2)
        names_to_mask.add(name)
        return prefix + DiagnosticMasker.mask_name(name)

    s = re.sub(
        r'(学生|同学|老师|辅导员)([\u4e00-\u9fa5]{2,4}(?:[·•][\u4e00-\u9fa5]{1,10})?)'
        r'(?=因|在|于|申|请|未|已|被|的|有|无|缺|迟|旷|报|反|告|录|登|核|查|\s|[，。；：！？、,.;:!?\(\)（）/\\]|$)',
        _mask_student_name,
        s,
    )

    # 3. 替换已知学生姓名
    if names_to_mask:
        for name in sorted(names_to_mask, key=len, reverse=True):
            if name in s:
                s = s.replace(name, DiagnosticMasker.mask_name(name))

    # 4. 独立数字学号
    s = re.sub(
        r'(?<!\d)\d{8,13}(?!\d)',
        lambda m: DiagnosticMasker.mask_sid(m.group(0)),
        s,
    )
    return s


def generate_diagnostic_log(
    config: Any,
    result: Optional[Any] = None,
    error_trace: Optional[str] = None,
    anomalies: Optional[List[Dict[str, Any]]] = None,
) -> str:
    """生成完全脱敏的工业级系统诊断与故障排查 Markdown 报告。

    包含：
    1. 运行环境与系统规格 (OS, Python, PySide6, 时间戳);
    2. 任务配置与文件健康状态 (脱敏路径, 文件大小, 哈希前缀, 健康校验);
    3. 表格结构指纹与列头元数据 (脱敏表头);
    4. 异常条目清单与统计摘要 (脱敏 scope 与 note);
    5. 错误堆栈 (脱敏文件路径).
    """
    masker = DiagnosticMasker
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    qt_version = 'N/A'
    if 'PySide6.QtCore' in sys.modules:
        try:
            from PySide6.QtCore import __version__ as qv
            qt_version = f'PySide6 {qv}'
        except Exception:
            pass

    lines: List[str] = [
        '# 本地考勤工具 · 诊断日志',
        '',
        '## 一、运行环境信息',
        f'- **生成时间**：{now_str}',
        f'- **操作系统**：{platform.platform()} ({platform.machine()})',
        f'- **Python 版本**：{platform.python_version()} ({platform.python_implementation()})',
        f'- **GUI 框架**：{qt_version}',
        f'- **工作目录**：`{masker.mask_path(os.getcwd())}`',
        '',
        '## 二、任务配置与文件指纹',
    ]

    detail_path = getattr(config, 'detail_path', '') or ''
    exempt_path = getattr(config, 'exempt_path', '') or ''
    exempt_paths = getattr(config, 'exemption_files', (exempt_path,) if exempt_path else ())
    week_no = getattr(config, 'week_no', 7)
    exempt_source = getattr(config, 'exempt_source', '未分类')
    campus = getattr(config, 'campus', '雁山')

    lines.append(f'- **统计周次**：第 {week_no} 周')
    lines.append(f'- **计入校区**：{campus}')
    lines.append(f'- **免勤来源分类**：{exempt_source}')

    # 明细表信息
    lines.append('')
    lines.append('### 1. 执行情况明细表')
    lines.append(f'- **配置路径**：`{masker.mask_path(detail_path)}`')
    if detail_path and os.path.exists(detail_path):
        size = os.path.getsize(detail_path)
        h_res = verify_excel_health(detail_path)
        sha = _file_sha256_prefix(detail_path)
        lines.append(f'- **存在状态**：存在 (大小: {_format_size(size)}, SHA256: `{sha}`)')
        lines.append(f'- **健康自检**：`{h_res.error_code}` ({h_res.user_message})')
    else:
        lines.append('- **存在状态**：文件不存在或路径为空')

    # 免勤表信息
    lines.append('')
    lines.append('### 2. 免勤人员统计表')
    if exempt_paths:
        for index, path in enumerate(exempt_paths, 1):
            lines.append(f'- **配置路径 {index}**：`{masker.mask_path(path)}`')
            if not os.path.isfile(path):
                lines.append('- **存在状态**：文件不存在')
                continue
            size = os.path.getsize(path)
            h_res = verify_excel_health(path)
            sha = _file_sha256_prefix(path)
            lines.append(f'- **存在状态**：存在 (大小: {_format_size(size)}, SHA256: `{sha}`)')
            lines.append(f'- **健康自检**：`{h_res.error_code}` ({h_res.user_message})')
    else:
        lines.append('- **配置路径**：未提供免勤表（纯明细统计）')

    # 表格结构与脱敏表头嗅探
    lines.append('')
    lines.append('## 三、表格结构与表头元数据')
    files = [('明细表', detail_path)]
    files.extend((f'免勤表 {index}', path) for index, path in enumerate(exempt_paths, 1))
    for label, path in files:
        if not path or not os.path.exists(path) or os.path.isdir(path):
            continue
        health = verify_excel_health(path)
        if not health.is_valid:
            lines.append(f'- **{label}**：自检未通过 ({health.error_code})，跳过结构嗅探')
            continue

        wb = None
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path, read_only=True, data_only=True)
            lines.append(f'- **{label} 工作表列表**：{wb.sheetnames}')
            for sname in wb.sheetnames[:3]:
                ws = wb[sname]
                header_cells = []
                for row in ws.iter_rows(max_row=1, values_only=True):
                    header_cells = [
                        _sanitize_anomaly_text(str(c).strip())
                        for c in row if c is not None
                    ][:15]
                    break
                lines.append(f'  - `{sname}` 首行表头采样: {header_cells}')
        except Exception as e:
            lines.append(f'- **{label} 结构嗅探异常**：{masker.mask_path(str(e))}')
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception:
                    pass

    # 异常条目清单与统计摘要
    lines.append('')
    lines.append('## 四、异常条目清单')
    target_anomalies = anomalies
    if target_anomalies is None and result is not None:
        target_anomalies = getattr(result, 'anomalies', [])

    if result is not None:
        ok = getattr(result, 'ok', False)
        err_msg = _sanitize_anomaly_text(str(getattr(result, "error", "") or ""))
        status_text = '成功完成' if ok else f'执行失败 ({err_msg})'
        lines.append(f'- **执行结果**：{status_text}')
        if hasattr(result, 'sheets'):
            lines.append(f'- **产物工作表**：{result.sheets}')
        rates = getattr(result, 'college_rates', [])
        lines.append(f'- **学院出勤率条目数**：{len(rates)}')

    anom_list = target_anomalies or []
    lines.append(f'- **异常清单总数**：{len(anom_list)}')
    if anom_list:
        lines.append('')
        lines.append('| 序号 | 级别 | 异常类型 | 范围 (已脱敏) | 说明 (已脱敏) |')
        lines.append('| :--- | :--- | :--- | :--- | :--- |')
        known_student_names: set = set()
        for idx, a in enumerate(anom_list[:20], start=1):
            lvl = a.get('level', '')
            atype = a.get('type', '')
            raw_scope = a.get('scope', '')
            raw_note = a.get('note', '')
            scope_masked = _sanitize_anomaly_text(raw_scope, known_student_names)
            note_masked = _sanitize_anomaly_text(raw_note, known_student_names)
            lines.append(f'| {idx} | {lvl} | {atype} | {scope_masked} | {note_masked} |')
        if len(anom_list) > 20:
            lines.append(f'| ... | ... | ... | 剩余 {len(anom_list) - 20} 条已省略 | ... |')

    # 错误堆栈
    lines.append('')
    lines.append('## 五、错误堆栈跟踪 (Traceback)')
    err = error_trace or (getattr(result, 'error', '') if result else '')
    if err:
        masked_err = _sanitize_anomaly_text(str(err))
        lines.append('```text')
        lines.append(masked_err.strip())
        lines.append('```')
    else:
        lines.append('无异常记录。')

    lines.append('')
    lines.append('---')
    lines.append('分享前请核对日志内容。')
    full_text = '\n'.join(lines)
    return _sanitize_anomaly_text(full_text)
