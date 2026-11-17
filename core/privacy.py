# -*- coding: utf-8 -*-
"""模型零隐私接入与双向隐私剔除引擎（Privacy Redaction & Anonymization）。

核心能力 [VERIFIED_ALGO]：
1. 学号不可逆同态伪匿名化：使用加盐 HMAC-SHA256 将 8–13 位学号（含测试号段与缺学号合成键）
   单向映射为 `ANON-XXXXXXXXXX`（异常位数映射为 `ABNORM-XXXXXXXXXX`），保证大模型及脚本
   仅能按匿名标识筛选、去重与统计唯一学号数，绝不暴露任何真实学号前后缀数字。
2. 姓名、病事由、个人班级、联系方式零字符残留剔除：
   - 姓名（含“张*”等半打码形态）彻底抹除为 `[已脱敏]`；
   - 免勤原因严格归并为四类白名单标签（`课务学业`、`社团竞赛训练`、`健康医疗`、`其他`）；
   - 个人班级在脱敏表格中同态匿名化为 `CLS-XXXXXXXX`（保证集体覆盖个人扣减 100% 等价），
     模型 JSON 视图中不透出任何个人班级明细、手机、身份证、邮箱或本地用户路径。
3. 双端覆盖：
   - 输入端预剔除：`strip_workbook_privacy` 生成最小化脱敏副本（同时展平公式），
     确保脱敏副本送入计算引擎仍产出 100% 一致的学院免勤人数与出勤率；
   - 输出端拦截：`build_model_safe_exempt_summary`、`redact_text_for_model` 与
     `assert_model_privacy_compliance` 守护 CLI（JSON/文本）、Skill 与诊断日志全通道。
"""

from __future__ import annotations

import hashlib
import hmac
import os
import re
from contextlib import closing
from io import BytesIO
from typing import Any, Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple

from openpyxl import Workbook, load_workbook

from .formulas import _atomic_save_workbook, flatten_workbook_formulas, xlsx_has_formulas
from .parsing import (extract_group_size, is_group_application, normalize_class,
                      normalize_text, strip_illegal, strip_spaces)

ANON_SID_PREFIX = 'ANON-'
ABNORM_SID_PREFIX = 'ABNORM-'
CLASS_TOKEN_PREFIX = 'CLS-'

ANON_SID_RE = re.compile(r'^(?:ANON|NOSID)-[0-9A-F]{8,16}$')
ABNORM_SID_RE = re.compile(r'^ABNORM-[0-9A-F]{8,16}$')

ALLOWED_FOUR_CATEGORIES = ('课务学业', '社团竞赛训练', '健康医疗', '其他')

_REASON_FOUR_CATEGORY_RULES: Tuple[Tuple[str, Tuple[str, ...]], ...] = (
    ('健康医疗', (
        '健康医疗', '病', '医', '院', '伤', '痛', '诊', '术', '药', '咳', '烧', '炎',
        '复查', '拔牙', '挂水', '急诊', '就医', '治疗', '住院', '生病', '感染', '骨折',
        '发热', '体检', '感冒', '牙', '胃', '发烧',
    )),
    ('课务学业', (
        '课务学业', '课', '学', '考', '试', '重修', '选修', '实验', '上机', '实训',
        '答辩', '四级', '六级', '教资', '冲突', '补考', '作业', '讲座', '学业',
    )),
    ('社团竞赛训练', (
        '社团竞赛训练', '竞赛', '训练', '比赛', '校运会', '运动会', '排练', '演出',
        '文体', '社团', '辩论', '公假', '党', '团', '学工', '学生会', '志愿', '会议',
        '值班', '活动', '组织', '例会', '协会', '迎新', '体测',
    )),
)

# 不得误判为学生姓名的系统常见词
_SAFE_NON_NAME_TERMS = frozenset({
    '雁山', '屏风', '早自习', '晚自习', '自习', '学院', '明细表', '免勤表', '汇总',
    '出勤率', '巡查率', '周平均', '平均', '已核对', '待核对', '已统计', '已回填',
    '已脱敏', '已剔除', '[已脱敏]', '[已剔除]', '[缺学号]', '[班级已剔除]',
    '集体', '全班', '全体', '学生', '测试', '未分类', '全免勤大表',
    '在线存档', '课务学业', '社团竞赛训练', '健康医疗', '其他', '地球科学学院',
    '计算机学院', '理学院', '化工学院', '机械学院', '测绘学院', '环境学院', '土木学院',
})

_USER_DIR_RE = re.compile(
    r'(?:[A-Za-z]:[\\/]+(?:Users|Documents and Settings)|/(?:Users|home))[\\/]+[^\\/\r\n\'"]+',
    re.IGNORECASE,
)
_ID_CARD_RE = re.compile(r'(?<![0-9A-Fa-f.\-])\d{17}[\dXx](?![0-9A-Fa-f.\-])')
_PHONE_RE = re.compile(r'(?<![0-9A-Fa-f.\-])1[3-9]\d{9}(?![0-9A-Fa-f.\-])')
_EMAIL_RE = re.compile(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}')
_RAW_SID_11_13_RE = re.compile(r'(?<![0-9A-Fa-f.\-])(?!99)\d{11,13}(?![0-9A-Fa-f.\-])')
_RAW_SID_8_13_RE = re.compile(r'(?<![0-9A-Fa-f.\-])\d{8,13}(?![0-9A-Fa-f.\-])')

_EXEMPT_SENSITIVE_COL_KEYWORDS = (
    '电话', '手机', '联系方式', '身份证', '证件号', 'qq', '微信', '邮箱', 'email',
    '寝室', '宿舍', '住址', '辅导员', '班主任', '审批', '签名', '联系人', '家长',
    '备注', '说明', '附件', '病历',
)
_DETAIL_REDACT_COL_KEYWORDS = (
    '学号', '姓名', '电话', '手机', '联系方式', '身份证', '巡查', '辅导员', '班主任',
    '备注', '签名', '人工记录', '原因',
)


class PrivacyViolationError(ValueError):
    """模型可见输出中检测到未剔除隐私信息时触发的硬阻断异常。"""


def mask_local_paths(text: Any) -> str:
    """将本地工作区及个人主目录路径脱敏为中性占位标记。"""
    if text is None:
        return ''
    s = str(text)
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
    except OSError:
        pass

    for env_path in (os.path.expanduser('~'), os.environ.get('USERPROFILE')):
        if env_path and env_path in s:
            s = s.replace(env_path, '%USERPROFILE%')

    s = _USER_DIR_RE.sub('%USERPROFILE%', s)
    username = os.environ.get('USERNAME') or os.environ.get('USER')
    if username and len(username) >= 2:
        s = re.sub(re.escape(username), '<USER>', s, flags=re.IGNORECASE)
    s = re.sub(r'\b[A-Za-z]:[\\/]+', '<DRIVE>/', s)
    return s


class PrivacyRedactor:
    """不可逆加盐同态伪匿名与模型侧全量隐私剔除器。"""

    def __init__(self, salt: Optional[bytes] = None) -> None:
        self._salt = salt if salt is not None else os.urandom(32)

    def anonymize_sid(self, sid: Any, *, valid_sid: Optional[bool] = None) -> str:
        """将任意学号/内部合成键单向映射为 `ANON-XXXXXXXXXX` 或 `ABNORM-XXXXXXXXXX`。"""
        if sid is None:
            return '[缺学号]'
        raw = str(sid).strip()
        if not raw:
            return '[缺学号]'
        if ANON_SID_RE.match(raw) or ABNORM_SID_RE.match(raw):
            return raw
        digest = hmac.new(self._salt, raw.encode('utf-8'), hashlib.sha256).hexdigest()[:10].upper()
        prefix = ANON_SID_PREFIX if (valid_sid is None or valid_sid) else ABNORM_SID_PREFIX
        return f'{prefix}{digest}'

    def anonymize_class(self, college: str, class_label: str) -> str:
        """将个人班级名称单向映射为确定性匿名键 `CLS-XXXXXXXX`（同班同校一致）。"""
        norm = normalize_class(class_label)
        payload = f'{strip_spaces(college)}|{norm}'.encode('utf-8')
        digest = hmac.new(self._salt, payload, hashlib.sha256).hexdigest()[:8].upper()
        return f'{CLASS_TOKEN_PREFIX}{digest}'

    @staticmethod
    def classify_reason_four_way(
        reason: Any,
        *,
        custom_classifier: Optional[Callable[[str], str]] = None,
    ) -> str:
        """将免勤原文统一映射为四大白名单分类之一，彻底抛弃原文细节。"""
        raw = strip_spaces(reason)
        if not raw:
            return '其他'
        for cat in ALLOWED_FOUR_CATEGORIES:
            if cat in raw:
                return cat
        if custom_classifier is not None:
            mapped = strip_spaces(custom_classifier(raw))
            if mapped in ALLOWED_FOUR_CATEGORIES:
                return mapped
            if mapped in ('党团事务', '公共事务', '文体活动'):
                return '社团竞赛训练'
            if mapped == '因事请假':
                return '其他'
        for cat, keywords in _REASON_FOUR_CATEGORY_RULES:
            if any(kw in raw for kw in keywords):
                return cat
        return '其他'

    def redact_text(
        self,
        text: Any,
        *,
        extra_names: Iterable[str] = (),
        extra_sids: Iterable[str] = (),
        extra_reasons: Iterable[str] = (),
        extra_classes: Iterable[str] = (),
    ) -> str:
        """对交付给模型的任意文本执行零姓氏、零原文、伪匿名学号深层清洗。"""
        if text is None:
            return ''
        s = mask_local_paths(strip_illegal(str(text)))
        if not s:
            return ''

        # 0. 先抹除内部缺学号合成键 NOSID:college:class:name
        s = re.sub(
            r'NOSID:[^,\s;；，。()（）]+',
            lambda m: self.anonymize_sid(m.group(0), valid_sid=True),
            s,
        )

        # 1. 替换已知真实学号 -> ANON-XXXXXXXXXX
        for sid in sorted({str(x).strip() for x in extra_sids if x and len(str(x).strip()) >= 8}, key=len, reverse=True):
            if sid in s:
                s = s.replace(sid, self.anonymize_sid(sid))

        # 2. 替换已知真实学生姓名（彻底变为 [已脱敏]，不留任何姓氏）
        safe_names = {
            str(n).strip() for n in extra_names
            if n and len(str(n).strip()) >= 2 and str(n).strip() not in _SAFE_NON_NAME_TERMS
        }
        for name in sorted(safe_names, key=len, reverse=True):
            if name in s:
                s = s.replace(name, '[已脱敏]')

        # 3. 替换已知班级名称与原始免勤事由正文
        for cls in sorted({str(c).strip() for c in extra_classes if c and len(str(c).strip()) >= 3}, key=len, reverse=True):
            if cls in s and cls not in _SAFE_NON_NAME_TERMS:
                s = s.replace(cls, '[班级已剔除]')

        for rsn in sorted({str(r).strip() for r in extra_reasons if r and len(str(r).strip()) >= 3}, key=len, reverse=True):
            if rsn in s and rsn not in ALLOWED_FOUR_CATEGORIES and rsn not in _SAFE_NON_NAME_TERMS:
                s = s.replace(rsn, f'[{self.classify_reason_four_way(rsn)}]')

        # 4. 结构化正则兜底：姓名(学号) -> [已脱敏](ANON-...)
        def _sub_name_sid(m: re.Match) -> str:
            return f'[已脱敏]{m.group(2)}{self.anonymize_sid(m.group(3))}{m.group(4)}'

        s = re.sub(
            r'([\u4e00-\u9fa5·•]{2,12})\s*([\(（])(\d{8,13})([\)\)）])',
            _sub_name_sid,
            s,
        )

        # 5. 结构化正则兜底：标签/前缀引导的姓名 -> 彻底清除为 [已脱敏]
        s = re.sub(
            r'(学生|同学|老师|辅导员|班主任|申请人|联系人|姓名\s*[:：=]?)\s*'
            r'([\u4e00-\u9fa5]{2,4}(?:[·•][\u4e00-\u9fa5]{1,10})?)'
            r'(?=因|在|于|申|请|未|已|被|的|有|无|缺|迟|旷|报|反|告|录|登|核|查|冲突|重名|\s|[，。；：！？、,.;:!?\(\)（）/\\]|$)',
            r'\1[已脱敏]',
            s,
        )

        # 6. 身份证号、手机号、邮箱与残留 8–13 位学号
        s = _ID_CARD_RE.sub('[身份证已剔除]', s)
        s = _PHONE_RE.sub('[手机号已剔除]', s)
        s = _EMAIL_RE.sub('[邮箱已剔除]', s)
        s = _RAW_SID_8_13_RE.sub(lambda m: self.anonymize_sid(m.group(0)), s)
        return s

    def redact_anomaly(
        self,
        item: Dict[str, Any],
        *,
        extra_names: Iterable[str] = (),
        extra_sids: Iterable[str] = (),
        extra_reasons: Iterable[str] = (),
        extra_classes: Iterable[str] = (),
    ) -> Dict[str, Any]:
        """对单条异常对象执行模型安全脱敏，移除任何可追溯个人的字段。"""
        cleaned: Dict[str, Any] = {}
        for key, val in item.items():
            if key in ('name', 'sid', 'reason', 'class_text', 'phone', 'id_card'):
                continue
            if isinstance(val, str):
                cleaned[key] = self.redact_text(
                    val,
                    extra_names=extra_names,
                    extra_sids=extra_sids,
                    extra_reasons=extra_reasons,
                    extra_classes=extra_classes,
                )
            elif isinstance(val, dict):
                cleaned[key] = self.redact_anomaly(
                    val,
                    extra_names=extra_names,
                    extra_sids=extra_sids,
                    extra_reasons=extra_reasons,
                    extra_classes=extra_classes,
                )
            elif isinstance(val, list):
                cleaned[key] = [
                    self.redact_anomaly(
                        el,
                        extra_names=extra_names,
                        extra_sids=extra_sids,
                        extra_reasons=extra_reasons,
                        extra_classes=extra_classes,
                    ) if isinstance(el, dict)
                    else (self.redact_text(
                        el,
                        extra_names=extra_names,
                        extra_sids=extra_sids,
                        extra_reasons=extra_reasons,
                        extra_classes=extra_classes,
                    ) if isinstance(el, str) else el)
                    for el in val
                ]
            else:
                cleaned[key] = val
        return cleaned


def collect_sensitive_entities(result: Any) -> Tuple[Set[str], Set[str], Set[str], Set[str]]:
    """从 RunResult 中采集全部真实学生姓名、学号、事由与个人班级，供出站校验与清洗。"""
    names: Set[str] = set()
    sids: Set[str] = set()
    reasons: Set[str] = set()
    classes: Set[str] = set()

    if result is None:
        return names, sids, reasons, classes

    college_whitelist = set(getattr(result, 'college_names', {}) or {}) | set(getattr(result, 'college_names', {}).values() if isinstance(getattr(result, 'college_names', None), dict) else ())

    for rec_list in (getattr(result, 'student_details', None) or {}).values():
        for rec in rec_list or ():
            if not isinstance(rec, dict):
                continue
            sid = str(rec.get('sid') or '').strip()
            name = str(rec.get('name') or '').strip()
            reason = str(rec.get('reason') or '').strip()
            cls = str(rec.get('class_text') or '').strip()
            if sid and not ANON_SID_RE.match(sid) and not ABNORM_SID_RE.match(sid) and not sid.startswith('NOSID:'):
                sids.add(sid)
            if name and len(name) >= 2 and name not in _SAFE_NON_NAME_TERMS and name not in college_whitelist and not name.startswith('全体学生'):
                names.add(name)
            if reason and len(reason) >= 3 and reason not in ALLOWED_FOUR_CATEGORIES and not reason.startswith('['):
                reasons.add(reason)
            if cls and len(cls) >= 3 and cls not in _SAFE_NON_NAME_TERMS and not cls.startswith(CLASS_TOKEN_PREFIX):
                classes.add(cls)

    for rec_list in (getattr(result, 'room_students', None) or {}).values():
        for rec in rec_list or ():
            if not isinstance(rec, dict):
                continue
            sid = str(rec.get('sid') or '').strip()
            name = str(rec.get('name') or '').strip()
            reason = str(rec.get('reason') or '').strip()
            if sid and not ANON_SID_RE.match(sid) and not ABNORM_SID_RE.match(sid) and not sid.startswith('NOSID:'):
                sids.add(sid)
            if name and len(name) >= 2 and name not in _SAFE_NON_NAME_TERMS and name not in college_whitelist:
                names.add(name)
            if reason and len(reason) >= 3 and reason not in ALLOWED_FOUR_CATEGORIES and not reason.startswith('['):
                reasons.add(reason)

    return names, sids, reasons, classes


def build_model_safe_exempt_summary(
    result: Any,
    redactor: Optional[PrivacyRedactor] = None,
    *,
    custom_classifier: Optional[Callable[[str], str]] = None,
) -> List[Dict[str, Any]]:
    """构建给大模型使用的免勤安全摘要：仅含学院免勤人数、4类宏观计数与伪匿名学号索引。"""
    red = redactor or PrivacyRedactor()
    student_details = getattr(result, 'student_details', None) or {}
    exempt_rows = getattr(result, 'exempt_rows', None) or []

    # 按 (学院, 星期, 自习) 聚合匿名学号与 4 类计数
    scope_index: Dict[Tuple[str, str, str], Dict[str, Any]] = {}
    for row in exempt_rows:
        college = str(row.get('学院') or '未识别学院')
        day = str(row.get('星期') or '')
        session = str(row.get('自习') or '')
        key = (college, day, session)
        scope_index[key] = {
            'college': college,
            'day': day,
            'session': session,
            'personal_exempt_count': int(row.get('个人免勤', 0) or 0),
            'group_exempt_count': int(row.get('集体免勤', 0) or 0),
            'total_exempt_count': int(row.get('免勤合计', 0) or 0),
            'status': str(row.get('免勤状态') or ''),
            'unique_anon_sids_count': 0,
            'anon_sids': [],
            'category_counts': {cat: 0 for cat in ALLOWED_FOUR_CATEGORIES},
        }

    for (college, day, session), students in student_details.items():
        key = (college, day, session)
        entry = scope_index.setdefault(key, {
            'college': college,
            'day': day,
            'session': session,
            'personal_exempt_count': len(students),
            'group_exempt_count': 0,
            'total_exempt_count': len(students),
            'status': '已统计',
            'unique_anon_sids_count': 0,
            'anon_sids': [],
            'category_counts': {cat: 0 for cat in ALLOWED_FOUR_CATEGORIES},
        })
        seen_anon: Set[str] = set(entry['anon_sids'])
        for stu in students or ():
            raw_sid = stu.get('sid')
            anon_sid = red.anonymize_sid(raw_sid)
            if anon_sid not in seen_anon:
                seen_anon.add(anon_sid)
                entry['anon_sids'].append(anon_sid)
            raw_cat = stu.get('reason_category') or stu.get('reason') or ''
            cat = red.classify_reason_four_way(raw_cat, custom_classifier=custom_classifier)
            entry['category_counts'][cat] = entry['category_counts'].get(cat, 0) + 1
        entry['anon_sids'].sort()
        entry['unique_anon_sids_count'] = len(entry['anon_sids'])

    return list(scope_index.values())


def assert_model_privacy_compliance(
    payload_text: str,
    *,
    forbidden_names: Iterable[str] = (),
    forbidden_sids: Iterable[str] = (),
    forbidden_reasons: Iterable[str] = (),
    forbidden_classes: Iterable[str] = (),
) -> None:
    """运行期硬阻断闸门：确保输出文本不含任何学生真实隐私信息。"""
    if not payload_text:
        return

    for sid in forbidden_sids:
        s = str(sid or '').strip()
        if s and len(s) >= 8 and not ANON_SID_RE.match(s) and not ABNORM_SID_RE.match(s) and s in payload_text:
            raise PrivacyViolationError('隐私阻断：检测到未脱敏学生学号外泄风险')

    # 扫描任意裸露 8–13 位学号（严格拦截全部真实 8–13 位数字标识）
    m_sid = _RAW_SID_8_13_RE.search(payload_text)
    if m_sid:
        raise PrivacyViolationError(
            f'隐私阻断：检测到未脱敏 8–13 位长数字标识（位置 {m_sid.start()}，长度 {len(m_sid.group(0))}）'
        )

    if _PHONE_RE.search(payload_text):
        raise PrivacyViolationError('隐私阻断：检测到疑似 11 位手机号码外泄风险')

    if _ID_CARD_RE.search(payload_text):
        raise PrivacyViolationError('隐私阻断：检测到疑似 18 位身份证号外泄风险')

    for name in forbidden_names:
        n = str(name or '').strip()
        if n and len(n) >= 2 and n not in _SAFE_NON_NAME_TERMS and n in payload_text:
            raise PrivacyViolationError('隐私阻断：检测到真实学生姓名外泄风险')

    for cls in forbidden_classes:
        c = str(cls or '').strip()
        if c and len(c) >= 4 and c not in _SAFE_NON_NAME_TERMS and c in payload_text:
            raise PrivacyViolationError('隐私阻断：检测到个人真实班级名称外泄风险')

    for rsn in forbidden_reasons:
        r = str(rsn or '').strip()
        if r and len(r) >= 4 and r not in ALLOWED_FOUR_CATEGORIES and r not in _SAFE_NON_NAME_TERMS and r in payload_text:
            raise PrivacyViolationError('隐私阻断：检测到真实免勤事由原文外泄风险')


def _strip_exemption_sheet_privacy(
    ws,
    redactor: PrivacyRedactor,
    *,
    sid_resolver: Optional[Dict[Tuple[str, str, str], str]] = None,
    custom_classifier: Optional[Callable[[str], str]] = None,
) -> Dict[str, int]:
    """对单张免勤工作表执行最小化隐私剔除，返回统计。"""
    stats = {'rows_stripped': 0, 'sids_anonymized': 0, 'names_removed': 0, 'sensitive_cells_cleared': 0}

    # 定位表头行
    header_row = None
    col_map: Dict[str, int] = {}
    sensitive_cols: List[int] = []
    for r in range(1, min(ws.max_row or 0, 25) + 1):
        probe: Dict[str, int] = {}
        sens: List[int] = []
        for c in range(1, (ws.max_column or 0) + 1):
            head = strip_spaces(ws.cell(r, c).value)
            if not head:
                continue
            low = head.lower()
            if head in ('学号', '学生学号') or ('学号' in head and '位数' not in head and '状态' not in head and '异常' not in head):
                if 'sid' not in probe or head == '学号':
                    if 'sid' in probe:
                        sens.append(probe['sid'])
                    probe['sid'] = c
                else:
                    sens.append(c)
            elif head in ('姓名', '学生姓名', '申请人', '申请人姓名') or (
                '姓名' in head and not any(x in head for x in ('辅导员', '老师', '班主任', '审批', '审核', '签'))
            ):
                if 'name' not in probe or head in ('姓名', '学生姓名'):
                    if 'name' in probe:
                        sens.append(probe['name'])
                    probe['name'] = c
                else:
                    sens.append(c)
            elif head in ('学院', '学院名称', '所属学院'):
                probe['college'] = c
            elif '班级' in head:
                probe['class'] = c
            elif '理由' in head or '原因' in head:
                probe['reason'] = c
            elif head in ('免勤周次', '请假周次') or '周次' in head:
                probe['week'] = c
            elif '时间' in head or '星期' in head:
                probe['time'] = c
            elif any(kw in low for kw in _EXEMPT_SENSITIVE_COL_KEYWORDS):
                sens.append(c)
        has_student_id = 'sid' in probe or 'name' in probe
        has_time_dim = any(k in probe for k in ('week', 'time', 'reason', 'class'))
        if ('college' in probe and any(k in probe for k in ('sid', 'name', 'class'))) or (has_student_id and has_time_dim):
            header_row = r
            col_map = probe
            sensitive_cols = sens
            break

    if header_row is None:
        return stats

    sid_col = col_map.get('sid')
    name_col = col_map.get('name')
    college_col = col_map.get('college')
    class_col = col_map.get('class')
    reason_col = col_map.get('reason')

    # 第一遍：收集标准 13 位同学号，保证同名同班缺学号可映射同一匿名 ID
    name_class_to_anon: Dict[Tuple[str, str, str], str] = {}
    if sid_resolver is not None:
        name_class_to_anon.update(sid_resolver)
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        college = normalize_text(ws.cell(r, college_col).value) if college_col else ''
        raw_sid = ws.cell(r, sid_col).value if sid_col else None
        raw_name = normalize_text(ws.cell(r, name_col).value) if name_col else ''
        raw_cls = normalize_text(ws.cell(r, class_col).value) if class_col else ''
        sid_txt = strip_spaces(str(int(raw_sid)) if isinstance(raw_sid, float) and raw_sid.is_integer() else raw_sid)
        if sid_txt.isdigit() and len(sid_txt) <= 13 and len(sid_txt) >= 11:
            norm_sid = sid_txt.zfill(13) if len(sid_txt) < 13 else sid_txt
            if raw_name and raw_cls:
                key = (college, raw_name.replace('・', '·').replace('•', '·'), normalize_class(raw_cls))
                name_class_to_anon[key] = redactor.anonymize_sid(norm_sid, valid_sid=True)

    # 第二遍：逐行剔除隐私
    for r in range(header_row + 1, (ws.max_row or 0) + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, (ws.max_column or 0) + 1)]
        if all(v is None or str(v).strip() == '' for v in row_vals):
            continue

        college = normalize_text(ws.cell(r, college_col).value) if college_col else ''
        raw_name = normalize_text(ws.cell(r, name_col).value) if name_col else ''
        raw_cls = normalize_text(ws.cell(r, class_col).value) if class_col else ''
        raw_reason = normalize_text(ws.cell(r, reason_col).value) if reason_col else ''
        raw_sid_val = ws.cell(r, sid_col).value if sid_col else None

        is_group = is_group_application(raw_name, raw_reason, raw_cls)
        group_size = extract_group_size(raw_name) or extract_group_size(raw_reason) or extract_group_size(raw_cls)

        # 1. 学号脱敏
        if sid_col:
            sid_txt = ''
            if isinstance(raw_sid_val, bool):
                sid_txt = ''
            elif isinstance(raw_sid_val, int):
                sid_txt = str(raw_sid_val)
            elif isinstance(raw_sid_val, float) and raw_sid_val.is_integer():
                sid_txt = str(int(raw_sid_val))
            elif raw_sid_val is not None:
                sid_txt = strip_spaces(raw_sid_val).strip("'\"‘’“”`")

            if not sid_txt:
                if is_group:
                    ws.cell(r, sid_col).value = ''
                elif raw_name and raw_cls:
                    lookup = (college, raw_name.replace('・', '·').replace('•', '·'), normalize_class(raw_cls))
                    if lookup in name_class_to_anon:
                        ws.cell(r, sid_col).value = name_class_to_anon[lookup]
                    else:
                        ws.cell(r, sid_col).value = redactor.anonymize_sid(f'NOSID:{lookup}', valid_sid=True)
                    stats['sids_anonymized'] += 1
            elif sid_txt in ('全班', '全体', '全部', '集体', '无', '待定', '学生', '同学', '/'):
                ws.cell(r, sid_col).value = sid_txt
            elif ANON_SID_RE.match(sid_txt) or ABNORM_SID_RE.match(sid_txt):
                pass
            elif sid_txt.isdigit():
                is_std = 11 <= len(sid_txt) <= 13
                padded = sid_txt.zfill(13) if (isinstance(raw_sid_val, (int, float)) and len(sid_txt) < 13) else sid_txt
                valid_flag = len(padded) == 13
                ws.cell(r, sid_col).value = redactor.anonymize_sid(padded, valid_sid=(valid_flag or is_std))
                stats['sids_anonymized'] += 1
            else:
                ws.cell(r, sid_col).value = redactor.anonymize_sid(sid_txt, valid_sid=False)
                stats['sids_anonymized'] += 1

        # 2. 班级匿名化（保持同班同校映射完全一致，使集体覆盖个人扣减 100% 等价）
        effective_cls = raw_cls
        if not effective_cls and is_group and re.search(r'\d+\s*班', raw_name):
            effective_cls = raw_name
        if class_col and effective_cls:
            if strip_spaces(effective_cls) in ('全班', '全体学生', '全体', '集体', '无', '待定', '-', '/'):
                ws.cell(r, class_col).value = strip_spaces(effective_cls)
            elif not effective_cls.startswith(CLASS_TOKEN_PREFIX):
                ws.cell(r, class_col).value = redactor.anonymize_class(college, effective_cls)

        # 3. 姓名彻底剥离（不留姓氏；集体单仅保留无姓名人数占位）
        if name_col and raw_name:
            if is_group and group_size is not None:
                ws.cell(r, name_col).value = f'全体学生（共{group_size}人）'
            else:
                ws.cell(r, name_col).value = '[已脱敏]'
            stats['names_removed'] += 1

        # 4. 免勤原因收敛为四大合规分类（集体单保留共N人标记）
        if reason_col and raw_reason:
            cat = redactor.classify_reason_four_way(raw_reason, custom_classifier=custom_classifier)
            if is_group and group_size is not None:
                ws.cell(r, reason_col).value = f'集体免勤 共{group_size}人 [{cat}]'
            else:
                ws.cell(r, reason_col).value = f'[{cat}]'

        # 5. 清除联系方式、身份证、备注、辅导员等高敏列
        for sc in sensitive_cols:
            cell = ws.cell(r, sc)
            if cell.value is not None and str(cell.value).strip() != '':
                cell.value = '[已剔除]'
                stats['sensitive_cells_cleared'] += 1

        stats['rows_stripped'] += 1

    return stats


def _strip_detail_or_summary_sheet_privacy(ws, redactor: PrivacyRedactor) -> Dict[str, int]:
    """对明细/汇总工作表清理个人辅助列与残留敏感文本，完整保留数字/出勤率/学院/教室。"""
    stats = {'sensitive_cells_cleared': 0}
    redact_cols: List[int] = []
    scan_rows = min(ws.max_row or 0, 6)
    for c in range(1, (ws.max_column or 0) + 1):
        head = '|'.join(
            strip_spaces(ws.cell(r, c).value)
            for r in range(1, scan_rows + 1)
            if ws.cell(r, c).value is not None
        ).lower()
        if not head:
            continue
        # 避免误伤核心计算列（应到/实到/免勤/出勤率/巡查率/学院/教室）
        if any(safe in head for safe in ('应到', '实到', '免勤', '出勤率', '巡查率', '学院', '教室', '专业班级', '序号')):
            continue
        if any(kw in head for kw in _DETAIL_REDACT_COL_KEYWORDS):
            redact_cols.append(c)

    if not redact_cols:
        return stats

    header_bottom = 3
    for r in range(header_bottom + 1, (ws.max_row or 0) + 1):
        for c in redact_cols:
            cell = ws.cell(r, c)
            if isinstance(cell.value, str) and cell.value.strip():
                cell.value = '[已剔除]'
                stats['sensitive_cells_cleared'] += 1
            elif isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                # 防止学号/手机号以数字存在备注/电话列
                if abs(cell.value) >= 10_000_000:
                    cell.value = '[已剔除]'
                    stats['sensitive_cells_cleared'] += 1
    return stats


def strip_workbook_privacy(
    input_path: str,
    output_path: str,
    *,
    salt: Optional[bytes] = None,
    dictionary_path: Optional[str] = None,
) -> Dict[str, Any]:
    """计算前对工作簿执行隐私剔除与公式展平，原子输出最小化脱敏副本 .xlsx。"""
    if not str(output_path).lower().endswith('.xlsx'):
        raise ValueError('输出文件须使用 .xlsx 扩展名')

    redactor = PrivacyRedactor(salt=salt)
    custom_classifier = None
    if dictionary_path:
        from .dictionary import Dictionary
        custom_classifier = Dictionary.load(dictionary_path).reason_class_of

    with open(input_path, 'rb') as fh:
        raw_bytes = fh.read()

    total_stats: Dict[str, Any] = {
        'rows_stripped': 0,
        'sids_anonymized': 0,
        'names_removed': 0,
        'sensitive_cells_cleared': 0,
        'formulas_frozen': 0,
        'residual_formulas': 0,
    }

    with closing(load_workbook(BytesIO(raw_bytes), data_only=False)) as wb:
        if xlsx_has_formulas(raw_bytes):
            with closing(load_workbook(BytesIO(raw_bytes), data_only=True)) as wb_values:
                flatten_stats = flatten_workbook_formulas(
                    wb, wb_values=wb_values, clear_unresolved=True, sanitize_strings=True
                )
        else:
            flatten_stats = flatten_workbook_formulas(
                wb, wb_values=None, clear_unresolved=True, sanitize_strings=True
            )
        total_stats['formulas_frozen'] = flatten_stats.formulas_frozen
        total_stats['residual_formulas'] = flatten_stats.residual_formulas

        for ws in wb.worksheets:
            ex_stats = _strip_exemption_sheet_privacy(ws, redactor, custom_classifier=custom_classifier)
            if ex_stats['rows_stripped'] > 0:
                total_stats['rows_stripped'] += ex_stats['rows_stripped']
                total_stats['sids_anonymized'] += ex_stats['sids_anonymized']
                total_stats['names_removed'] += ex_stats['names_removed']
                total_stats['sensitive_cells_cleared'] += ex_stats['sensitive_cells_cleared']
            else:
                dt_stats = _strip_detail_or_summary_sheet_privacy(ws, redactor)
                total_stats['sensitive_cells_cleared'] += dt_stats['sensitive_cells_cleared']

        _atomic_save_workbook(wb, output_path)

    return total_stats
