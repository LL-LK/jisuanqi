# -*- coding: utf-8 -*-
"""文本解析与归一化。

本模块是「三早一晚」免勤统计系统的地基：真实数据里的周次、星期、校区、班级
写法都存在大量不规范形式，解析必须容错且**不得漏算**——漏一个周次或一个
星期，就会让某天的免勤人数整体偏低，出勤率直接失真。

已在本项目真实样本中观察到的脏格式（均必须支持）：
    周次：'第7周' '第7-18周' '第6-7，9-13周' '第7—12周' '第七周，第八周'
          '第7、8、9周' '6-10周' '第7周、第8周'
    星期：'周一、周三、周四、周五' '周一 ，周三 ，周四' '周一到周五'
          '周一周二周三周四周五' '周一、二、四、五'
    校区：'雁山校区' '雁山校园' '燕山校区'(错字) '屏风' '雁门校区'
"""

from __future__ import annotations

import re
import unicodedata
from typing import Iterable, Optional, Tuple

# ---------------------------------------------------------------- 基础归一化

_FULLWIDTH_MAP = {
    '，': ',', '。': '.', '；': ';', '：': ':', '（': '(', '）': ')',
    '［': '[', '］': ']', '、': '、', '　': ' ', '～': '~', '－': '-',
    '—': '-', '–': '-', '―': '-', '〜': '~', '＝': '=',
}

WEEKDAY_NAMES = ('周一', '周二', '周三', '周四', '周五', '周六', '周日')
_WEEKDAY_BY_CN = {'一': '周一', '二': '周二', '三': '周三', '四': '周四',
                  '五': '周五', '六': '周六', '日': '周日', '天': '周日'}

_CN_DIGITS = {'零': 0, '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
              '六': 6, '七': 7, '八': 8, '九': 9}


# XML 1.0 不允许大部分控制字符，直接写进 Excel 会抛
# "All strings must be XML compatible"。所有进入系统的文本都要先过一遍。
_ILLEGAL_XML = re.compile(
    '[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f\ud800-\udfff\ufffe\uffff]')


def strip_illegal(value: str) -> str:
    """去掉 XML 非法字符。"""
    return _ILLEGAL_XML.sub('', value)


def normalize_text(value) -> str:
    """全角转半角 + 去首尾空白 + 折叠内部空白 + 去 XML 非法字符。"""
    if value is None:
        return ''
    text = str(value)
    text = strip_illegal(text)
    text = unicodedata.normalize('NFKC', text)
    for src, dst in _FULLWIDTH_MAP.items():
        text = text.replace(src, dst)
    # NFKC 已把部分全角符号转好，这里再兜底处理残留
    text = text.replace('，', ',').replace('、', '、')
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def strip_spaces(text: str) -> str:
    """去掉所有空白，用于紧凑匹配。"""
    return re.sub(r'\s+', '', normalize_text(text))


# ---------------------------------------------------------------- 周次解析

def _cn_number(text: str) -> Optional[int]:
    """把 '七' / '十七' / '二十' 这类中文数字转成 int。"""
    text = strip_spaces(text)
    if not text:
        return None
    if text.isdigit():
        return int(text)
    if '十' in text:
        head, _, tail = text.partition('十')
        tens = _CN_DIGITS.get(head, 1) if head else 1
        ones = _CN_DIGITS.get(tail, 0) if tail else 0
        return tens * 10 + ones
    total = 0
    for ch in text:
        if ch not in _CN_DIGITS:
            return None
        total = total * 10 + _CN_DIGITS[ch]
    return total


_CN_NUM_CHARS = '零〇一二三四五六七八九十'
DEFAULT_WEEK_NO = 7          # 与词典「统计周次」默认值一致
MAX_WEEK_NO = 40             # 学期周次上限：防止把时钟/其它数字读成周次


def _cn_to_digits(text: str) -> str:
    """把**周次上下文**里的中文数字转成阿拉伯数字。

    只在「数字 + 周」或「中文数字 + 区间符 + 段尾『周』」的上下文里转换；
    裸中文数字（'第一 三 四'）**不转换** —— 按裁定 R10 它属于星期。
    若全局转换，'第一 三 四' 会被读成周次 {1,3,4}，凭空造出没人申请过的周次。

    >>> _cn_to_digits('第七周，第八周')
    '第7周,第8周'
    >>> _cn_to_digits('第六-十七周')
    '6-17周'
    >>> _cn_to_digits('第一 三 四')
    '第一 三 四'
    """
    text = normalize_text(text)
    if not text:
        return text

    def _span(m):
        lo = _cn_number(m.group(1))
        hi = _cn_number(m.group(2))
        if lo is None or hi is None:
            return m.group(0)
        return '%d-%d' % (lo, hi)

    # 区间优先：'第六-十七周' / '第六至十七周'
    text = re.sub(r'(?:第)?([%s]+)\s*(?:-|—|–|~|至|到)\s*(?:第)?([%s]+)\s*周'
                  % (_CN_NUM_CHARS, _CN_NUM_CHARS), _span, text)

    # 单值：'第七周' / '第八周'
    def _one(m):
        num = _cn_number(m.group(1))
        return '第%d周' % num if num is not None else m.group(0)

    return re.sub(r'第([%s]+)周' % _CN_NUM_CHARS, _one, text)


def _mask(text: str, pattern: str, handler) -> str:
    """按 pattern 挖空（用 span 重建字符串）。

    **禁止** str.replace 原地替换后继续沿用旧 offset：位置整体错位是本项目
    踩过的真实坑（挖掉第一个「周一」后，后面所有匹配的起点都会偏）。
    """
    out = []
    last = 0
    for m in re.finditer(pattern, text):
        out.append(text[last:m.start()])
        handler(m)
        out.append(' ')
        last = m.end()
    out.append(text[last:])
    return ''.join(out)


def _weekday_span(start: str, end: str) -> set:
    """'一'..'五' → {'周一'..'周五'}（区间式；'周一到周五' 必须展开成 5 天）。"""
    lo = WEEKDAY_NAMES.index(_WEEKDAY_BY_CN[start])
    hi = WEEKDAY_NAMES.index(_WEEKDAY_BY_CN[end])
    if lo > hi:
        lo, hi = hi, lo
    return set(WEEKDAY_NAMES[lo:hi + 1])


def _segment_facts(seg: str, default_kind: str, window_week: Optional[int] = None):
    """识别**一个段**内的周次与星期，返回 (weeks:set[int], weekdays:set[str])。

    default_kind：'WEEK'（免勤周次列）/ 'TIME'（免勤时间列）。
    window_week：统计窗口周次；'全学期' 写法按窗口周计（裁定 R4）。
    """
    text = normalize_text(seg)
    weeks: set = set()
    weekdays: set = set()
    if not text:
        return weeks, weekdays

    has_week_marker = bool(re.search(r'\d\s*周|第\s*[\d%s]+\s*周' % _CN_NUM_CHARS, text))

    # ① 星期区间：'周一到周五'（必须展开，否则漏周二/周三/周四）
    def _span(m):
        weekdays.update(_weekday_span(m.group(1), m.group(2)))
    text = _mask(text, r'周([一二三四五六日天])\s*(?:到|至|-|—|–|~)\s*周?([一二三四五六日天])', _span)

    # ② 星期单点：'周一' / '星期二' / '礼拜天'
    def _point(m):
        weekdays.add(_WEEKDAY_BY_CN[m.group(1)])
    text = _mask(text, r'(?:周|星期|礼拜)\s*([一二三四五六日天])', _point)

    # ③ 周次区间（中文数字）：'第六-十七周' → {6..17}
    def _cn_range(m):
        lo = _cn_number(m.group(1))
        hi = _cn_number(m.group(2))
        if lo is None or hi is None or hi > MAX_WEEK_NO:
            return
        if lo > hi:
            lo, hi = hi, lo
        weeks.update(range(lo, hi + 1))
    text = _mask(text, r'第?\s*([%s]+)\s*(?:-|—|–|~|至|到)\s*第?\s*([%s]+)\s*周'
                 % (_CN_NUM_CHARS, _CN_NUM_CHARS), _cn_range)

    # ④ 周次区间（阿拉伯数字）：'第6-7' / '第7周-第8周' / '9-13周' / '6-10周'
    def _num_range(m):
        lo = int(m.group(1))
        hi = int(m.group(2))
        if hi > MAX_WEEK_NO:
            return
        if lo > hi:
            lo, hi = hi, lo
        weeks.update(range(lo, hi + 1))
    if default_kind == 'WEEK' or has_week_marker:
        text = _mask(text, r'第?\s*(\d{1,3})\s*周?\s*(?:-|—|–|~|至|到)\s*第?\s*(\d{1,3})\s*周?', _num_range)

    # ⑤ 周次单值：'第7周' / '9周'
    def _num_point(m):
        d = int(m.group(1))
        if d <= MAX_WEEK_NO:
            weeks.add(d)
    text = _mask(text, r'第?\s*(\d{1,3})\s*周', _num_point)

    # ⑥ 周次单值（中文数字）：'第七周' / '第八周'
    def _cn_point(m):
        d = _cn_number(m.group(1))
        if d is not None and d <= MAX_WEEK_NO:
            weeks.add(d)
    text = _mask(text, r'第\s*([%s]+)\s*周' % _CN_NUM_CHARS, _cn_point)

    # ⑦ 裸星期：'第X'（后面不跟「周」）→ '第一' / '第二'
    def _bare_d(m):
        weekdays.add(_WEEKDAY_BY_CN[m.group(1)])
    text = _mask(text, r'第([一二三四五六日天])(?![周])', _bare_d)

    # ⑧ 裸中文数字串（不跟「周」）→ 星期列举：'周一 三 四' 里的 三/四
    #    注意：裸「天」不算星期（'今天' 不是周日）；只有带「周/星期/礼拜」前缀的才算，
    #    那部分已在 ② 里处理。
    def _bare_cn(m):
        for ch in m.group(0):
            if ch == '天':
                continue
            day = _WEEKDAY_BY_CN.get(ch)
            if day:
                weekdays.add(day)
    text = _mask(text, r'[一二三四五六日天]+(?![周])', _bare_cn)

    # ⑨ 剩余阿拉伯数字：周次列当周次；时间列只在 ≤7 且无周次标记时当星期
    for token in re.split(r'[,;、/\s]+', text):
        m = re.search(r'(?<!\d)(\d{1,3})(?!\d)', token)
        if not m:
            continue
        d = int(m.group(1))
        if default_kind == 'WEEK':
            if d <= MAX_WEEK_NO:
                weeks.add(d)
        elif d <= 7:
            weekdays.add(WEEKDAY_NAMES[d - 1])

    # ⑩ '全学期' / '本学期' / '整学期' / '一学期' → 覆盖统计窗口周（裁定 R4）
    if re.search(r'(?:全|本|整|一)学期', text) and window_week is not None:
        weeks.add(int(window_week))
    return weeks, weekdays


def resolve_pairs(week_text, time_text, window_week: Optional[int] = None):
    """把「免勤周次」与「免勤时间」两列合成 (周次, 星期) 事件集合。

    **段内配对优先**：一个单元格内同时出现周次与星期时（真实数据存在
    '6-8周周一 三 四，8-10周周四' 这类写法），按段内配对；
    只有当两列各自单调（周次列只有周次、时间列只有星期）时才跨列笛卡尔积。
    混排单元格**绝不跨列外推**——外推会造出没人申请过的「第9-10周 × 周一」。

    返回 (pairs:set[(int,str)], warnings:list[str])。
    """
    segs = []
    for text, kind in ((week_text, 'WEEK'), (time_text, 'TIME')):
        raw = normalize_text(text)
        if not raw:
            continue
        for seg in re.split(r'[,;、]+', raw):
            if not seg.strip():
                continue
            w, d = _segment_facts(seg, kind, window_week)
            if w or d:
                segs.append((w, d, kind))
    if not segs:
        return set(), ['免勤周次与免勤时间均为空，无法定位事件']

    joint = [s for s in segs if s[0] and s[1]]
    if not joint:
        weeks = set()
        days = set()
        for w, d, _k in segs:
            weeks |= w
            days |= d
        if not weeks or not days:
            return set(), ['周次或星期解析为空']
        return {(w, d) for w in weeks for d in days}, []

    pairs = set()
    for w, d, _k in joint:
        pairs |= {(x, y) for x in w for y in d}

    warns = []
    covered_days = {y for _x, y in pairs}
    for w, d, kind in segs:
        if w and not d:
            same = [s for s in segs if s[2] == kind and s[1]]
            if same:
                days = set()
                for _w, dd, _k in same:
                    days |= dd
                pairs |= {(x, y) for x in w for y in days}
            else:
                warns.append('周次段缺星期且同列无可继承值：%s' % sorted(w))
        elif d and not w:
            same = [s for s in segs if s[2] == kind and s[0]]
            if same:
                weeks = set()
                for ww, _dd, _k in same:
                    weeks |= ww
                pairs |= {(x, y) for x in weeks for y in d}
            elif not (d & covered_days):
                warns.append('星期段缺周次且同列无可继承值：%s' % sorted(d))
    return pairs, warns


def parse_weeks(text, week_no: Optional[int] = None) -> set:
    """解析免勤周次，返回周号集合。

    week_no：统计窗口周次（默认 7）。'全学期' 写法覆盖窗口内任意一周，
    按窗口周次计（裁定 R4）——**不得**当作「周次无法识别」丢弃。

    >>> sorted(parse_weeks('第7、8、9周'))
    [7, 8, 9]
    >>> sorted(parse_weeks('第6-7，9-13周'))
    [6, 7, 9, 10, 11, 12, 13]
    >>> sorted(parse_weeks('第七周，第八周'))
    [7, 8]
    >>> sorted(parse_weeks('第7—12周'))
    [7, 8, 9, 10, 11, 12]
    >>> sorted(parse_weeks('6-10周'))
    [6, 7, 8, 9, 10]
    >>> sorted(parse_weeks('第六-十七周'))
    [6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]
    >>> sorted(parse_weeks('全学期'))
    [7]
    >>> sorted(parse_weeks('第一 三 四'))
    []
    """
    raw = normalize_text(text)
    if not raw:
        return set()
    window = int(week_no) if week_no is not None else DEFAULT_WEEK_NO
    weeks = set()
    for seg in re.split(r'[,;、]+', raw):
        if not seg.strip():
            continue
        w, _d = _segment_facts(seg, 'WEEK', window)
        weeks |= w
    return weeks


def parse_week_span(text) -> Optional[set]:
    """从 '周一到周五' 这类区间写法展开成星期集合（不是枚举式）。"""
    raw = normalize_text(text)
    m = re.search(r'周([一二三四五六日天])\s*(?:到|至|-|~|—)\s*周?([一二三四五六日天])', raw)
    if not m:
        return None
    start = _WEEKDAY_BY_CN.get(m.group(1))
    end = _WEEKDAY_BY_CN.get(m.group(2))
    if not start or not end:
        return None
    lo, hi = WEEKDAY_NAMES.index(start), WEEKDAY_NAMES.index(end)
    if lo > hi:
        lo, hi = hi, lo
    return set(WEEKDAY_NAMES[lo:hi + 1])


def parse_weekdays(text) -> set:
    """解析免勤时间里的星期，返回 '周一'…'周日' 集合。

    >>> sorted(parse_weekdays('周一、周三、周四、周五'))
    ['周五', '周四', '周三', '周一']
    >>> sorted(parse_weekdays('周一 ，周三 ，周四'))
    ['周三', '周四', '周一']
    >>> sorted(parse_weekdays('周一到周五'))
    ['周五', '周四', '周三', '周二', '周一']
    >>> sorted(parse_weekdays('周一周二周三周四周五'))
    ['周五', '周四', '周三', '周二', '周一']
    >>> sorted(parse_weekdays('周一、二、四、五'))
    ['周五', '周二', '周四', '周一']
    """
    raw = normalize_text(text)
    if not raw:
        return set()
    # 「星期天 / 周天」先归一成「周日」，避免去掉「周」前缀后把
    # 「今天」这类普通词里的「天」误判成星期日。
    raw = raw.replace('星期天', '周日').replace('周天', '周日')
    days = set()
    for seg in re.split(r'[,;、]+', raw):
        if not seg.strip():
            continue
        _w, d = _segment_facts(seg, 'TIME')
        days |= d
    return days


# ---------------------------------------------------------------- 校区 / 学院 / 班级

def normalize_campus(text, custom_map: Optional[dict] = None) -> str:
    """校区归一。'雁山校园'/'燕山校区' 都是 '雁山校区' 的手写变体。

    >>> normalize_campus('燕山校区')
    '雁山'
    >>> normalize_campus('雁山校园')
    '雁山'
    >>> normalize_campus('屏风')
    '屏风'
    """
    raw = strip_spaces(text)
    if not raw:
        return '待确认'
    if custom_map:
        if raw in custom_map:
            return custom_map[raw]
        for k, v in custom_map.items():
            if k and k in raw:
                return v
    if '屏风' in raw:
        return '屏风'
    if '雁山' in raw or '燕山' in raw:
        return '雁山'
    if '雁门' in raw:
        return '雁门'
    return '待确认'


def parse_college_label(text):
    """解析 '地球科学学院（460人）' → ('地球科学学院', 460)。

    括号全半角都支持；缺人数时 total 为 None。
    """
    raw = normalize_text(text)
    if not raw:
        return '', None
    m = re.search(r'^(.*?)\s*[（(]\s*(\d+)\s*人\s*[)）]\s*$', raw)
    if m:
        return m.group(1).strip(), int(m.group(2))
    m = re.search(r'[（(]\s*(\d+)\s*人\s*[)）]', raw)
    if m:
        return raw[:m.start()].strip(), int(m.group(1))
    return raw, None


def normalize_class(text, custom_map: Optional[dict] = None) -> str:
    """班级名归一，用于把集体申请与个人申请对应到同一把钥匙。

    '应用统计学2026-1班' 与 '应用统计学26-1' 归一为同一 key。
    """
    raw = strip_spaces(text)
    if not raw:
        return ''
    if custom_map and raw in custom_map:
        raw = strip_spaces(custom_map[raw])
    raw = re.sub(r'全体学生', '', raw)
    raw = re.sub(r'[（(]?\s*共\s*\d+\s*人\s*[)）]?', '', raw)
    if custom_map and raw in custom_map:
        raw = strip_spaces(custom_map[raw])
    raw = raw.replace('班', '').replace('级', '')
    raw = re.sub(r'20(\d{2})', r'\1', raw)      # 2026 -> 26
    raw = raw.replace('-', '').replace('_', '').replace('/', '')
    res = raw.strip()
    if custom_map and res in custom_map:
        res = normalize_class(custom_map[res], None)
    return res


def extract_group_size(text) -> Optional[int]:
    """从 '应用统计学26-1班全体学生（共43人）' 里取出 43。"""
    raw = normalize_text(text)
    m = re.search(r'共\s*(\d+)\s*人', raw)
    return int(m.group(1)) if m else None


def is_group_application(*texts) -> bool:
    """判断是否为集体申请（无学号的整班免勤）。"""
    for text in texts:
        if text and ('全体学生' in str(text) or '全班' in str(text)):
            return True
    return False


# ---------------------------------------------------------------- 数值

def parse_int(value) -> Optional[int]:
    """宽松数值解析。空、'免勤'、'无' 等一律返回 None（视为缺失，不是 0）。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value != value:   # NaN
            return None
        return int(value)
    raw = strip_spaces(value)
    if not raw:
        return None
    m = re.search(r'-?\d+(?:\.\d+)?', raw)
    if not m:
        return None
    try:
        return int(round(float(m.group(0))))
    except ValueError:
        return None


def parse_session(text) -> Optional[str]:
    """'早' → '早自习'；'晚' → '晚自习'。"""
    raw = strip_spaces(text)
    if not raw:
        return None
    if '早' in raw:
        return '早自习'
    if '晚' in raw:
        return '晚自习'
    return None


def normalize_room(text) -> str:
    """教室号归一：去空白、去 '教室' 后缀、保留前导零原样。"""
    raw = strip_spaces(text)
    if not raw:
        return ''
    raw = raw.replace('教室', '')
    return raw


def pct(value: Optional[float]) -> Optional[float]:
    """统一成 0~1 的比例值。"""
    if value is None:
        return None
    return float(value)


# ---------------------------------------------------------------- 班级特征解析与教室匹配

MAJOR_ABBR_MAP = {
    # 地球科学
    "资勘": "资源勘查工程", "资源勘查": "资源勘查工程",
    "勘技": "勘查技术与工程", "勘查技术": "勘查技术与工程",
    "地信": "地球信息科学与技术", "地球信息": "地球信息科学与技术", "地求信息": "地球信息科学与技术",
    "地工": "地质工程", "地质": "地质工程",
    "地化": "地球化学",
    # 环工
    "环工": "环境工程",
    "给排水": "给排水科学与工程", "水务": "给排水科学与工程",
    "环科": "环境科学",
    # 计信 / 软工 / 数统
    "软工": "软件工程",
    "计科": "计算机科学与技术", "计算机": "计算机科学与技术",
    "物联": "物联网工程", "物联网": "物联网工程",
    "信安": "信息安全",
    "大数据": "数据科学与大数据技术", "数科": "数据科学与大数据技术",
    "应统": "应用统计学", "统计": "应用统计学",
    "信计": "信息与计算科学", "信算": "信息与计算科学",
    # 土木
    "土木": "土木工程",
    # 化工
    "化工": "化学工程与工艺",
    "应化": "应用化学",
    "生工": "生物工程",
}


def parse_class_feature(text: str) -> tuple[str, str]:
    """从班级文本中提取 (规范专业名, 班级特征)。

    例：
        '地球化学2026-1班' -> ('地球化学', '26-1')
        '地球化学2026拔尖班' -> ('地球化学', '26-拔尖')
        '资勘东盟2026-1班' -> ('资源勘查工程', '26-1-东盟')
        '资勘26_2' -> ('资源勘查工程', '26-2')
        '地球信息科学与技术2026-1半' -> ('地球信息科学与技术', '26-1')
    """
    if not text:
        return ("", "")
    raw = str(text).strip().replace(" ", "").replace("　", "")
    raw = raw.replace("半", "班").replace("—", "-").replace("–", "-").replace("_", "-")

    # 统一 2026 -> 26, 2025 -> 25 等
    raw = re.sub(r'20(\d{2})', r'\1', raw)

    special = ""
    for spec in ["拔尖", "创新", "实验", "东盟", "国际", "专升本", "卓越"]:
        if spec in raw:
            special = spec
            break

    class_num_match = re.search(r'(\d{2})[-_]?(\d+|[A-Za-z]+)', raw)
    class_id = ""
    if class_num_match:
        yr = class_num_match.group(1)
        num = class_num_match.group(2)
        class_id = f"{yr}-{num}"
    else:
        yr_match = re.search(r'(\d{2})', raw)
        if yr_match:
            class_id = yr_match.group(1)

    if special:
        class_id = f"{class_id}-{special}" if class_id else special

    major_part = re.sub(r'\d+', '', raw)
    major_part = re.sub(r'[班级\(\)（）\-_]', '', major_part)
    for spec in ["拔尖", "创新", "实验", "东盟", "国际", "专升本", "卓越", "全体学生", "全班"]:
        major_part = major_part.replace(spec, "")

    canonical_major = major_part
    for abbr, full in MAJOR_ABBR_MAP.items():
        if major_part == abbr or major_part.startswith(abbr) or abbr in major_part:
            canonical_major = full
            break

    return (canonical_major, class_id)


def match_class_to_room_classes(student_class: str, room_classes_text: str) -> bool:
    """判断学生的班级是否精准命中了明细表中该教室的分配专业班级。

    明细表中 classes_text 常见如：'地球信息2026-1班、勘技2026-1班、勘技2026-2班'
    学生班级如：'地球信息科学与技术2026-1'
    """
    if not student_class or not room_classes_text:
        return False
    stu_major, stu_cid = parse_class_feature(student_class)
    if not stu_major or not stu_cid:
        return False

    room_items = re.split(r'[,;、\s]+', room_classes_text)
    for item in room_items:
        if not item.strip():
            continue
        r_major, r_cid = parse_class_feature(item)
        if not r_major or not r_cid:
            continue
        # 班级编号必须完全一致
        if stu_cid == r_cid:
            # 专业名称完全一致，或者互相包含/简称命中
            if stu_major == r_major or stu_major in r_major or r_major in stu_major:
                return True
    return False

