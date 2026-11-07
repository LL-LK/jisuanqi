# -*- coding: utf-8 -*-
"""隐私扫描器：打包与推送前必须跑一遍（V2 双闸门的唯一实现）。

检查五类泄露：
  1. 绝对路径（本机盘符、用户名、/mnt/data、/home/、Python 安装路径等）
  2. 真实学号（13 位数字串）
  3. 真实姓名与真实免勤数据（与本地真实免勤表交叉比对，不靠正则）
  4. 真实 Excel 数据文件被放进项目目录 / 产物 / 仓库
  5. 被否口径的「旋钮」复活（cap_pct / apply_cap / 临时请假封顶…）

三种扫描面：
  --root <目录>              源码树（文本文件）
  --artifact <目录>          产物目录（**含二进制 EXE** 与 xlsx 单元格）
  --push-repo <clone>        git 分支（**提交内容 + 完整 patch 历史**）
  --push-branch <ref>        分支引用，默认 origin/desktop

用法：
    python -X utf8 tools/privacy_scan.py                          # 扫源码树
    python -X utf8 tools/privacy_scan.py --out 报告.md
    python -X utf8 tools/privacy_scan.py --data "<真实免勤表.xlsx>"
    python -X utf8 tools/privacy_scan.py --artifact 发布版 --out 产物报告.md
    python -X utf8 tools/privacy_scan.py --push-repo <clone> \
        --push-branch origin/desktop --out push回执.md
"""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import subprocess
import sys
import tempfile
import zipfile
from typing import Dict, List, Set, Tuple

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WORKSPACE = os.path.dirname(ROOT)

TEXT_EXT = {'.py', '.bat', '.ps1', '.md', '.txt', '.json', '.yml', '.yaml',
            '.spec', '.cfg', '.ini', '.toml', '.csv', '.html', '.css', '.js'}
SKIP_DIR = {'__pycache__', '.git', '.pytest_cache', 'dist', 'dist_linux', 'build', '.hypothesis',
            '.buildenv', '.venv', 'venv', '_synthetic', '发布版', '发布包', '_work', 'data_sample',
            'node_modules', 'site-packages', '.agents', 'tests_e2e'}

# 绝对路径与身份痕迹
# 第三列 text_only=True 表示该模式**只对文本内容**生效：
# 通用盘符正则放进 EXE/ZIP 等压缩字节里会随机命中（压缩流里 3 字节就可能出现 "F:\x"），
# 因此二进制只保留高信号特征串（具体盘符+具体目录名）。
PATH_PATTERNS = [
    (r'[A-Za-z]:\\\\?[\w一-鿿][^\s"\'<>|]*', 'Windows 盘符绝对路径', True),
    (r'/mnt/[A-Za-z0-9_./-]+', 'Linux /mnt 挂载路径', True),
    (r'/home/[A-Za-z0-9_.-]+', 'Linux 用户目录', True),
    (r'/Users/[A-Za-z0-9_.-]+', 'macOS 用户目录', True),
    (r'C:\\\\?Users\\\\?22140', '本机用户名片段', False),
    (r'python\.data', '本机 Python 安装路径', False),
    (r'三早一晚\s*ruanjian', '本机工作区绝对路径片段', False),
    (r'F:\\\\?三早一晚', '本机工作区绝对路径片段', False),
    (r'G:\\\\?python\.data', '本机 Python 安装路径', False),
    (r'/mnt/data', 'Linux /mnt 挂载路径', False),
    (r'/home/[A-Za-z0-9_.-]+', 'Linux 用户目录', False),
]
# 13 位数字串：两侧不能是十六进制字符，否则会命中 git sha / 文件哈希里的连续数字
SID_PATTERN = re.compile(r'(?<![0-9A-Fa-f])\d{13}(?![0-9A-Fa-f])')

# 被否口径的「旋钮」：本项目口径是免勤不封顶，代码与词典里不允许留可切换的封顶开关。
# 扫描范围：.py 源码与界面文案 + **xlsx 单元格内容**（V2 补丁：旋钮住在词典里）；
# 不扫 .md 文档——文档里作为口径说明出现的字样不算命中，否则口径文档会被误杀。
REJECTED_KNOBS = ('cap_pct', 'apply_cap', '临时请假封顶', '免勤封顶比例', '免勤封顶')
SCAN_EXT = {'.py'}
SCAN_XLSX_CELLS = True

# ---- 第 4 类闸门：随包「免勤理由分类」sheet 白名单逐格比对（Q3 裁定）----
# 裁定原文：词典「免勤理由分类」只收通用类别与中性词，禁止人名、学号、班级名、
# 学院名、具体到个人的课程安排、任何具体诊断或症状描述；健康医疗类只允许
# 类别名与中性通用词；源理由文本只用于当次分类、不落盘、不随包分发。
# 实现要点：白名单是**逐格**的——关键词、分类、说明三列分别白名单化；
# 出现白名单之外的任何一行 ⇒ 判高危（阻断打包与发布），因为新增行可能含
# 具体诊断/症状、人名或班级。判定只许一处实现：三个扫描面共用本处。
REASON_SHEET = '免勤理由分类'
ALLOWED_CATEGORIES = ('课务学业', '社团竞赛训练', '健康医疗', '其他')
ALLOWED_REASON_KEYWORDS = ('有选修课', '军事理论课冲突', '生病', '就医', '住院', '治疗')
ALLOWED_REASON_NOTES = ('通用课程/事务类别，不含个人指向',
                         '全校通用课程冲突',
                         '通用健康情形词，不含具体诊断或症状')
REASON_VIOLATION = '被否口径越界（理由分类白名单外）'

# 占位名：这些是表格里的"占位写法"，不是真实个人姓名，扫描时应排除
PLACEHOLDER_NAMES = {'全班', '全体', '全部', '集体', '无', '待定', '同学',
                      '学生', '未填写', '同上', '略', '/'}


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


def iter_text_files(root: str, skip_paths=()):
    """遍历待扫描文本文件。

    skip_paths：本次运行正在写出的报告文件（避免自指）。
    注意：**不再按文件名跳过「隐私扫描报告.md」**——报告本身会随 EXE 分发，
    必须与其它文件一样接受扫描；报告里的路径必须写成脱敏形式。
    """
    skip_abs = {os.path.abspath(p) for p in skip_paths if p}
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIR]
        for name in filenames:
            path = os.path.join(dirpath, name)
            if os.path.splitext(name)[1].lower() not in TEXT_EXT:
                continue
            if os.path.basename(path) == 'privacy_scan.py':
                continue          # 扫描器自身含模式定义，避免自指误报
            if os.path.abspath(path) in skip_abs:
                continue          # 本次正在写出的报告
            yield path


def load_real_personal(data_path: str):
    """从真实免勤表里取出学号与姓名，用于交叉比对。"""
    sids: Set[str] = set()
    names: Set[str] = set()
    try:
        from openpyxl import load_workbook
    except ImportError:
        return sids, names
    if not data_path or not os.path.exists(data_path):
        return sids, names
    wb = load_workbook(data_path, data_only=True, read_only=True)
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows(min_row=3, values_only=True):
            if not row:
                continue
            sid = str(row[1]).strip() if len(row) > 1 and row[1] else ''
            name = str(row[3]).strip() if len(row) > 3 and row[3] else ''
            if sid and re.fullmatch(r'\d{13}', sid):
                sids.add(sid)
            if name and len(name) >= 2 and len(name) <= 4 and re.fullmatch(r'[一-鿿]{2,4}', name):
                if name in PLACEHOLDER_NAMES:
                    continue
                names.add(name)
    wb.close()
    return sids, names


def _cn_name_tokens(text: str, real_names) -> list:
    """从文本里抽出所有 2~4 字的连续中文片段，与真实姓名库求交集。

    比「逐个姓名做子串搜索」快若干个数量级（2601 次子串搜索 vs 一次抽词）。
    """
    if not real_names:
        return []
    cands = set()
    for m in re.finditer(r'[一-鿿]+', text):
        s = m.group(0)
        if len(s) > 60:
            s = s[:60]
        L = len(s)
        for i in range(L):
            for j in range(i + 2, min(i + 5, L + 1)):
                cands.add(s[i:j])
    return sorted(cands & real_names)


def _scan_text(content: str, rel: str, real_sids, real_names, findings,
               knobs: bool = False, skip_self: bool = False, binary: bool = False):
    """对一段文本做四类检查。

    skip_self=True  用于扫描器自身文件：它按定义必然包含全部路径模式串与
                    旋钮词表，自指命中必须排除，否则闸门会永远「不通过」。
    binary=True     用于 EXE/ZIP 等压缩内容：通用盘符正则在此不适用（随机命中），
                    只保留高信号特征串。
    """
    if not skip_self:
        for pattern, label, text_only in PATH_PATTERNS:
            if binary and text_only:
                continue
            for m in re.finditer(pattern, content):
                findings.append({'kind': label, 'file': rel,
                                 'detail': m.group(0)[:120],
                                 'line': content[:m.start()].count('\n') + 1})
    for m in SID_PATTERN.finditer(content):
        value = m.group(0)
        if value.startswith('99'):
            continue                      # 合成测试学号：99 开头是约定虚构号段
        findings.append({'kind': '13 位数字串（疑似学号）', 'file': rel,
                         'detail': value,
                         'line': content[:m.start()].count('\n') + 1})
    if real_sids:
        hit = sorted(s for s in real_sids if s in content)
        if hit:
            findings.append({'kind': '真实学号', 'file': rel,
                             'detail': '命中 %d 个真实学号' % len(hit), 'line': 0})
    if real_names:
        # 姓名交叉比对是**子串**匹配：常见词（如「XX基线」里的「XX」）可能恰好
        # 是某位学生的姓名。因此按严重度分级：
        #   · 同一行/同一区域同时出现**真实学号** ⇒ 高危（真实数据行形态）
        #   · 只出现姓名            ⇒ 低危（常见词，需人工确认）
        for n in _cn_name_tokens(content, real_names):
            idx = content.find(n)
            line_start = content.rfind('\n', 0, idx) + 1
            line_end = content.find('\n', idx)
            line = content[line_start: line_end if line_end > 0 else len(content)]
            window = content[max(0, idx - 200): idx + 200]
            with_sid = any(s in line or s in window for s in real_sids)
            findings.append({'kind': '真实姓名（与学号共现）' if with_sid
                             else '真实姓名（疑似常见词，需人工确认）',
                             'file': rel, 'detail': n,
                             'line': content[:idx].count('\n') + 1})
    if knobs and not skip_self:
        for knob in REJECTED_KNOBS:
            if knob in content:
                findings.append({'kind': '被否口径旋钮', 'file': rel,
                                 'detail': knob,
                                 'line': content[:content.index(knob)].count('\n') + 1})


def scan(root: str, data_path: str, skip_paths=()):
    findings: List[Dict] = []
    synthetic_sids = {}
    real_sids, real_names = load_real_personal(data_path)
    files = list(iter_text_files(root, skip_paths))

    # 源码树里的表格也要逐单元格扫描（V2 补丁：被否口径旋钮住在词典 xlsx 里，
    # 只扫文本文件会漏掉它——打包闸门第 7 步走的正是这个扫描面）
    xlsx_hits: List[Dict] = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIR]
        for name in filenames:
            if os.path.splitext(name)[1].lower() not in ('.xlsx', '.xls'):
                continue
            path = os.path.join(dirpath, name)
            rel = os.path.relpath(path, root)
            xlsx_hits += scan_xlsx_cells(path, rel, real_sids, real_names)

    for path in files:
        rel = os.path.relpath(path, root)
        try:
            content = open(path, encoding='utf-8', errors='replace').read()
        except Exception:
            continue
        knobs = os.path.splitext(path)[1].lower() in SCAN_EXT
        _scan_text(content, rel, real_sids, real_names, findings, knobs=knobs)
        for m in SID_PATTERN.finditer(content):
            if m.group(0).startswith('99'):
                v = m.group(0)
                synthetic_sids[v] = synthetic_sids.get(v, 0) + 1
    return findings + xlsx_hits, files, real_sids, real_names, synthetic_sids


def scan_xlsx_cells(path: str, rel: str, real_sids, real_names) -> List[Dict]:
    """逐单元格扫描 xlsx：真实学号 / 真实姓名 / 绝对路径 / 被否口径旋钮。"""
    findings: List[Dict] = []
    try:
        from openpyxl import load_workbook
    except ImportError:
        return findings
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:
        findings.append({'kind': 'xlsx 无法解析', 'file': rel,
                         'detail': str(exc)[:100], 'line': 0})
        return findings
    # 词典必须有「免勤理由分类」表：缺表 ⇒ 分类列会回落到默认「未分类」，违反 Q3
    if '词典' in os.path.basename(path) and REASON_SHEET not in wb.sheetnames:
        findings.append({'kind': REASON_VIOLATION, 'file': rel,
                         'detail': '词典缺少「%s」表，分类列将回落到默认「未分类」' % REASON_SHEET,
                         'line': 0})
    for sn in wb.sheetnames:
        ws = wb[sn]
        for ri, row in enumerate(ws.iter_rows(values_only=True), 1):
            for ci, cell in enumerate(row, 1):
                if cell is None:
                    continue
                s = str(cell)
                where = '%s[%s!R%dC%d]' % (rel, sn, ri, ci)
                if sn == REASON_SHEET and ri >= 2:
                    # ---- 第 4 类闸门：逐格白名单 ----
                    if ci == 1 and s.strip() not in ALLOWED_REASON_KEYWORDS:
                        findings.append({'kind': REASON_VIOLATION, 'file': where,
                                         'detail': '关键词不在白名单：%s' % s[:40], 'line': 0})
                    elif ci == 2 and s.strip() not in ALLOWED_CATEGORIES:
                        findings.append({'kind': REASON_VIOLATION, 'file': where,
                                         'detail': '分类不在白名单：%s' % s[:40], 'line': 0})
                    elif ci == 3 and s.strip() and s.strip() not in ALLOWED_REASON_NOTES:
                        findings.append({'kind': REASON_VIOLATION, 'file': where,
                                         'detail': '说明不在白名单：%s' % s[:40], 'line': 0})
                for knob in REJECTED_KNOBS:
                    if knob in s:
                        findings.append({'kind': '被否口径旋钮（xlsx 单元格）',
                                         'file': where, 'detail': knob, 'line': 0})
                for m in SID_PATTERN.finditer(s):
                    if m.group(0).startswith('99'):
                        continue
                    findings.append({'kind': '13 位数字串（xlsx 单元格）',
                                     'file': where, 'detail': m.group(0), 'line': 0})
                if real_sids and s in real_sids:
                    findings.append({'kind': '真实学号', 'file': where,
                                     'detail': s, 'line': 0})
                if real_names and s in real_names:
                    findings.append({'kind': '真实姓名（疑似常见词，需人工确认）',
                                     'file': where, 'detail': s, 'line': 0})
                for pattern, label, text_only in PATH_PATTERNS:
                    if re.search(pattern, s):
                        findings.append({'kind': label + '（xlsx 单元格）',
                                         'file': where, 'detail': s[:80], 'line': 0})
    wb.close()
    return findings


def scan_data_files(root: str, data_path: str):
    """项目目录里是否混入了真实数据文件。"""
    real_hashes = set()
    if data_path and os.path.exists(data_path):
        real_hashes.add(sha256(data_path))
    for name in os.listdir(WORKSPACE):
        path = os.path.join(WORKSPACE, name)
        if os.path.isfile(path) and os.path.splitext(name)[1].lower() in ('.xlsx', '.xls', '.csv'):
            real_hashes.add(sha256(path))
    hits = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIR]
        for name in filenames:
            if os.path.splitext(name)[1].lower() not in ('.xlsx', '.xls', '.csv'):
                continue
            path = os.path.join(dirpath, name)
            rel = os.path.relpath(path, root)
            digest = sha256(path)
            kind = '真实数据文件副本' if digest in real_hashes else '项目内 Excel（需人工确认）'
            hits.append((rel, kind, os.path.getsize(path)))
    return hits


# ---------------------------------------------------------------- 产物扫描
def scan_artifact(root: str, data_path: str) -> Tuple[List[Dict], List[Tuple[str, str, int]]]:
    """扫描发布目录：含二进制 EXE 的逐字节反查 + xlsx 单元格。"""
    findings: List[Dict] = []
    real_sids, real_names = load_real_personal(data_path)
    real_hashes = set()
    if data_path and os.path.exists(data_path):
        real_hashes.add(sha256(data_path))
    for name in os.listdir(WORKSPACE):
        p = os.path.join(WORKSPACE, name)
        if os.path.isfile(p) and os.path.splitext(name)[1].lower() in ('.xlsx', '.xls', '.csv'):
            real_hashes.add(sha256(p))

    files_report = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in ('__pycache__', '.git')]
        for name in filenames:
            path = os.path.join(dirpath, name)
            rel = os.path.relpath(path, root)
            ext = os.path.splitext(name)[1].lower()
            if ext in ('.xlsx', '.xls'):
                findings += scan_xlsx_cells(path, rel, real_sids, real_names)
                files_report.append((rel, '表格（已逐单元格扫描）', os.path.getsize(path)))
                try:
                    with zipfile.ZipFile(path) as zf:
                        for member in zf.namelist():
                            _scan_text(zf.read(member).decode('utf-8', 'replace'),
                                       '%s!%s' % (rel, member), real_sids, real_names,
                                       findings)
                except Exception:
                    pass
                continue
            with open(path, 'rb') as fh:
                blob = fh.read()
            if ext in ('.xls', '.csv'):
                digest = hashlib.sha256(blob).hexdigest()
                kind = ('真实数据文件副本' if digest in real_hashes
                        else '表格（需人工确认）')
                files_report.append((rel, kind, len(blob)))
            else:
                files_report.append((rel, '二进制/文档（已逐字节扫描）', len(blob)))
            # 文本文件走完整四类检查；EXE/ZIP 等压缩内容只查高信号特征串
            is_text = ext in TEXT_EXT
            _scan_text(blob.decode('utf-8', 'replace'), rel, real_sids, real_names,
                       findings, binary=not is_text)
            # UTF-16LE 通道：只补查路径特征（中文在 UTF-16LE 下会错位，
            # 因此不复用完整检查，避免重复计入）。通用盘符正则只对小文件生效——
            # 文本类文件才可能是 UTF-16 编码，50MB 级二进制走全量正则纯属空转。
            u16 = blob.decode('utf-16-le', 'replace')[1::2]
            for pattern, label, text_only in PATH_PATTERNS:
                if text_only and len(blob) > 8 * 1024 * 1024:
                    continue
                for m in re.finditer(pattern, u16):
                    findings.append({'kind': label, 'file': rel,
                                     'detail': m.group(0)[:120], 'line': 0})
    return findings, files_report


# ---------------------------------------------------------------- 推送侧扫描
def _git(repo: str, args: List[str]) -> str:
    p = subprocess.run(['git', '-C', repo, '-c', 'core.quotePath=false'] + args,
                       capture_output=True, text=True, encoding='utf-8',
                       errors='replace')
    return p.stdout or ''


def scan_push(repo: str, branch: str, data_path: str) -> Tuple[List[Dict], List[str], List[str]]:
    """扫描分支的提交内容与完整 patch 历史。"""
    findings: List[Dict] = []
    real_sids, real_names = load_real_personal(data_path)
    tree = [t for t in _git(repo, ['ls-tree', '-r', '--name-only', branch]).splitlines() if t.strip()]

    tmp = tempfile.mkdtemp(prefix='privacyscan_push_')
    archive = os.path.join(tmp, 'branch.zip')
    with open(archive, 'wb') as fh:
        subprocess.run(['git', '-C', repo, 'archive', '--format=zip', branch],
                       stdout=fh, stderr=subprocess.DEVNULL)
    work = os.path.join(tmp, 'x')
    try:
        with zipfile.ZipFile(archive) as zf:
            zf.extractall(work)
    except Exception:
        work = ''

    if work:
        for dirpath, _dirs, filenames in os.walk(work):
            for name in filenames:
                path = os.path.join(dirpath, name)
                rel = os.path.relpath(path, work).replace('\\', '/')
                ext = os.path.splitext(name)[1].lower()
                is_scanner = os.path.basename(rel) == 'privacy_scan.py'
                if ext in ('.xlsx', '.xls'):
                    findings += scan_xlsx_cells(path, rel, real_sids, real_names)
                    continue
                try:
                    content = open(path, encoding='utf-8', errors='replace').read()
                except Exception:
                    continue
                _scan_text(content, rel, real_sids, real_names, findings,
                           knobs=(ext in SCAN_EXT), skip_self=is_scanner)

    base = _git(repo, ['merge-base', branch, 'origin/main']).strip()
    hist = _git(repo, ['log', '-p', '--no-color',
                       ('%s..%s' % (base, branch)) if base else branch])
    # patch 里必然出现扫描器自身的模式定义行：按文件分段剔除，避免自指误报
    kept, skip = [], False
    for line in hist.splitlines():
        if line.startswith('+++ b/'):
            skip = os.path.basename(line[6:].strip()) == 'privacy_scan.py'
        if not skip:
            kept.append(line)
    hist_clean = '\n'.join(kept)
    _scan_text(hist_clean, '提交历史(patch)', real_sids, real_names, findings)
    try:
        import shutil
        shutil.rmtree(tmp, ignore_errors=True)
    except Exception:
        pass
    return findings, tree, []


# ---------------------------------------------------------------- 报告
HARD_KINDS = ('真实学号', '真实数据文件副本',
              '被否口径旋钮', '被否口径旋钮（xlsx 单元格）',
              '被否口径越界（理由分类白名单外）',
              '真实姓名（与学号共现）',
              'Windows 盘符绝对路径', '本机用户名片段', '本机 Python 安装路径',
              '本机工作区绝对路径片段', 'Linux /mnt 挂载路径',
              'Linux 用户目录', 'macOS 用户目录')
# 低危：姓名子串命中但无学号共现（常见词误报），不阻断，但需人工确认
SOFT_KINDS = ('真实姓名（疑似常见词，需人工确认）',
              '13 位数字串（疑似学号）', '13 位数字串（xlsx 单元格）')


def _table(findings, mask_names: bool = True) -> List[str]:
    out = ['| 类型 | 文件 | 行 | 命中内容 |', '|---|---|---:|---|']
    for item in findings:
        detail = str(item['detail'])
        if mask_names and item['kind'].startswith('真实姓名'):
            # 报告随 EXE / 发布包分发 ⇒ 命中姓名**不回显原文**，只留位置供人工复核
            detail = '（姓名命中，已脱敏；需原文用 --show-names 本地查看）'
        out.append('| %s | `%s` | %s | `%s` |'
                   % (item['kind'], item['file'], item['line'],
                      detail.replace('|', '\\|')))
    return out


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', default=ROOT)
    parser.add_argument('--data', default=None, help='真实免勤表路径（对照数据源）')
    parser.add_argument('--out', default=None, help='报告输出路径')
    parser.add_argument('--artifact', default=None, help='产物目录（含二进制 EXE）')
    parser.add_argument('--push-repo', default=None, help='git clone 路径（推送侧扫描）')
    parser.add_argument('--push-branch', default='origin/desktop', help='分支引用')
    parser.add_argument('--show-names', action='store_true',
                        help='本地排障时回显命中姓名原文（默认脱敏：报告随 EXE 分发）')
    args = parser.parse_args()

    data_path = args.data
    if not data_path:
        for name in os.listdir(WORKSPACE):
            if '免勤人员统计表' in name:
                data_path = os.path.join(WORKSPACE, name)
                break

    lines: List[str] = ['# 隐私扫描报告', '']
    all_findings: List[Dict] = []
    real_sids, real_names = load_real_personal(data_path)

    # ---- 源码树 ----
    findings, files, real_sids, real_names, synthetic_sids = scan(
        args.root, data_path, skip_paths=[args.out] if args.out else [])
    data_hits = scan_data_files(args.root, data_path)
    all_findings += findings
    lines += ['- 扫描范围：`<项目根>`（本机绝对路径已脱敏）',
              '- 扫描文本文件：%d 个' % len(files),
              '- 对照数据源：`真实免勤表（本机，仓外）`（文件名不写入报告，避免真实表名外泄）',
              '- 对照库：真实学号 %d 个、真实姓名 %d 个' % (len(real_sids), len(real_names)),
              '- 合成测试学号（99 开头，虚构）：%d 个不同号段、%d 处引用'
              % (len(synthetic_sids), sum(synthetic_sids.values())),
              '- 被否口径旋钮扫描范围：`.py` 源码与界面文案 ＋ **xlsx 单元格内容**（不扫 `.md`）',
              '- 第 4 类闸门：词典「免勤理由分类」**逐格白名单**比对'
              '（关键词/分类/说明三列分别白名单化，白名单外任何行 ⇒ 高危阻断）', '']
    lines += ['## 一、代码与文本中的命中', '']
    lines += _table(findings, mask_names=not args.show_names) if findings else ['未发现绝对路径、真实学号或真实姓名。']
    lines += ['', '## 二、项目目录内的表格文件', '']
    if data_hits:
        lines += ['| 文件 | 判定 | 大小 |', '|---|---|---:|']
        for rel, kind, size in data_hits:
            lines.append('| %s | %s | %d |' % (rel, kind, size))
    else:
        lines.append('项目目录内没有表格文件。')

    # ---- 产物 ----
    if args.artifact:
        art_findings, art_files = scan_artifact(args.artifact, data_path)
        all_findings += art_findings
        # 报告随 EXE 分发，目录名只写末级名，绝不把本机绝对路径写进报告正文
        art_label = os.path.basename(os.path.normpath(args.artifact))
        lines += ['', '## 三、产物目录扫描（`%s`，含二进制）' % art_label, '']
        lines.append('- 扫描文件：%d 个（含 EXE 等二进制，UTF-8 与 UTF-16LE 双通道）' % len(art_files))
        lines.append('- 表格文件已逐单元格扫描')
        lines += ['', '### 命中', '']
        lines += _table(art_findings, mask_names=not args.show_names) if art_findings else ['未发现真实学号、真实姓名、绝对路径或被否口径旋钮。']
        lines += ['', '### 产物文件清单', '', '| 文件 | 说明 | 字节 |', '|---|---|---:|']
        for rel, kind, size in art_files:
            lines.append('| %s | %s | %d |' % (rel, kind, size))

    # ---- 推送侧 ----
    if args.push_repo:
        push_findings, tree, _ = scan_push(args.push_repo, args.push_branch, data_path)
        all_findings += push_findings
        data_files = [t for t in tree if os.path.splitext(t)[1].lower()
                      in ('.xlsx', '.xls', '.csv', '.zip')]
        lines += ['', '## 四、推送侧扫描（分支 %s）' % args.push_branch, '']
        lines.append('- 分支文件数：%d' % len(tree))
        lines.append('- 扫描面：**提交内容**（逐文件，含 xlsx 单元格）＋ **完整 patch 历史**')
        lines.append('- 分支内的数据/压缩文件：%s' % (', '.join(data_files) if data_files else '无'))
        lines += ['', '### 命中', '']
        lines += _table(push_findings, mask_names=not args.show_names) if push_findings else [
            '提交内容与提交历史均未发现真实学号、真实姓名、绝对路径或被否口径旋钮。']

    # ---- 结论 ----
    hard = [f for f in all_findings if f['kind'] in HARD_KINDS]
    soft = [f for f in all_findings if f['kind'] not in HARD_KINDS]
    lines += ['', '## 结论', '']
    if hard:
        lines.append('**不通过：存在高危命中，必须清理后再打包发布与推送。**')
    elif soft:
        lines.append('**有条件通过：无真实个人数据，仅存低危命中（需人工确认）。**')
    else:
        lines.append('**通过：未发现真实个人数据、绝对路径与被否口径旋钮。**')
    if all_findings:
        lines += ['', '命中总数：%d（高危 %d / 低危 %d）'
                  % (len(all_findings), len(hard), len(soft))]
    if soft:
        lines += ['', '### 低危命中（不阻断，但需人工确认）', '']
        lines += _table(soft, mask_names=not args.show_names)
    lines += ['', '> **分级规则**：真实姓名与**真实学号共现** ⇒ 高危（真实数据行形态，直接阻断）；'
              '仅姓名子串命中 ⇒ 低危（常见词误报：姓名库是子串匹配，'
              '如「XX基线」中的「XX」恰是某位学生的姓名），不阻断但需人工确认。']

    report = '\n'.join(lines)
    print(report)
    if args.out:
        out_path = args.out if os.path.isabs(args.out) else os.path.join(args.root, args.out)
        with open(out_path, 'w', encoding='utf-8') as fh:
            fh.write(report + '\n')
        print('\n报告已写出：%s' % out_path)
    return 1 if hard else 0


if __name__ == '__main__':
    raise SystemExit(main())
