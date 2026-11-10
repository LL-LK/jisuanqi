# -*- coding: utf-8 -*-
"""临时文件、缓存目录与运行垃圾统一管理中心。

确保所有运行时中间文件、字体自检图、原子写入缓冲文件与缓存碎片
全量收敛至系统 %TEMP%/AttendanceStats_temp 目录下，
并在退出时通过 atexit 自动清扫，实现对程序工作区零垃圾残留。
"""

from __future__ import annotations

import os
import sys
import shutil
import atexit
import tempfile
from typing import List, Optional


TEMP_SUBDIR = "AttendanceStats_temp"
CACHE_SUBDIR = "AttendanceStats_cache"


def get_base_temp_dir() -> str:
    """获取应用专属系统临时根目录路径。"""
    sys_temp = tempfile.gettempdir()
    target = os.path.join(sys_temp, TEMP_SUBDIR)
    os.makedirs(target, exist_ok=True)
    return os.path.abspath(target)


def get_temp_dir(sub: Optional[str] = None) -> str:
    """获取或创建临时子目录。"""
    base = get_base_temp_dir()
    if sub:
        target = os.path.join(base, sub)
        os.makedirs(target, exist_ok=True)
        return os.path.abspath(target)
    return base


def get_cache_dir() -> str:
    """获取缓存目录路径。"""
    return get_temp_dir("cache")


def get_atomic_temp_dir() -> str:
    """获取原子写入专用的安全临时缓冲目录。"""
    return get_temp_dir("atomic_writes")


def get_font_selftest_dir() -> str:
    """获取字体自检图片的输出目录。"""
    return get_temp_dir("font_tests")


def cleanup_temp_files(dry_run: bool = False) -> int:
    """清理本应用专属临时目录中的所有文件与子目录。
    
    返回被成功删除的条目数。
    """
    base = os.path.join(tempfile.gettempdir(), TEMP_SUBDIR)
    if not os.path.exists(base):
        return 0

    removed_count = 0
    for root, dirs, files in os.walk(base, topdown=False):
        for f in files:
            fp = os.path.join(root, f)
            if not dry_run:
                try:
                    os.chmod(fp, 0o777)
                    os.remove(fp)
                    removed_count += 1
                except Exception:
                    pass
        for d in dirs:
            dp = os.path.join(root, d)
            if not dry_run:
                try:
                    os.rmdir(dp)
                    removed_count += 1
                except Exception:
                    pass
    return removed_count


def cleanup_legacy_workspace_garbage(workspace_dir: str = ".") -> List[str]:
    """排查并安全清理软件工作目录下的历史遗留垃圾。
    
    例如: _work/ 目录、font_selftest.* 文件。
    返回已清理的路径列表。
    """
    cleaned: List[str] = []
    abs_ws = os.path.abspath(workspace_dir)

    # 1. 清理 legacy _work 目录 (若仅为中间产物)
    legacy_work = os.path.join(abs_ws, "_work")
    if os.path.exists(legacy_work) and os.path.isdir(legacy_work):
        try:
            shutil.rmtree(legacy_work, ignore_errors=True)
            cleaned.append(legacy_work)
        except Exception:
            pass

    # 2. 清理 font_selftest 相关文件
    for name in os.listdir(abs_ws):
        if name.startswith("font_selftest."):
            target = os.path.join(abs_ws, name)
            try:
                os.remove(target)
                cleaned.append(target)
            except Exception:
                pass

    return cleaned


# 注册退出自动清理钩子
atexit.register(cleanup_temp_files)
