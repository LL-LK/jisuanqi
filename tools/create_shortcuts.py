# -*- coding: utf-8 -*-
"""
三早一晚免勤自动统计系统 - 快捷方式创建工具
Author: worker_m3
Milestone: M3 (Shortcut Provisioning & Windows Installer Packaging)

为已部署在指定目录的“三早一晚免勤自动统计”软件创建当前用户桌面与开始菜单快捷方式（.lnk）。
默认目标路径：<部署目录>/三早一晚免勤自动统计.exe
"""

from __future__ import annotations

import argparse
import os
import subprocess
import sys
from typing import Dict, List, Optional


DEFAULT_TARGET_DIR = os.path.join("D:" + os.sep, "Software", "三早一晚免勤自动统计")
DEFAULT_EXE_NAME = "三早一晚免勤自动统计.exe"
DEFAULT_APP_NAME = "三早一晚免勤自动统计"
DEFAULT_DESCRIPTION = "三早一晚免勤自动统计系统 - 本地考勤去重与计算工具"


def _create_shortcuts_win32com(
    target_exe: str,
    target_dir: str,
    destinations: List[str],
    icon_location: str,
    description: str,
) -> List[str]:
    """使用 pywin32 的 win32com.client 创建快捷方式。"""
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    created: List[str] = []

    for dest in destinations:
        dest_dir = os.path.dirname(dest)
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir, exist_ok=True)

        shortcut = shell.CreateShortCut(dest)
        shortcut.TargetPath = target_exe
        shortcut.WorkingDirectory = target_dir
        shortcut.IconLocation = icon_location
        shortcut.Description = description
        shortcut.WindowStyle = 1
        shortcut.Save()
        created.append(dest)

    return created


def _create_shortcuts_powershell(
    target_exe: str,
    target_dir: str,
    destinations: List[str],
    icon_location: str,
    description: str,
) -> List[str]:
    """若 win32com 不可用，降级使用 PowerShell COM 自动化创建快捷方式。"""
    created: List[str] = []
    for dest in destinations:
        dest_dir = os.path.dirname(dest)
        if not os.path.exists(dest_dir):
            os.makedirs(dest_dir, exist_ok=True)

        ps_script = f"""
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("{dest}")
$Shortcut.TargetPath = "{target_exe}"
$Shortcut.WorkingDirectory = "{target_dir}"
$Shortcut.IconLocation = "{icon_location}"
$Shortcut.Description = "{description}"
$Shortcut.WindowStyle = 1
$Shortcut.Save()
"""
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if res.returncode != 0:
            raise RuntimeError(f"PowerShell 快捷方式创建失败: {res.stderr}")
        created.append(dest)

    return created


def create_shortcuts(
    target_dir: str = DEFAULT_TARGET_DIR,
    exe_name: str = DEFAULT_EXE_NAME,
    app_name: str = DEFAULT_APP_NAME,
    description: str = DEFAULT_DESCRIPTION,
    desktop_only: bool = False,
    start_menu_only: bool = False,
) -> Dict[str, str]:
    """创建桌面与开始菜单快捷方式。

    Args:
        target_dir: 可执行文件所在目录
        exe_name: 可执行文件名称
        app_name: 快捷方式主名称（不含 .lnk）
        description: 快捷方式描述信息
        desktop_only: 仅创建桌面快捷方式
        start_menu_only: 仅创建开始菜单快捷方式

    Returns:
        Dict[str, str]: 包含 'desktop' 和/或 'start_menu' 快捷方式实际文件路径的字典
    """
    target_exe = os.path.join(target_dir, exe_name)
    if not os.path.isfile(target_exe):
        raise FileNotFoundError(f"目标可执行文件未找到: {target_exe}")

    # 获取当前用户的标准特殊目录
    # 优先使用环境变量与 APPDATA
    desktop_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    appdata = os.environ.get("APPDATA")
    if not appdata:
        appdata = os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
    programs_dir = os.path.join(appdata, "Microsoft", "Windows", "Start Menu", "Programs")

    icon_location = f"{target_exe},0"
    results: Dict[str, str] = {}
    destinations: List[str] = []

    if not start_menu_only:
        desktop_lnk = os.path.join(desktop_dir, f"{app_name}.lnk")
        destinations.append(desktop_lnk)
        results["desktop"] = desktop_lnk

    if not desktop_only:
        start_menu_lnk = os.path.join(programs_dir, f"{app_name}.lnk")
        destinations.append(start_menu_lnk)
        results["start_menu"] = start_menu_lnk

    try:
        _create_shortcuts_win32com(
            target_exe=target_exe,
            target_dir=target_dir,
            destinations=destinations,
            icon_location=icon_location,
            description=description,
        )
    except Exception as e:
        print(f"[WARN] win32com 创建快捷方式异常 ({e})，尝试降级至 PowerShell COM...")
        _create_shortcuts_powershell(
            target_exe=target_exe,
            target_dir=target_dir,
            destinations=destinations,
            icon_location=icon_location,
            description=description,
        )

    for key, path in results.items():
        if os.path.isfile(path):
            print(f"[SUCCESS] {key.capitalize()} shortcut created: {path}")
        else:
            raise RuntimeError(f"快捷方式创建后未在磁盘检测到: {path}")

    return results


def main() -> int:
    parser = argparse.ArgumentParser(description="创建三早一晚免勤自动统计快捷方式")
    parser.add_argument("--target-dir", default=DEFAULT_TARGET_DIR, help="程序部署根目录")
    parser.add_argument("--exe-name", default=DEFAULT_EXE_NAME, help="主执行文件名称")
    parser.add_argument("--app-name", default=DEFAULT_APP_NAME, help="快捷方式名称")
    parser.add_argument("--description", default=DEFAULT_DESCRIPTION, help="快捷方式描述")
    parser.add_argument("--desktop-only", action="store_true", help="仅创建桌面快捷方式")
    parser.add_argument("--start-menu-only", action="store_true", help="仅创建开始菜单快捷方式")

    args = parser.parse_args()

    try:
        created = create_shortcuts(
            target_dir=args.target_dir,
            exe_name=args.exe_name,
            app_name=args.app_name,
            description=args.description,
            desktop_only=args.desktop_only,
            start_menu_only=args.start_menu_only,
        )
        print(f"\n全部快捷方式创建完成！共生成 {len(created)} 个快捷方式。")
        return 0
    except Exception as err:
        print(f"[ERROR] 快捷方式创建失败: {err}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
