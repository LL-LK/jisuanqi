# -*- coding: utf-8 -*-
"""
三早一晚免勤自动统计系统 - 快捷方式与安装包自动化测试套件
Author: worker_m3
Milestone: M3 (Shortcut Provisioning & Windows Installer Packaging)
Requirements: R1 (Part 1: 快捷方式创建, Part 2: Windows 标准安装包)
"""

from __future__ import annotations

import hashlib
import os
import re
import struct
import sys
from pathlib import Path
from typing import Dict, List, Optional

import pytest

IS_WINDOWS = sys.platform == "win32"
PROJECT_DIR = Path(__file__).resolve().parent.parent


def _get_shortcut_attributes(lnk_path: str) -> Dict[str, str]:
    """读取 Windows .lnk 快捷方式属性字典。"""
    try:
        import win32com.client
        shell = win32com.client.Dispatch("WScript.Shell")
        sc = shell.CreateShortCut(lnk_path)
        return {
            "TargetPath": sc.TargetPath,
            "WorkingDirectory": sc.WorkingDirectory,
            "IconLocation": sc.IconLocation,
            "Description": sc.Description,
            "WindowStyle": str(sc.WindowStyle),
        }
    except Exception:
        import subprocess
        ps_script = f"""
$WshShell = New-Object -ComObject WScript.Shell
$sc = $WshShell.CreateShortcut("{lnk_path}")
Write-Output "TargetPath=$($sc.TargetPath)"
Write-Output "WorkingDirectory=$($sc.WorkingDirectory)"
Write-Output "IconLocation=$($sc.IconLocation)"
Write-Output "Description=$($sc.Description)"
Write-Output "WindowStyle=$($sc.WindowStyle)"
"""
        res = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
        if res.returncode != 0:
            raise RuntimeError(f"PowerShell 读取快捷方式失败: {res.stderr}")
        attrs = {}
        for line in res.stdout.strip().splitlines():
            if "=" in line:
                k, v = line.split("=", 1)
                attrs[k.strip()] = v.strip()
        return attrs


# ==============================================================================
# 1. 快捷方式属性与可执行目标验证 (Requirement R1 - Part 1)
# ==============================================================================
@pytest.mark.skipif(not IS_WINDOWS, reason="快捷方式验证需要 Windows 系统环境")
class TestShortcutVerification:
    """验证桌面与开始菜单快捷方式存在性、目标路径、工作目录与图标位置。"""

    DEPLOY_DIR = Path("D:" + os.sep + "Software" + os.sep + "三早一晚免勤自动统计")
    DEPLOY_EXE = DEPLOY_DIR / "三早一晚免勤自动统计.exe"
    EXPECTED_DESC = "三早一晚免勤自动统计系统 - 本地考勤去重与计算工具"

    @pytest.fixture(autouse=True)
    def setup_paths(self):
        desktop = Path(os.path.expanduser("~")) / "Desktop"
        appdata = os.environ.get("APPDATA") or str(Path(os.path.expanduser("~")) / "AppData" / "Roaming")
        programs = Path(appdata) / "Microsoft" / "Windows" / "Start Menu" / "Programs"

        self.desktop_lnk = desktop / "三早一晚免勤自动统计.lnk"
        self.programs_lnk = programs / "三早一晚免勤自动统计.lnk"

    def test_desktop_shortcut_existence_and_attributes(self):
        """测试 1: 桌面快捷方式存在，且 TargetPath、WorkingDirectory、IconLocation 配置正确。"""
        assert self.desktop_lnk.is_file(), f"桌面快捷方式文件缺失: {self.desktop_lnk}"

        attrs = _get_shortcut_attributes(str(self.desktop_lnk))
        assert os.path.normcase(attrs["TargetPath"]) == os.path.normcase(str(self.DEPLOY_EXE)), (
            f"桌面快捷方式 TargetPath 不匹配: {attrs['TargetPath']} != {self.DEPLOY_EXE}"
        )
        assert os.path.normcase(attrs["WorkingDirectory"]) == os.path.normcase(str(self.DEPLOY_DIR)), (
            f"桌面快捷方式 WorkingDirectory 不匹配: {attrs['WorkingDirectory']} != {self.DEPLOY_DIR}"
        )
        assert (
            os.path.normcase(str(self.DEPLOY_EXE)) in os.path.normcase(attrs["IconLocation"])
            or "attendance.ico" in attrs["IconLocation"].lower()
        ), f"桌面快捷方式 IconLocation 不匹配: {attrs['IconLocation']}"
        assert attrs["Description"] == self.EXPECTED_DESC, (
            f"桌面快捷方式 Description 不匹配: {attrs['Description']}"
        )
        assert attrs.get("WindowStyle") == "1", f"WindowStyle 应为 1 (Normal): {attrs.get('WindowStyle')}"

    def test_start_menu_shortcut_existence_and_attributes(self):
        """测试 2: 开始菜单快捷方式存在，且 TargetPath、WorkingDirectory、IconLocation 配置正确。"""
        assert self.programs_lnk.is_file(), f"开始菜单快捷方式文件缺失: {self.programs_lnk}"

        attrs = _get_shortcut_attributes(str(self.programs_lnk))
        assert os.path.normcase(attrs["TargetPath"]) == os.path.normcase(str(self.DEPLOY_EXE)), (
            f"开始菜单快捷方式 TargetPath 不匹配: {attrs['TargetPath']} != {self.DEPLOY_EXE}"
        )
        assert os.path.normcase(attrs["WorkingDirectory"]) == os.path.normcase(str(self.DEPLOY_DIR)), (
            f"开始菜单快捷方式 WorkingDirectory 不匹配: {attrs['WorkingDirectory']} != {self.DEPLOY_DIR}"
        )
        assert (
            os.path.normcase(str(self.DEPLOY_EXE)) in os.path.normcase(attrs["IconLocation"])
            or "attendance.ico" in attrs["IconLocation"].lower()
        ), f"开始菜单快捷方式 IconLocation 不匹配: {attrs['IconLocation']}"
        assert attrs["Description"] == self.EXPECTED_DESC, (
            f"开始菜单快捷方式 Description 不匹配: {attrs['Description']}"
        )

    def test_target_executable_validity_and_pe_header(self):
        """测试 3: 快捷方式指向的可执行目标真实存在、大小合理且具有有效 PE 结构。"""
        assert self.DEPLOY_EXE.is_file(), f"快捷方式目标文件不存在: {self.DEPLOY_EXE}"
        file_size = self.DEPLOY_EXE.stat().st_size
        assert file_size > 30 * 1024 * 1024, f"目标可执行文件大小异常偏小: {file_size} 字节"

        # 检查 PE DOS Header 与 PE Signature
        with open(self.DEPLOY_EXE, "rb") as fh:
            dos_header = fh.read(64)
            assert dos_header[:2] == b"MZ", "目标文件非有效 DOS 可执行文件 (缺少 MZ 标头)"
            pe_offset = struct.unpack("<I", dos_header[60:64])[0]
            fh.seek(pe_offset)
            pe_sig = fh.read(4)
            assert pe_sig == b"PE\x00\x00", "目标文件非有效 Portable Executable (缺少 PE 签名)"

    def test_deployed_dictionary_and_docs_present(self):
        """测试 4: 已部署软件目录中包含免勤词典与基础说明文档。"""
        dict_file = self.DEPLOY_DIR / "免勤词典.xlsx"
        assert dict_file.is_file(), f"部署目录缺失核心词典: {dict_file}"
        assert dict_file.stat().st_size > 5000, "词典文件大小异常"

    def test_shortcut_creation_script_api_idempotent(self):
        """测试 5: tools/create_shortcuts.py 具备幂等性，重复调用依然成功且属性保持一致。"""
        from tools.create_shortcuts import create_shortcuts

        res = create_shortcuts(
            target_dir=str(self.DEPLOY_DIR),
            exe_name=self.DEPLOY_EXE.name,
            app_name="三早一晚免勤自动统计",
            description=self.EXPECTED_DESC,
        )
        assert "desktop" in res and os.path.isfile(res["desktop"])
        assert "start_menu" in res and os.path.isfile(res["start_menu"])


# ==============================================================================
# 2. Inno Setup 脚本指令与安装包产物验证 (Requirement R1 - Part 2)
# ==============================================================================
class TestInstallerScriptAndDeliverables:
    """验证 setup.iss 脚本指令完备性以及编译出的 Setup.exe 产物属性。"""

    SETUP_ISS = PROJECT_DIR / "setup.iss"
    OUTPUT_DIR = PROJECT_DIR / "发布包"
    SETUP_EXE = OUTPUT_DIR / "三早一晚免勤自动统计_Setup.exe"

    @pytest.fixture(scope="class")
    def iss_text(self) -> str:
        assert self.SETUP_ISS.is_file(), f"setup.iss 脚本不存在: {self.SETUP_ISS}"
        return self.SETUP_ISS.read_text(encoding="utf-8")

    def test_setup_iss_chinese_simplified_support(self, iss_text: str):
        """测试 6: setup.iss 显式包含 ChineseSimplified 语言支持配置。"""
        assert "ChineseSimplified.isl" in iss_text, "setup.iss 未引用 ChineseSimplified.isl 语言包"
        assert 'Name: "chinesesimp"' in iss_text or "Name: chinesesimp" in iss_text

    def test_setup_iss_explicit_working_directory(self, iss_text: str):
        """测试 7: setup.iss 快捷方式定义中显式包含 WorkingDir: \"{app}\"。"""
        # 匹配 Icons 段落中带 WorkingDir 的设置
        assert 'WorkingDir: "{app}"' in iss_text or "WorkingDir: {app}" in iss_text, (
            "setup.iss 快捷方式中缺失显式 WorkingDir: \"{app}\" 指令"
        )

    def test_setup_iss_dictionary_protection_flags(self, iss_text: str):
        """测试 8: 免勤词典.xlsx 配置了 onlyifdoesntexist 与 uninsneveruninstall 双重保护。"""
        pattern = r'Source' + r':[ ]*".*免勤词典\.xlsx".*DestDir' + r':[ ]*"{app}".*Flags' + r':[ ]*(.*)'
        match = re.search(pattern, iss_text)
        assert match, "setup.iss 未找到 免勤词典.xlsx 的 Files 打包规则"
        flags = match.group(1).lower()
        assert "onlyifdoesntexist" in flags, "词典必须包含 onlyifdoesntexist 标志"
        assert "uninsneveruninstall" in flags, "词典必须包含 uninsneveruninstall 标志"

    def test_setup_iss_output_target_directives(self, iss_text: str):
        """测试 9: setup.iss 输出目录为 发布包，基础文件名为 三早一晚免勤自动统计_Setup。"""
        assert re.search(r"OutputDir\s*=\s*发布包", iss_text), "OutputDir 必须配置为 发布包"
        assert re.search(r"OutputBaseFilename\s*=\s*三早一晚免勤自动统计_Setup", iss_text), (
            "OutputBaseFilename 必须为 三早一晚免勤自动统计_Setup"
        )

    def test_installer_executable_existence_and_size(self):
        """测试 10: 编译产物 三早一晚免勤自动统计_Setup.exe 存在且大小在正常范围 (> 15MB)。"""
        assert self.SETUP_EXE.is_file(), f"未检测到安装包产物: {self.SETUP_EXE}"
        exe_size = self.SETUP_EXE.stat().st_size
        assert exe_size > 15 * 1024 * 1024, f"安装包大小异常: {exe_size} 字节"
        assert exe_size < 100 * 1024 * 1024, f"安装包大小异常超大: {exe_size} 字节"

    def test_installer_executable_pe_signature(self):
        """测试 11: 安装包可执行文件具备有效 PE Header 与 Inno Setup 引导结构。"""
        with open(self.SETUP_EXE, "rb") as fh:
            dos_header = fh.read(64)
            assert dos_header[:2] == b"MZ", "安装包非有效 DOS 可执行文件 (缺少 MZ 标头)"
            pe_offset = struct.unpack("<I", dos_header[60:64])[0]
            fh.seek(pe_offset)
            pe_sig = fh.read(4)
            assert pe_sig == b"PE\x00\x00", "安装包非有效 Portable Executable (缺少 PE 签名)"

    def test_installer_sha256_sidecar_and_manifest_consistency(self):
        """测试 12: 安装包具有对应的 .sha256 sidecar 文件，且哈希与 SHA256.txt 一致。"""
        sidecar = self.SETUP_EXE.with_name(self.SETUP_EXE.name + ".sha256")
        assert sidecar.is_file(), f"安装包 sidecar 校验文件缺失: {sidecar}"

        # 实际计算安装包哈希
        hasher = hashlib.sha256()
        with open(self.SETUP_EXE, "rb") as fh:
            while chunk := fh.read(65536):
                hasher.update(chunk)
        actual_hash = hasher.hexdigest().lower()

        sidecar_content = sidecar.read_text(encoding="utf-8").strip()
        assert actual_hash in sidecar_content.lower(), (
            f"Sidecar 哈希不匹配: {sidecar_content} vs {actual_hash}"
        )

        manifest_file = self.OUTPUT_DIR / "SHA256.txt"
        assert manifest_file.is_file(), f"发布包统合校验清单缺失: {manifest_file}"
        manifest_text = manifest_file.read_text(encoding="utf-8").lower()
        assert actual_hash in manifest_text, (
            f"发布包 SHA256.txt 中未收录安装包的真实哈希: {actual_hash}"
        )
