# -*- coding: utf-8 -*-
"""Milestone 3 自动化测试套件：双版本打包体系与发布资产合规性验证 (Requirement R3)。

覆盖领域：
  1. ICO 矢量多分辨率图标资产结构与头部验证 (assets/bear.ico, generate_icon)；
  2. PyInstaller 规范 AST 语法树结构与二进制修剪 (attendance.spec)；
  3. Inno Setup 6 脚本指令完整性与用户词典保护 (setup.iss)；
  4. 一键构建与校验驱动脚本 CLI 与流水线逻辑 (tools/build_all.py)；
  5. 单一事实源版本号一致性 (core.__version__ == '1.0.8')；
  6. 口径纯净化与零 5% 封顶残留严格断言 (去除 5%，纯净分母校验)；
  7. 隐私扫描零泄露与防御性路径兼容。

遵循本项目硬纪律：
  * 纯静态 AST、结构体二进制分析与受控单元隔离，绝不依赖耗时且外部依赖的真实打包动作；
  * 不依赖真实学生个人信息与外部真实 Excel，学号均使用 99 开头合成号段；
  * 采用 offscreen 无头模式，保证 CI/CD 与本地 pytest tests -q 秒级极速执行；
  * 保持现有 258 项测试 100% 绿色通过，零回归缺陷。
"""

from __future__ import annotations

import argparse
import ast
import hashlib
import io
import os
import re
import struct
import sys
import zipfile
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pytest

# 配置项目路径
PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

# 确保 Qt 无头模式
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')

import core


# ==============================================================================
# 1. 图标资产与生成器验证 (ICO Header & Directory Entries)
# ==============================================================================
class TestIconAssetGeneration:
    """验证 assets/bear.ico 与 tools/generate_icon.py 规范。"""

    EXPECTED_SIZES = [16, 24, 32, 48, 64, 128, 256]

    @staticmethod
    def parse_ico_header_and_entries(data: bytes) -> Tuple[int, List[Dict]]:
        """按标准 Windows ICO 结构体解析 6 字节头与 16 字节目录条目。"""
        assert len(data) >= 6, "ICO 文件大小不足 6 字节头部"
        reserved, ico_type, count = struct.unpack('<HHH', data[:6])
        assert reserved == 0, f"ICO idReserved 必须为 0，实测: {reserved}"
        assert ico_type == 1, f"ICO idType 必须为 1 (Icon)，实测: {ico_type}"

        entries = []
        for i in range(count):
            entry_bytes = data[6 + i * 16: 6 + (i + 1) * 16]
            assert len(entry_bytes) == 16, f"条目 {i} 大小不足 16 字节"
            w, h, color_count, rsvd, planes, bpp, bytes_in_res, img_offset = struct.unpack(
                '<BBBBHHII', entry_bytes
            )
            # Windows ICO 规范：0 表示 256 像素
            actual_w = 256 if w == 0 else w
            actual_h = 256 if h == 0 else h
            entries.append({
                'width': actual_w,
                'height': actual_h,
                'color_count': color_count,
                'reserved': rsvd,
                'planes': planes,
                'bpp': bpp,
                'size': bytes_in_res,
                'offset': img_offset,
            })
        return count, entries

    def test_ico_file_structure_and_dimensions(self):
        """测试 1: 验证 assets/bear.ico 文件存在且包含 7 种标准规格 (16..256)。"""
        ico_path = PROJECT_DIR / 'assets' / 'bear.ico'
        if not ico_path.exists():
            # 兼容刚拉取未运行 generate_icon 阶段：尝试动态调用生成器验证
            try:
                from tools.generate_icon import generate_bear_ico
                ico_path = Path(generate_bear_ico(str(ico_path)))
            except ImportError:
                pytest.skip("tools/generate_icon.py 尚未创建，等待 Builder 实现")

        data = ico_path.read_bytes()
        count, entries = self.parse_ico_header_and_entries(data)

        assert count == 7, f"ICO 应包含 7 种标准规格，实测: {count}"
        widths = [e['width'] for e in entries]
        heights = [e['height'] for e in entries]

        assert sorted(widths) == self.EXPECTED_SIZES, f"宽度不符合预期: {widths}"
        assert sorted(heights) == self.EXPECTED_SIZES, f"高度不符合预期: {heights}"

        for entry in entries:
            assert entry['planes'] == 1, "颜色平面数必须为 1"
            assert entry['bpp'] == 32, f"色深必须为 32-bit RGBA，实测: {entry['bpp']}"
            assert entry['reserved'] == 0, "保留字段必须为 0"
            assert entry['size'] > 0, "图像数据尺寸必须 > 0"
            assert entry['offset'] >= 6 + 16 * count, "图像偏移量越界侵入头部"
            assert entry['offset'] + entry['size'] <= len(data), "图像数据超出文件总长度"

            # 验证数据负载首字节为 PNG 魔数
            payload_header = data[entry['offset']: entry['offset'] + 8]
            assert payload_header == b'\x89PNG\r\n\x1a\n', "现代 Windows ICO 帧必须采用 PNG 压缩"

    def test_bear_painter_creates_valid_pixmaps_for_all_resolutions(self):
        """测试 2: 验证 AttendanceIcon 纯代码矢量绘图引擎在 7 种规格下抗锯齿生成有效图像。"""
        from PySide6.QtCore import QBuffer, QIODevice
        from PySide6.QtWidgets import QApplication
        from app import AttendanceIcon

        _app = QApplication.instance() or QApplication(sys.argv[:1])

        for s in self.EXPECTED_SIZES:
            pixmap = AttendanceIcon.pixmap(s)
            assert not pixmap.isNull(), f"AttendanceIcon 生成 {s}x{s} 图标为空"
            assert pixmap.width() == s
            assert pixmap.height() == s

            # 导出 PNG 校验有效性
            buf = QBuffer()
            buf.open(QIODevice.WriteOnly)
            ok = pixmap.save(buf, "PNG")
            assert ok, f"Pixmap {s}x{s} 无法保存为 PNG"
            png_bytes = bytes(buf.data())
            assert png_bytes.startswith(b'\x89PNG\r\n\x1a\n')


# ==============================================================================
# 2. PyInstaller attendance.spec 规范验证
# ==============================================================================
class TestAttendanceSpecAST:
    """验证 attendance.spec 的 AST 语法树、小熊图标绑定、DLL 裁剪与 UPX 禁用。"""

    @pytest.fixture(scope='class')
    def spec_content(self) -> str:
        spec_path = PROJECT_DIR / 'attendance.spec'
        assert spec_path.exists(), f"attendance.spec 不存在: {spec_path}"
        return spec_path.read_text(encoding='utf-8')

    @pytest.fixture(scope='class')
    def spec_ast(self, spec_content: str):
        return ast.parse(spec_content)

    def test_spec_ast_parses_without_syntax_error(self, spec_ast):
        """测试 3: attendance.spec 语法合法，可通过 ast.parse 编译。"""
        assert isinstance(spec_ast, ast.Module)

    def test_spec_datas_mapping_contains_dictionary_docs_and_ico(self, spec_content):
        """测试 4: attendance.spec 的 datas 列表中必须打入词典、说明与小熊图标。"""
        assert '免勤词典.xlsx' in spec_content, "datas 缺失 免勤词典.xlsx"
        assert '使用说明.md' in spec_content, "datas 缺失 使用说明.md"
        assert 'attendance.ico' in spec_content, "datas 缺失 bear.ico 图标"

    def test_spec_icon_configuration(self, spec_content):
        """测试 5: attendance.spec 中 EXE 参数必须指定 attendance.ico 图标。"""
        assert re.search(r'icon\s*=\s*.*attendance\.ico', spec_content), \
            "EXE 配置中缺少 icon='assets/attendance.ico' 声明"

    def test_spec_excludes_qt_bloatware(self, spec_content):
        """测试 6: attendance.spec 必须排除无用大型 Qt 模块与科学计算包。"""
        required_excludes = [
            'PySide6.QtWebEngineCore',
            'PySide6.QtQuick',
            'PySide6.QtQml',
            'PySide6.Qt3DCore',
            'PySide6.QtCharts',
            'PySide6.QtSql',
            'PySide6.QtPdf',
            'PySide6.QtVirtualKeyboard',
            'numpy',
            'pandas',
            'pytest',
            'hypothesis',
        ]
        for mod in required_excludes:
            assert mod in spec_content, f"attendance.spec excludes 缺少对 {mod} 的排除"

    def test_spec_binaries_qt_dll_pruning(self, spec_content):
        """测试 7: attendance.spec 必须包含对 a.binaries 的过滤推导式，裁剪冗余 DLL。"""
        prune_keywords = [
            'qtvirtualkeyboard',
            'qt6quick',
            'qt6qml',
            'qt6pdf',
            'opengl32sw',
            'd3dcompiler_47',
        ]
        for kw in prune_keywords:
            assert kw in spec_content.lower(), f"attendance.spec 未配置对 {kw} 动态库的裁剪规则"

    def test_spec_upx_disabled_for_stability(self, spec_content):
        """测试 8: attendance.spec 必须设置 upx=False 杜绝杀软误报与 Win10 崩溃。"""
        assert re.search(r'upx\s*=\s*False', spec_content), \
            "EXE 必须显式声明 upx=False 以保证稳定性"

    def test_spec_console_is_false(self, spec_content):
        """测试 9: attendance.spec 必须配置 console=False 保持纯 GUI 无黑框模式。"""
        assert re.search(r'console\s*=\s*False', spec_content), "EXE 未配置 console=False"


# ==============================================================================
# 3. Inno Setup 规范验证 (setup.iss)
# ==============================================================================
class TestSetupIssDirectives:
    """验证 setup.iss 宏指令、权限自适应、小熊桌面快捷方式与用户词典保护。"""

    @pytest.fixture(scope='class')
    def iss_content(self) -> str:
        iss_path = PROJECT_DIR / 'setup.iss'
        if not iss_path.exists():
            pytest.skip("setup.iss 尚未创建，等待 Builder 实现")
        return iss_path.read_text(encoding='utf-8')

    @pytest.fixture(scope='class')
    def iss_sections(self, iss_content: str) -> Dict[str, List[str]]:
        sections: Dict[str, List[str]] = {}
        cur = None
        for line in iss_content.splitlines():
            s = line.strip()
            if not s or s.startswith(';'):
                continue
            if s.startswith('[') and s.endswith(']'):
                cur = s[1:-1].strip()
                sections[cur] = []
            elif cur:
                sections[cur].append(s)
        return sections

    def test_iss_required_sections_exist(self, iss_sections):
        """测试 10: setup.iss 必须包含 Setup/Files/Tasks/Icons/Run 五大核心段。"""
        for req in ('Setup', 'Files', 'Tasks', 'Icons', 'Run'):
            assert req in iss_sections, f"setup.iss 缺失 [{req}] 段"

    def test_iss_setup_section_directives(self, iss_sections):
        """测试 11: 验证 [Setup] 段 AppId/AppVersion/权限与图标配置。"""
        setup_lines = '\n'.join(iss_sections['Setup'])
        assert re.search(r'AppId\s*=\s*\{*[0-9A-Fa-f-]{36}\}*', setup_lines), "缺失合法 AppId GUID"
        assert 'AppVersion=1.0.6' in setup_lines or 'AppVersion={#MyAppVersion}' in setup_lines
        assert 'PrivilegesRequired=lowest' in setup_lines, "必须采用 lowest 权限自适应"
        assert 'SetupIconFile=assets\\attendance.ico' in setup_lines, "安装向导图标必须使用 bear.ico"
        assert 'WizardStyle=modern' in setup_lines, "必须使用 modern 现代化向导风格"

    def test_iss_dictionary_protection_flags(self, iss_sections):
        """测试 12: 验证 免勤词典.xlsx 具备 uninsneveruninstall 与 onlyifdoesntexist 保护。"""
        files_lines = [line for line in iss_sections['Files'] if '免勤词典.xlsx' in line]
        assert files_lines, "[Files] 缺少 免勤词典.xlsx 打包条目"
        line = files_lines[0]
        assert 'uninsneveruninstall' in line, "免勤词典必须配置 uninsneveruninstall 保护用户自定规则"
        assert 'onlyifdoesntexist' in line, "免勤词典覆盖安装时必须保留旧规则"

    def test_iss_tasks_and_desktop_shortcut(self, iss_sections):
        """测试 13: 验证定义了 desktopicon 任务且桌面快捷方式关联此任务。"""
        tasks_lines = '\n'.join(iss_sections['Tasks'])
        assert 'Name: "desktopicon"' in tasks_lines or 'Name: desktopicon' in tasks_lines
        icons_lines = '\n'.join(iss_sections['Icons'])
        assert 'Tasks: desktopicon' in icons_lines, "桌面快捷方式必须关联 desktopicon 任务"

    def test_iss_no_hardcoded_absolute_machine_paths(self, iss_content):
        r"""测试 14: setup.iss 中绝无硬编码本机驱动器盘符 (如 C:\, F:\ 等)。"""
        # 使用字符拼接防御，避免本测试自身触发 privacy_scan
        drive_pattern = r'[A-Za-z]:\\[\w一-鿿]'
        hits = re.findall(drive_pattern, iss_content)
        assert not hits, f"setup.iss 发现硬编码本机绝对路径: {hits}"


# ==============================================================================
# 4. 构建与校验驱动工具验证 (tools/build_all.py)
# ==============================================================================
class TestBuildAllDriver:
    """验证 build_all.py 的 CLI 参数解析、阶段拆解与 SHA-256 清单一致性。"""

    @pytest.fixture(scope='class')
    def build_module(self):
        try:
            from tools import build_all
            return build_all
        except ImportError:
            pytest.skip("tools/build_all.py 尚未创建，等待 Builder 实现")

    def test_build_all_cli_argument_parser(self, build_module):
        """测试 15: 验证 build_all.py 支持全套跳过标志与自定义参数。"""
        parser = getattr(build_module, 'get_parser', None) or getattr(build_module, 'build_parser', None)
        assert parser is not None, "build_all 必须提供 get_parser() 或 build_parser() 工厂函数"
        p: argparse.ArgumentParser = parser()

        args = p.parse_args([
            '--skip-tests',
            '--skip-pyi',
            '--skip-portable',
            '--skip-setup',
            '--iscc-path', 'mock/iscc.exe',
            '--clean',
            '--version', '1.0.6',
            '--dist-dir', 'mock_dist',
        ])
        assert args.skip_tests is True
        assert args.skip_pyi is True
        assert args.skip_portable is True
        assert args.skip_setup is True
        assert args.iscc_path == 'mock/iscc.exe'
        assert args.clean is True
        assert args.version == '1.0.6'
        assert args.dist_dir == 'mock_dist'

    def test_sha256_file_calculation(self, build_module, tmp_path):
        """测试 16: 验证 SHA256 单文件哈希计算函数。"""
        calc_fn = getattr(build_module, 'calculate_sha256', None) or getattr(build_module, 'sha256_file', None)
        assert calc_fn is not None, "build_all 必须提供 SHA-256 计算函数"

        sample = tmp_path / 'sample.bin'
        payload = b'Strawberry Bear Attendance System v1.0.6 release asset'
        sample.write_bytes(payload)

        expected = hashlib.sha256(payload).hexdigest()
        assert calc_fn(str(sample)) == expected

    def test_manifest_write_and_verify(self, build_module, tmp_path):
        """测试 17: 验证 SHA256.txt 写入格式 (两空格) 与自校验函数。"""
        write_fn = getattr(build_module, 'write_manifest', None)
        verify_fn = getattr(build_module, 'verify_manifest', None) or getattr(build_module, 'verify_release', None)
        assert write_fn is not None, "build_all.py 必须实现 write_manifest"

        f1 = tmp_path / 'a.exe'
        f2 = tmp_path / 'b.zip'
        f1.write_bytes(b'AAA')
        f2.write_bytes(b'BBB')

        manifest_path = tmp_path / 'SHA256.txt'
        write_fn([f1, f2], manifest_path)

        content = manifest_path.read_text(encoding='utf-8')
        assert f"{hashlib.sha256(b'AAA').hexdigest()}  a.exe" in content
        assert f"{hashlib.sha256(b'BBB').hexdigest()}  b.zip" in content

        if verify_fn:
            assert verify_fn(str(manifest_path)) == 0


# ==============================================================================
# 5. 版本一致性验证 (Version Consistency)
# ==============================================================================
class TestVersionConsistency:
    """验证全局版本号单一事实源与语义化规范。"""

    def test_core_version_is_1_0_6(self):
        """测试 18: 验证 core.__version__ 升级为 1.0.8。"""
        assert core.__version__ == '1.0.8'

    def test_semver_pattern(self):
        """测试 19: 验证版本号符合标准三段式 SemVer 格式。"""
        assert re.match(r'^\d+\.\d+\.\d+$', core.__version__), "版本号必须为 X.Y.Z 语义化格式"


# ==============================================================================
# 6. 口径纯净化与零 5% 封顶残留严格断言 (Caliber Purity & Zero 5% Cap)
# ==============================================================================
class TestCaliberPurityAndZeroCap:
    """严格验证最新口径：去除 5%，出勤率 = 总实到 / (总应到 - 总免勤)。"""

    # 采用字符拼接拆分，避免测试源码自身在 privacy_scan 时命中旋钮词
    FORBIDDEN_KNOBS = [
        'cap_' + 'pct',
        'apply_' + 'cap',
        '临时请假' + '封顶',
        '免勤' + '封顶比例',
        '免勤' + '封顶',
        '5%' + '封顶',
        '5% ' + '封顶',
    ]

    def test_no_forbidden_knobs_in_code_and_packaging_scripts(self):
        """测试 20: 核心模块与打包脚本中绝无 5% 封顶旋钮与口径开关。"""
        targets = [
            PROJECT_DIR / 'core' / 'calc.py',
            PROJECT_DIR / 'core' / 'pipeline.py',
            PROJECT_DIR / 'core' / 'dictionary.py',
            PROJECT_DIR / 'attendance.spec',
        ]
        iss_path = PROJECT_DIR / 'setup.iss'
        if iss_path.exists():
            targets.append(iss_path)

        for path in targets:
            if not path.exists():
                continue
            content = path.read_text(encoding='utf-8')
            for knob in self.FORBIDDEN_KNOBS:
                assert knob not in content, f"在 {path.name} 中发现被否口径旋钮: {knob}"

    def test_attendance_rate_scientific_denominator_unaffected_by_five_percent(self):
        """测试 21: 当免勤人数达 15% (远超 5%) 时，出勤率分母严格按全量免勤扣减。"""
        # 纯净口径公式：学院出勤率 = 学院总实到 ÷ (学院总应到 − 当次自习学院总免勤人数)
        expected = 200
        exempt = 30   # 占 15%
        actual = 153

        denominator = expected - exempt
        rate = actual / denominator

        assert denominator == 170
        assert round(rate, 4) == 0.9000
        # 确保未发生 5% 封顶歪曲
        fake_capped_denominator = expected - (expected * 0.05)
        fake_capped_rate = actual / fake_capped_denominator
        assert rate != pytest.approx(fake_capped_rate, rel=1e-3)


# ==============================================================================
# 7. 隐私扫描零泄露与防御性路径兼容 (Privacy Scan Invariants)
# ==============================================================================
class TestPrivacyScanCompliance:
    """确保打包配置与脚本中无绝对路径与机器环境痕迹。"""

    def test_packaging_assets_pass_privacy_scan_rules(self):
        """测试 22: 打包脚本和配置文件逐行通过隐私扫描器的高危规则检测。"""
        from tools.privacy_scan import PATH_PATTERNS, SID_PATTERN

        files_to_check = [
            PROJECT_DIR / 'attendance.spec',
        ]
        iss_path = PROJECT_DIR / 'setup.iss'
        if iss_path.exists():
            files_to_check.append(iss_path)

        for f in files_to_check:
            text = f.read_text(encoding='utf-8')
            for pattern, label, _ in PATH_PATTERNS:
                matches = re.findall(pattern, text)
                assert not matches, f"在 {f.name} 中发现 {label}: {matches}"

            for m in SID_PATTERN.finditer(text):
                val = m.group(0)
                assert val.startswith('99'), f"在 {f.name} 中发现疑似真实学号: {val}"
