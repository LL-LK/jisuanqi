# -*- coding: utf-8 -*-
"""三早一晚免勤自动统计系统 · 全流程自动化构建驱动脚本 (build_all.py)

八大流水线阶段 (8-Stage Pipeline)：
  1. Pre-flight Checks: 环境与依赖自检 (Python 3.10+, 必需库, 源码文件)
  2. Icon Generation: 草莓小熊多尺寸 ICO 资产检查与自动生成
  3. Automated Testing Gate: 全量 pytest 自动化测试门禁 (零容忍失败)
  4. PyInstaller Executable Compilation: 单文件 EXE 编译与动态库裁剪
  5. Assemble Release Directory: 组装发布目录 (发布版/) 并生成目录 SHA-256
  6. Build Portable ZIP Package: 组装便携版 ZIP (三早一晚自动统计_vX_Windows便携版.zip)
  7. Inno Setup Compilation: 编译向导式安装包 (三早一晚自动统计_vX_Setup.exe)
  8. SHA256 Manifest & Verification: 统一计算发布包哈希并执行逐位校验

出勤率算法口径：学院总实到 / (学院总应到 - 学院总免勤)，纯净计算无 5% 封顶。
"""

from __future__ import annotations

import argparse
import hashlib
import os
import shutil
import subprocess
import sys
import time
import tarfile
import zipfile
from typing import Any, Dict, List, Optional, Tuple

# 强化 Windows/CI 终端编码：规避 cp1252 编码环境下的中文输出 UnicodeEncodeError
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# 路径基准解析 (无论从项目根目录还是 tools/ 执行均能正确锚定)
TOOLS_DIR = os.path.dirname(os.path.abspath(__file__))
DESKTOP_DIR = os.path.dirname(TOOLS_DIR)
if DESKTOP_DIR not in sys.path:
    sys.path.insert(0, DESKTOP_DIR)

try:
    from core import __version__ as DEFAULT_VERSION
except ImportError:
    DEFAULT_VERSION = '1.0.6'


def log_stage(stage_num: int, total_stages: int, title: str) -> None:
    """输出阶段分隔标题栏。"""
    bar = "=" * 70
    print("\n" + bar)
    print(f" [STAGE {stage_num}/{total_stages}] {title}")
    print(bar)


def log_info(msg: str) -> None:
    print(f"[INFO] {msg}")


def log_warn(msg: str) -> None:
    print(f"[WARN] {msg}")


def log_error(msg: str) -> None:
    print(f"[ERROR] {msg}")


def log_success(msg: str) -> None:
    print(f"[SUCCESS] {msg}")


def sha256_file(path: str) -> str:
    """计算大文件 SHA-256 哈希值（流式分块读取，防止内存暴涨）。"""
    h = hashlib.sha256()
    with open(path, 'rb') as fh:
        for chunk in iter(lambda: fh.read(65536), b''):
            h.update(chunk)
    return h.hexdigest()


calculate_sha256 = sha256_file


def write_manifest(files: List[Any], manifest_path: Any) -> str:
    """写入文件列表的 SHA256 清单文件（两空格标准格式）。"""
    manifest_lines = []
    for f in files:
        f_str = str(f)
        if os.path.isfile(f_str):
            fname = os.path.basename(f_str)
            digest = sha256_file(f_str)
            manifest_lines.append(f"{digest}  {fname}")
    manifest_content = '\n'.join(manifest_lines) + '\n'
    manifest_path_str = str(manifest_path)
    os.makedirs(os.path.dirname(os.path.abspath(manifest_path_str)), exist_ok=True)
    with open(manifest_path_str, 'w', encoding='utf-8') as fh:
        fh.write(manifest_content)
    return manifest_path_str


def verify_manifest(manifest_path: str) -> int:
    """校验清单中所有文件的 SHA256 哈希值。成功返回 0，失败返回 1。"""
    manifest_path_str = str(manifest_path)
    if not os.path.exists(manifest_path_str):
        log_error(f"清单文件不存在: {manifest_path_str}")
        return 1
    base_dir = os.path.dirname(os.path.abspath(manifest_path_str))
    with open(manifest_path_str, 'r', encoding='utf-8') as fh:
        lines = [line.strip() for line in fh if line.strip()]
    for line in lines:
        parts = line.split(maxsplit=1)
        if len(parts) != 2:
            continue
        expected_hash, fname = parts
        fname = fname.lstrip('*').strip()
        fpath = os.path.join(base_dir, fname)
        if not os.path.exists(fpath):
            log_error(f"清单中引用的文件不存在: {fpath}")
            return 1
        actual_hash = sha256_file(fpath)
        if actual_hash.lower() != expected_hash.lower():
            log_error(f"文件哈希校验不匹配: {fname} (预期 {expected_hash}, 实际 {actual_hash})")
            return 1
    return 0


verify_release = verify_manifest


# ------------------------------------------------------------------ Stage 1
def stage_1_preflight(args: argparse.Namespace) -> None:
    """阶段 1：环境与依赖自检。"""
    log_stage(1, 8, "Pre-flight Environment & Dependency Checks")

    # 1. Python 版本
    py_ver = sys.version_info
    log_info(f"Python 运行环境: {py_ver.major}.{py_ver.minor}.{py_ver.micro}")
    if py_ver < (3, 10):
        log_error("Python 3.10 或更高版本为强制要求！")
        sys.exit(1)

    # 2. 核心模块可用性
    required_modules = ['PySide6', 'openpyxl', 'pytest', 'PyInstaller']
    missing_modules = []
    for mod in required_modules:
        try:
            __import__(mod)
        except ImportError:
            missing_modules.append(mod)

    if missing_modules:
        log_error(f"缺失必要 Python 依赖库: {', '.join(missing_modules)}")
        log_error("请先执行: pip install -r requirements.txt")
        sys.exit(1)
    log_info("核心依赖库校验通过: PySide6, openpyxl, pytest, PyInstaller")

    # 3. 必需源码文件完备性
    required_files = [
        os.path.join(DESKTOP_DIR, 'app.py'),
        os.path.join(DESKTOP_DIR, 'attendance.spec'),
        os.path.join(DESKTOP_DIR, '免勤词典.xlsx'),
        os.path.join(DESKTOP_DIR, '使用说明.md'),
        os.path.join(DESKTOP_DIR, 'README.md'),
        os.path.join(DESKTOP_DIR, '隐私扫描报告.md'),
    ]
    for fp in required_files:
        if not os.path.exists(fp):
            log_error(f"缺失关键工程文件: {os.path.basename(fp)}")
            sys.exit(1)
    log_info("工程核心源文件完备。")


# ------------------------------------------------------------------ Stage 2
def stage_2_ensure_icon(args: argparse.Namespace) -> None:
    """阶段 2：草莓小熊图标资产检查与生成。"""
    log_stage(2, 8, "Multi-resolution Bear Icon Verification & Generation")
    ico_path = os.path.join(DESKTOP_DIR, 'assets', 'bear.ico')
    gen_script = os.path.join(TOOLS_DIR, 'generate_icon.py')

    if os.path.exists(ico_path) and os.path.getsize(ico_path) > 1024:
        log_info(f"检测到现存小熊图标: assets/bear.ico ({os.path.getsize(ico_path)} 字节)")
        return

    log_info("assets/bear.ico 缺失或为空，准备调用矢量生成器...")
    if os.path.exists(gen_script):
        cmd = [sys.executable, gen_script]
        ret = subprocess.run(cmd, cwd=DESKTOP_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace')
        if ret.returncode != 0:
            log_error("generate_icon.py 执行失败:")
            print(ret.stderr)
            sys.exit(1)
        log_success("已成功生成 assets/bear.ico")
    else:
        log_warn("tools/generate_icon.py 暂不存在，尝试动态生成...")
        try:
            from tools.generate_icon import generate_bear_ico
            generate_bear_ico(ico_path)
            log_success(f"已动态生成: {ico_path}")
        except Exception as exc:
            log_warn(f"动态生成失败: {exc}，PyInstaller 将回退为默认图标。")


# ------------------------------------------------------------------ Stage 3
def stage_3_test_gate(args: argparse.Namespace) -> None:
    """阶段 3：自动化测试套件门禁 (pytest)。"""
    log_stage(3, 8, "Automated Testing Gate (pytest)")
    if args.skip_tests:
        log_warn("用户传入 --skip-tests，跳过自动化测试套件。")
        return

    log_info("启动 pytest 全量测试套件 (包含 258+ 项回归与性质测试)...")
    t0 = time.time()
    cmd = [sys.executable, "-m", "pytest", "tests", "-q"]
    ret = subprocess.run(cmd, cwd=DESKTOP_DIR)
    elapsed = time.time() - t0

    if ret.returncode != 0:
        log_error(f"自动化测试门禁失败！(耗时 {elapsed:.2f}s)")
        log_error("构建立即中止。请修复全部测试用例后再行打包！")
        sys.exit(1)

    log_success(f"测试套件 100% 绿色通过！(耗时 {elapsed:.2f}s)")


# ------------------------------------------------------------------ Stage 4
def stage_4_pyinstaller(args: argparse.Namespace) -> None:
    """阶段 4：PyInstaller 单文件编译。"""
    log_stage(4, 8, "PyInstaller Executable Compilation")
    target_exe = os.path.join(DESKTOP_DIR, 'dist', 'AttendanceStats.exe')

    if args.skip_pyi:
        if os.path.exists(target_exe) and os.path.getsize(target_exe) > 10 * 1024 * 1024:
            log_warn(f"用户传入 --skip-pyi，复用现存 EXE: {target_exe} ({os.path.getsize(target_exe) / 1048576:.1f} MB)")
            return
        else:
            log_error("--skip-pyi 生效但 dist/AttendanceStats.exe 不存在或损坏！无法继续。")
            sys.exit(1)

    log_info("执行 PyInstaller 编译流程 (使用 attendance.spec)...")
    spec_path = os.path.join(DESKTOP_DIR, 'attendance.spec')
    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", spec_path]
    t0 = time.time()
    ret = subprocess.run(cmd, cwd=DESKTOP_DIR)
    elapsed = time.time() - t0

    if ret.returncode != 0 or not os.path.exists(target_exe):
        log_error(f"PyInstaller 编译失败！(耗时 {elapsed:.2f}s)")
        sys.exit(1)

    log_success(f"单文件 EXE 编译完成: {target_exe} ({os.path.getsize(target_exe) / 1048576:.1f} MB, 耗时 {elapsed:.2f}s)")


# ------------------------------------------------------------------ Stage 5
def stage_5_assemble_release(args: argparse.Namespace) -> str:
    """阶段 5：组装发布目录 (发布版/) 并生成内部 SHA256。"""
    log_stage(5, 8, "Assemble Release Directory (发布版/)")
    rel_dir = os.path.join(DESKTOP_DIR, '发布版')
    os.makedirs(rel_dir, exist_ok=True)

    # 清理旧文件
    for name in os.listdir(rel_dir):
        fp = os.path.join(rel_dir, name)
        if os.path.isfile(fp):
            os.remove(fp)
            log_info(f"清理旧产物: {name}")

    # 复制发布清单项
    items = [
        (os.path.join(DESKTOP_DIR, 'dist', 'AttendanceStats.exe'), '三早一晚免勤自动统计.exe'),
        (os.path.join(DESKTOP_DIR, '免勤词典.xlsx'), '免勤词典.xlsx'),
        (os.path.join(DESKTOP_DIR, '使用说明.md'), '使用说明.md'),
        (os.path.join(DESKTOP_DIR, 'README.md'), 'README.md'),
        (os.path.join(DESKTOP_DIR, '隐私扫描报告.md'), '隐私扫描报告.md'),
    ]

    for src, dst_name in items:
        if not os.path.exists(src):
            log_error(f"必需的组装文件缺失: {src}")
            sys.exit(1)
        shutil.copy2(src, os.path.join(rel_dir, dst_name))
        log_info(f"已装配文件: {dst_name}")

    # 生成发布版/SHA256.txt
    release_files = [os.path.join(rel_dir, name) for name in sorted(os.listdir(rel_dir)) if name != 'SHA256.txt']
    manifest_path = os.path.join(rel_dir, 'SHA256.txt')
    write_manifest(release_files, manifest_path)

    log_success(f"发布版目录装配完成: {rel_dir} (含 SHA256.txt)")
    return rel_dir


# ------------------------------------------------------------------ Stage 6
def stage_6_build_portable(args: argparse.Namespace, version: str, dist_dir: str) -> Optional[str]:
    """阶段 6：构建 Windows 便携版 ZIP 压缩包。"""
    log_stage(6, 8, "Build Portable ZIP Package")
    if args.skip_portable:
        log_warn("用户传入 --skip-portable，跳过便携版 ZIP 打包。")
        return None

    rel_dir = os.path.join(DESKTOP_DIR, '发布版')
    os.makedirs(dist_dir, exist_ok=True)

    zip_filename = "三早一晚免勤自动统计_Windows便携版.zip"
    zip_path = os.path.join(dist_dir, zip_filename)

    # 清理旧同名包
    if os.path.exists(zip_path):
        os.remove(zip_path)

    log_info(f"开始压缩打包: {zip_filename}...")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for name in sorted(os.listdir(rel_dir)):
            fp = os.path.join(rel_dir, name)
            if os.path.isfile(fp) and not name.lower().endswith('.zip'):
                arcname = f"三早一晚免勤自动统计/{name}"
                zf.write(fp, arcname=arcname)

    zip_size_mb = os.path.getsize(zip_path) / 1048576
    digest = sha256_file(zip_path)

    # 写入单文件 sidecar .sha256
    with open(zip_path + '.sha256', 'w', encoding='utf-8') as fh:
        fh.write(f"{digest}  {zip_filename}\n")

    log_success(f"便携版 ZIP 产出成功: {zip_filename} ({zip_size_mb:.1f} MB)")
    log_info(f"便携版 SHA-256: {digest}")
    return zip_path


# ------------------------------------------------------------------ Stage 7
def find_iscc(explicit_path: Optional[str]) -> Optional[str]:
    """多路径探测 Inno Setup 编译器 ISCC.exe（防御性环境变量探测，避免硬编码盘符路径）。"""
    if explicit_path:
        if os.path.isfile(explicit_path):
            return explicit_path
        log_warn(f"指定的 --iscc-path 不存在: {explicit_path}")

    # 1. PATH 搜索
    which_iscc = shutil.which('ISCC.exe') or shutil.which('iscc')
    if which_iscc:
        return which_iscc

    # 2. 环境变量动态拼接探测候选路径（杜绝硬编码盘符）
    env_keys = [
        ('ProgramFiles(x86)', ['Inno Setup 6', 'ISCC.exe']),
        ('ProgramFiles', ['Inno Setup 6', 'ISCC.exe']),
        ('LOCALAPPDATA', ['Programs', 'Inno Setup 6', 'ISCC.exe']),
        ('ChocolateyInstall', ['bin', 'ISCC.exe']),
    ]

    for env_var, sub_paths in env_keys:
        base = os.environ.get(env_var)
        if base:
            candidate = os.path.join(base, *sub_paths)
            if os.path.isfile(candidate):
                return candidate

    return None


def stage_7_build_inno_setup(args: argparse.Namespace, version: str, dist_dir: str) -> Optional[str]:
    """阶段 7：调用 Inno Setup 编译 Setup 安装包。"""
    log_stage(7, 8, "Inno Setup Compiler (setup.iss)")
    if args.skip_setup:
        log_warn("用户传入 --skip-setup，跳过 Inno Setup 安装包编译。")
        return None

    iscc_path = find_iscc(args.iscc_path)
    is_ci = os.environ.get('CI') == 'true' or os.environ.get('GITHUB_ACTIONS') == 'true'

    if not iscc_path:
        msg = [
            "=" * 70,
            "[WARN] 未在宿主机检测到 Inno Setup 编译器 (ISCC.exe)！",
            "-" * 70,
            "• 便携版 ZIP 构建不受影响，已正常完成。",
            "• 若需编译 Windows 安装包 (Setup.exe)，请按以下方式配置：",
            "    1. 官网下载安装: https://jrsoftware.org/isdl.php",
            "    2. Chocolatey:  choco install innosetup",
            "    3. 命令行指定:  python tools/build_all.py --iscc-path \"<InnoSetupDir>\\ISCC.exe\"",
            "=" * 70,
        ]
        if is_ci:
            log_error("CI/CD 环境下必须安装 Inno Setup！构建阻断。")
            sys.exit(1)
        else:
            print("\n".join(msg))
            return None

    log_info(f"使用 Inno Setup 编译器: {iscc_path}")
    iss_file = os.path.join(DESKTOP_DIR, 'setup.iss')
    if not os.path.exists(iss_file):
        log_error(f"缺失 Inno Setup 脚本: {iss_file}")
        sys.exit(1)

    cmd = [
        iscc_path,
        f"/DMyAppVersion={version}",
        f"/O{dist_dir}",
        iss_file,
    ]
    t0 = time.time()
    ret = subprocess.run(cmd, cwd=DESKTOP_DIR, capture_output=True, text=True, encoding='utf-8', errors='replace')
    elapsed = time.time() - t0

    setup_filename = "三早一晚免勤自动统计_Setup.exe"
    setup_path = os.path.join(dist_dir, setup_filename)

    if ret.returncode != 0 or not os.path.exists(setup_path):
        log_error(f"Inno Setup 编译失败！(耗时 {elapsed:.2f}s)")
        print(ret.stderr or ret.stdout)
        if is_ci:
            sys.exit(1)
        return None

    setup_size_mb = os.path.getsize(setup_path) / 1048576
    digest = sha256_file(setup_path)

    # 写入单文件 sidecar .sha256
    with open(setup_path + '.sha256', 'w', encoding='utf-8') as fh:
        fh.write(f"{digest}  {setup_filename}\n")

    log_success(f"安装包 Setup.exe 产出成功: {setup_filename} ({setup_size_mb:.1f} MB, 耗时 {elapsed:.2f}s)")
    log_info(f"Setup.exe SHA-256: {digest}")
    return setup_path


# ------------------------------------------------------------------ Stage 6B/6C/6D: Multi-OS & Clean Source
EXCLUDED_SOURCE_DIRS = {
    'tests', 'tests_e2e', '_synthetic', 'data_sample', '.pytest_cache',
    '.hypothesis', '__pycache__', '.git', '.agents', '.buildenv', '.venv',
    'venv', 'build', 'dist', 'dist_linux', '发布版', '发布包', 'node_modules', 'docs',
}
EXCLUDED_SOURCE_FILES = {
    'pytest.ini', 'requirements-dev.txt', 'e2e_report.json',
    'test_double_count.py',
}


def _is_test_or_excluded_relpath(rel_parts: Tuple[str, ...]) -> bool:
    if any(p in EXCLUDED_SOURCE_DIRS for p in rel_parts[:-1]):
        return True
    fname = rel_parts[-1]
    if fname in EXCLUDED_SOURCE_FILES:
        return True
    low = fname.lower()
    if low.startswith('test_') or low.endswith('_test.py') or low.endswith(('.pyc', '.pyo', '.db', '.db-wal', '.db-shm')):
        return True
    return False


def build_clean_source_packages(dist_dir: str) -> List[str]:
    """构建不含任何测试组件的纯净源码包 (.zip 与 .tar.gz)。"""
    os.makedirs(dist_dir, exist_ok=True)
    zip_name = "三早一晚免勤自动统计_Source源码包.zip"
    tar_name = "三早一晚免勤自动统计_Source源码包.tar.gz"
    zip_path = os.path.join(dist_dir, zip_name)
    tar_path = os.path.join(dist_dir, tar_name)

    for p in (zip_path, tar_path):
        if os.path.exists(p):
            os.remove(p)

    source_files: List[Tuple[str, str]] = []
    for root, dirs, files in os.walk(DESKTOP_DIR):
        dirs[:] = sorted(d for d in dirs if d not in EXCLUDED_SOURCE_DIRS)
        for fname in sorted(files):
            full_p = os.path.join(root, fname)
            rel_p = os.path.relpath(full_p, DESKTOP_DIR)
            parts = tuple(rel_p.replace("\\", "/").split("/"))
            if _is_test_or_excluded_relpath(parts):
                continue
            arc_p = "三早一晚免勤自动统计_Source/" + "/".join(parts)
            source_files.append((full_p, arc_p))

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        for full_p, arc_p in source_files:
            zf.write(full_p, arcname=arc_p)

    with tarfile.open(tar_path, 'w:gz') as tf:
        for full_p, arc_p in source_files:
            tf.add(full_p, arcname=arc_p)

    for pkg_p, pkg_n in ((zip_path, zip_name), (tar_path, tar_name)):
        digest = sha256_file(pkg_p)
        with open(pkg_p + '.sha256', 'w', encoding='utf-8') as fh:
            fh.write(f"{digest}  {pkg_n}\n")
        log_success(f"纯净源码包产出成功 (零测试组件): {pkg_n} ({os.path.getsize(pkg_p) / 1048576:.2f} MB)")

    return [zip_path, tar_path]


def build_macos_package(dist_dir: str, version: str) -> str:
    """构建 macOS Universal (.app Bundle) 便携发布包 (.tar.gz)，不含任何测试组件。"""
    os.makedirs(dist_dir, exist_ok=True)
    tar_name = "三早一晚免勤自动统计_macOS便携版.tar.gz"
    tar_path = os.path.join(dist_dir, tar_name)
    if os.path.exists(tar_path):
        os.remove(tar_path)

    import io
    prefix = "三早一晚免勤自动统计_macOS"
    app_prefix = f"{prefix}/三早一晚免勤自动统计.app/Contents"

    info_plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleDisplayName</key>
    <string>三早一晚免勤自动统计</string>
    <key>CFBundleExecutable</key>
    <string>AttendanceStats</string>
    <key>CFBundleIdentifier</key>
    <string>org.sanzaowan.attendancestats</string>
    <key>CFBundleName</key>
    <string>AttendanceStats</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleShortVersionString</key>
    <string>{version}</string>
    <key>CFBundleVersion</key>
    <string>{version}</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
""".encode("utf-8")

    macos_launcher = """#!/usr/bin/env bash
set -e
DIR="$(cd "$(dirname "$0")/../Resources/app" && pwd)"
export PYTHONPATH="$DIR:$PYTHONPATH"
if command -v python3 >/dev/null 2>&1; then
    exec python3 "$DIR/app.py" "$@"
else
    exec python "$DIR/app.py" "$@"
fi
""".encode("utf-8")

    run_command = """#!/usr/bin/env bash
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
APP_BIN="$HERE/三早一晚免勤自动统计.app/Contents/MacOS/AttendanceStats"
chmod +x "$APP_BIN" 2>/dev/null || true
exec "$APP_BIN" "$@"
""".encode("utf-8")

    def _add_bytes(tf: tarfile.TarFile, arcname: str, data: bytes, mode: int = 0o644) -> None:
        ti = tarfile.TarInfo(name=arcname)
        ti.size = len(data)
        ti.mtime = int(time.time())
        ti.mode = mode
        tf.addfile(ti, io.BytesIO(data))

    with tarfile.open(tar_path, 'w:gz') as tf:
        _add_bytes(tf, f"{app_prefix}/Info.plist", info_plist, mode=0o644)
        _add_bytes(tf, f"{app_prefix}/MacOS/AttendanceStats", macos_launcher, mode=0o755)
        _add_bytes(tf, f"{prefix}/run_mac.command", run_command, mode=0o755)

        # 打包核心运行时文件至 Contents/Resources/app（严格排除测试组件）
        for rel_item in ('app.py', 'cli.py', 'requirements.txt', '免勤词典.xlsx', '使用说明.md', 'README.md', '隐私扫描报告.md'):
            fp = os.path.join(DESKTOP_DIR, rel_item)
            if os.path.exists(fp):
                tf.add(fp, arcname=f"{app_prefix}/Resources/app/{rel_item}")
                if rel_item in ('免勤词典.xlsx', '使用说明.md', 'README.md', '隐私扫描报告.md'):
                    tf.add(fp, arcname=f"{prefix}/{rel_item}")

        for folder in ('core', 'assets'):
            folder_p = os.path.join(DESKTOP_DIR, folder)
            for root, dirs, files in os.walk(folder_p):
                dirs[:] = [d for d in dirs if d != '__pycache__']
                for fname in sorted(files):
                    if fname.endswith(('.pyc', '.pyo')):
                        continue
                    full_p = os.path.join(root, fname)
                    rel_p = os.path.relpath(full_p, DESKTOP_DIR).replace("\\", "/")
                    tf.add(full_p, arcname=f"{app_prefix}/Resources/app/{rel_p}")

    digest = sha256_file(tar_path)
    with open(tar_path + '.sha256', 'w', encoding='utf-8') as fh:
        fh.write(f"{digest}  {tar_name}\n")
    log_success(f"macOS 便携版产出成功: {tar_name} ({os.path.getsize(tar_path) / 1048576:.2f} MB)")
    return tar_path


def build_huawei_openeuler_package(dist_dir: str, version: str) -> List[str]:
    """构建华为平台 (Huawei Kunpeng / EulerOS / openEuler / UOS / Kylin) 与 Linux 便携版 (.tar.gz)。"""
    os.makedirs(dist_dir, exist_ok=True)
    linux_tar = os.path.join(dist_dir, "三早一晚免勤自动统计_Linux便携版.tar.gz")
    huawei_tar = os.path.join(dist_dir, "三早一晚免勤自动统计_华为openEuler便携版.tar.gz")

    linux_bin = os.path.join(DESKTOP_DIR, 'dist_linux', 'AttendanceStats')
    if not os.path.exists(linux_bin):
        linux_bin = os.path.join(DESKTOP_DIR, 'dist', 'AttendanceStats')

    import io

    huawei_launcher = """#!/usr/bin/env bash
# 华为平台 (openEuler / EulerOS / Kylin / UOS) & 通用 Linux 启动器
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
export LD_LIBRARY_PATH="$HERE/lib:$HERE:$LD_LIBRARY_PATH"
export QT_AUTO_SCREEN_SCALE_FACTOR=1
if [ -z "$DISPLAY" ] && [ -z "$WAYLAND_DISPLAY" ]; then
    export QT_QPA_PLATFORM=offscreen
fi
if [ -x "$HERE/三早一晚免勤自动统计" ]; then
    exec "$HERE/三早一晚免勤自动统计" "$@"
elif command -v python3 >/dev/null 2>&1; then
    exec python3 "$HERE/runtime/app.py" "$@"
else
    exec python "$HERE/runtime/app.py" "$@"
fi
""".encode("utf-8")

    desktop_entry = f"""[Desktop Entry]
Type=Application
Version={version}
Name=三早一晚免勤自动统计
Comment=三早一晚免勤自动统计系统 (支持华为 openEuler / Linux 平台)
Exec=./run_huawei_linux.sh
Terminal=false
Categories=Office;Education;Utility;
""".encode("utf-8")

    def _add_bytes(tf: tarfile.TarFile, arcname: str, data: bytes, mode: int = 0o644) -> None:
        ti = tarfile.TarInfo(name=arcname)
        ti.size = len(data)
        ti.mtime = int(time.time())
        ti.mode = mode
        tf.addfile(ti, io.BytesIO(data))

    built_paths = []
    for target_tar in (linux_tar, huawei_tar):
        if os.path.exists(target_tar):
            os.remove(target_tar)
        prefix = "三早一晚免勤自动统计"
        manifest_entries = []
        with tarfile.open(target_tar, 'w:gz') as tf:
            if os.path.exists(linux_bin):
                tf.add(linux_bin, arcname=f"{prefix}/三早一晚免勤自动统计")
                manifest_entries.append(f"{sha256_file(linux_bin)}  三早一晚免勤自动统计")
            _add_bytes(tf, f"{prefix}/run_huawei_linux.sh", huawei_launcher, mode=0o755)
            _add_bytes(tf, f"{prefix}/attendance.desktop", desktop_entry, mode=0o644)

            for doc_name in ('免勤词典.xlsx', '使用说明.md', 'README.md', '隐私扫描报告.md'):
                fp = os.path.join(DESKTOP_DIR, doc_name)
                if os.path.exists(fp):
                    tf.add(fp, arcname=f"{prefix}/{doc_name}")
                    manifest_entries.append(f"{sha256_file(fp)}  {doc_name}")

            # 附带轻量纯 Python 跨架构 runtime 以支持鲲鹏 ARM64 / 昇腾 / LoongArch64 原生运行
            for rel_item in ('app.py', 'cli.py', 'requirements.txt'):
                fp = os.path.join(DESKTOP_DIR, rel_item)
                if os.path.exists(fp):
                    tf.add(fp, arcname=f"{prefix}/runtime/{rel_item}")
            for folder in ('core', 'assets'):
                folder_p = os.path.join(DESKTOP_DIR, folder)
                for root, dirs, files in os.walk(folder_p):
                    dirs[:] = [d for d in dirs if d != '__pycache__']
                    for fname in sorted(files):
                        if fname.endswith(('.pyc', '.pyo')):
                            continue
                        full_p = os.path.join(root, fname)
                        rel_p = os.path.relpath(full_p, DESKTOP_DIR).replace("\\", "/")
                        tf.add(full_p, arcname=f"{prefix}/runtime/{rel_p}")

            sha_bytes = ("\n".join(manifest_entries) + "\n").encode("utf-8")
            _add_bytes(tf, f"{prefix}/SHA256.txt", sha_bytes, mode=0o644)

        tar_name = os.path.basename(target_tar)
        digest = sha256_file(target_tar)
        with open(target_tar + '.sha256', 'w', encoding='utf-8') as fh:
            fh.write(f"{digest}  {tar_name}\n")
        log_success(f"Linux/华为平台发布包产出成功: {tar_name} ({os.path.getsize(target_tar) / 1048576:.2f} MB)")
        built_paths.append(target_tar)

    return built_paths


# ------------------------------------------------------------------ Stage 8
def stage_8_manifest_and_verify(args: argparse.Namespace, dist_dir: str, version: str) -> None:
    """阶段 8：生成统一发布包 SHA256.txt 并执行自校验。"""
    log_stage(8, 8, "SHA-256 Manifest Calculation & Release Verification")

    deliverables: List[Tuple[str, str, float]] = []
    files_to_hash: List[str] = []

    for name in sorted(os.listdir(dist_dir)):
        fp = os.path.join(dist_dir, name)
        # 只收录终版交付物 (.zip 与 .exe)，排除 sidecar 与旧 manifest
        if os.path.isfile(fp) and (name.lower().endswith('.zip') or name.lower().endswith('.exe') or name.lower().endswith('.tar.gz')):
            digest = sha256_file(fp)
            size_mb = os.path.getsize(fp) / 1048576
            deliverables.append((name, digest, size_mb))
            files_to_hash.append(fp)

    if not deliverables:
        log_warn("发布包目录下未检测到交付物！")
        return

    # 写入发布包/SHA256.txt
    manifest_path = os.path.join(dist_dir, 'SHA256.txt')
    write_manifest(files_to_hash, manifest_path)
    log_info(f"已生成统合校验清单: {manifest_path}")

    # 自校验
    ret = verify_manifest(manifest_path)
    if ret != 0:
        log_error("自校验发现哈希不匹配或文件缺失！")
        sys.exit(1)

    for name, actual_hash, _ in deliverables:
        log_info(f"[VERIFIED] {name} -> {actual_hash[:16]}... OK")

    # 打印最终成果表格
    print("\n" + "=" * 70)
    print("  🎉 构建与打包全部完成！(Release Deliverables Ready)")
    print("=" * 70)
    print("  当前版本: 最新版 (草莓小熊版)")
    print(f"  输出目录: {dist_dir}")
    print("-" * 70)
    print(f"  {'产物文件名':<38} {'大小':<10} {'SHA-256 摘要'}")
    print("-" * 70)
    for name, digest, size_mb in deliverables:
        print(f"  {name:<36} {size_mb:>6.1f} MB  {digest[:16]}...")
    print("=" * 70 + "\n")


# ------------------------------------------------------------------ Parser & Main Entry
def build_parser() -> argparse.ArgumentParser:
    """构建命令行参数解析器。"""
    parser = argparse.ArgumentParser(
        description="三早一晚免勤自动统计系统 · 全流程自动化构建驱动脚本",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument('--skip-tests', action='store_true', help="跳过全量 pytest 测试套件")
    parser.add_argument('--skip-pyi', action='store_true', help="跳过 PyInstaller 编译过程 (复用现存 dist/ 产物)")
    parser.add_argument('--skip-portable', action='store_true', help="跳过构建便携版 ZIP 压缩包")
    parser.add_argument('--skip-setup', action='store_true', help="跳过构建 Inno Setup 安装包")
    parser.add_argument('--iscc-path', type=str, default=None, help="显式指定 ISCC.exe 的物理绝对路径")
    parser.add_argument('--clean', action='store_true', help="构建前清理旧缓存 (build/, dist/, 发布版/, 发布包/)")
    parser.add_argument('--version', type=str, default=DEFAULT_VERSION, help=f"发布版本号 (默认: {DEFAULT_VERSION})")
    parser.add_argument('--dist-dir', type=str, default=None, help="产物输出目录 (默认为 发布包/)")
    return parser


get_parser = build_parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    version = args.version.lstrip('v')
    dist_dir = os.path.abspath(args.dist_dir) if args.dist_dir else os.path.join(DESKTOP_DIR, '发布包')

    if args.clean:
        log_info("正在执行全面深度清理 (--clean)...")
        for folder in ['build', 'dist', 'dist_linux', '发布版', '发布包']:
            p = os.path.join(DESKTOP_DIR, folder)
            if os.path.exists(p):
                shutil.rmtree(p, ignore_errors=True)
                log_info(f"已清理目录: {folder}/")

    t_start = time.time()
    try:
        stage_1_preflight(args)
        stage_2_ensure_icon(args)
        stage_3_test_gate(args)
        stage_4_pyinstaller(args)
        stage_5_assemble_release(args)
        stage_6_build_portable(args, version, dist_dir)
        if not getattr(args, 'skip_portable', False):
            build_clean_source_packages(dist_dir)
            build_macos_package(dist_dir, version)
            build_huawei_openeuler_package(dist_dir, version)
        stage_7_build_inno_setup(args, version, dist_dir)
        stage_8_manifest_and_verify(args, dist_dir, version)
    except KeyboardInterrupt:
        log_error("\n构建被用户手动中断！")
        return 130
    except Exception as e:
        log_error(f"构建遇到未处理异常: {e}")
        import traceback
        traceback.print_exc()
        return 1

    total_time = time.time() - t_start
    log_success(f"构建总耗时: {total_time:.2f} 秒。一切就绪！\n")
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
