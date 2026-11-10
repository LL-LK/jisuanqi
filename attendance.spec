# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller packaging specification for Local Attendance Tool.

Architecture:
  - Output: Single-file executable (AttendanceStats.exe) in dist/
  - Multi-resolution Vector Icon: assets/attendance.ico embedded into PE header
  - Data Files: 免勤词典.xlsx, 使用说明.md, assets/attendance.ico
  - Deep DLL Pruning: Filters out ~39.2MB unused Qt6 dynamic libraries
    (Qt6VirtualKeyboard, Qt6Quick, Qt6Qml, Qt6Pdf, Qt6OpenGL, opengl32sw, d3dcompiler_47)
  - Security & Stability: upx=False (eliminates AV false-positives & Win10 0xc0000005 crashes)
  - Window Mode: console=False (no terminal black box)
  - PE Metadata: Windows version resource information
"""

import os
import sys

block_cipher = None
PROJECT_DIR = os.path.dirname(os.path.abspath(SPEC))

# 1. Assets and Data Files
icon_path = os.path.join(PROJECT_DIR, 'assets', 'attendance.ico')
datas = [
    ('免勤词典.xlsx', '.'),
    ('使用说明.md', '.'),
    ('assets/strawberry-bear.png', 'assets'),
]
if os.path.exists(icon_path):
    datas.append(('assets/attendance.ico', 'assets'))

# 2. Hidden Imports for Dynamic Modules & openpyxl Internals
hiddenimports = [
    'openpyxl',
    'openpyxl.cell._writer',
    'openpyxl.styles.stylesheet',
    'core',
    'core.pipeline',
    'core.calc',
    'core.dictionary',
    'core.export',
    'core.io_detail',
    'core.io_exempt',
    'core.parsing',
    'core.model',
    'core.excel_safety',
    'core.storage',
    'core.exempt_report',
    'core.summary_template',
    'core.temp_manager',
    'ui',
    'ui.theme',
    'ui.icons',
    'ui.widgets',
    'ui.dialogs',
]

# 3. Excludes: Heavyweight unused Qt modules and dev/scientific dependencies
excludes = [
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets', 'PySide6.QtWebEngineQuick',
    'PySide6.QtQuick', 'PySide6.QtQuick3D', 'PySide6.QtQml', 'PySide6.QtQuickWidgets',
    'PySide6.Qt3DCore', 'PySide6.Qt3DRender', 'PySide6.Qt3DAnimation', 'PySide6.Qt3DExtras',
    'PySide6.Qt3DInput', 'PySide6.Qt3DLogic', 'PySide6.Qt3DPhysics',
    'PySide6.QtCharts', 'PySide6.QtDataVisualization', 'PySide6.QtMultimedia',
    'PySide6.QtMultimediaWidgets', 'PySide6.QtBluetooth', 'PySide6.QtNfc', 'PySide6.QtPositioning',
    'PySide6.QtSql', 'PySide6.QtTest', 'PySide6.QtHelp', 'PySide6.QtDesigner',
    'PySide6.QtSerialPort', 'PySide6.QtWebSockets', 'PySide6.QtWebChannel',
    'PySide6.QtSpatialAudio', 'PySide6.QtPdf', 'PySide6.QtPdfWidgets',
    'PySide6.QtOpenGL', 'PySide6.QtOpenGLWidgets', 'PySide6.QtSvgWidgets',
    'PySide6.QtVirtualKeyboard',
    'tkinter', 'unittest', 'pydoc', 'doctest', 'pytest', 'hypothesis',
    'numpy', 'pandas', 'scipy', 'matplotlib', 'PIL',
    'setuptools', 'pip',
]

linux_binaries = []
if sys.platform.startswith('linux'):
    lib_dir = os.path.join(sys.prefix, 'lib')
    for so_name in ('libOpenGL.so.0', 'libGLX.so.0', 'libGLdispatch.so.0', 'libEGL.so.1', 'libGL.so.1'):
        so_path = os.path.join(lib_dir, so_name)
        if os.path.exists(so_path):
            linux_binaries.append((so_path, '.'))

a = Analysis(
    ['app.py'],
    pathex=[PROJECT_DIR],
    binaries=linux_binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# 4. Binary Pruning: Remove ~39.2MB of unreferenced Qt C++ DLLs
prune_keywords = [
    'qtvirtualkeyboard',
    'qt6virtualkeyboard',
    'qt6quick',
    'qt6qml',
    'qt6pdf',
    'qt6opengl',
    'opengl32sw',
    'd3dcompiler_47',
]
a.binaries = [
    b for b in a.binaries
    if not any(kw in b[0].lower() for kw in prune_keywords)
]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# 5. Version Information Metadata
version_file = os.path.join(PROJECT_DIR, 'file_version_info.txt')
version_param = version_file if (sys.platform == 'win32' and os.path.exists(version_file)) else None

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AttendanceStats',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon_path if os.path.exists(icon_path) else 'assets/attendance.ico',
    version=version_param,
)
