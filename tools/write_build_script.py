# -*- coding: utf-8 -*-
"""把 build_windows.bat 写成 cmd.exe 能正确解析的编码。

cmd.exe 按系统 ANSI 代码页（简体中文环境为 GBK/936）读取 .bat；
若文件是 UTF-8，中文会被解成乱码，个别字节组合会被当成引号或括号，
导致整个脚本解析失败（"xx was unexpected at this time"）。

因此：写 GBK + CRLF，且不���用 chcp。
"""

from __future__ import annotations

import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BAT = r'''@echo off
chcp 936 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo ============================================================
echo   Sanzao Wanwai Attendance Statistics - Windows Build
echo   三早一晚免勤自动统计系统 - Windows 打包脚本
echo ============================================================
echo.

echo [1/5] Checking Python ...
where py >nul 2>nul
if errorlevel 1 goto :try_python
set PY=py -3
goto :python_found
:try_python
where python >nul 2>nul
if errorlevel 1 goto :no_python
set PY=python
goto :python_found
:no_python
echo [ERROR] Python not found. Please install Python 3.10+ and tick "Add to PATH".
exit /b 1
:python_found
%PY% -c "import sys; sys.exit(0 if sys.version_info >= (3,10) else 1)"
if errorlevel 1 (
    echo [ERROR] Python 3.10 or newer is required.
    exit /b 1
)
for /f "delims=" %%v in ('%PY% -c "import sys;print(sys.version.split()[0])"') do set PYVER=%%v
echo       Python version: %PYVER%

echo [2/5] Preparing build environment ...
if exist ".buildenv" (
    echo       reusing existing .buildenv
) else (
    %PY% -m venv .buildenv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        exit /b 1
    )
)
set VENV=%CD%\.buildenv\Scripts
if not exist "%VENV%\python.exe" (
    echo [ERROR] Virtual environment is incomplete.
    exit /b 1
)

echo [3/5] Installing dependencies ...
"%VENV%\python.exe" -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo [ERROR] Failed to upgrade pip.
    exit /b 1
)
"%VENV%\python.exe" -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies. Check your network.
    exit /b 1
)

echo [4/5] Running test suite - build aborts if any test fails ...
"%VENV%\python.exe" -m pytest tests -q
if errorlevel 1 (
    echo.
    echo [ABORTED] Tests failed. Fix them before building.
    exit /b 1
)

echo [5/5] Building Windows executable ...
if exist "dist" rmdir /s /q "dist" >nul 2>&1
if exist "build" rmdir /s /q "build" >nul 2>&1
"%VENV%\python.exe" -m PyInstaller --noconfirm --clean attendance.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    exit /b 1
)

"%VENV%\python.exe" tools\make_release.py
if errorlevel 1 (
    echo [ERROR] Failed to assemble the release folder.
    exit /b 1
)

echo.
echo ============================================================
echo   Build finished. Output folder: 发布版
echo     - 三早一晚免勤自动统计.exe   double-click to run
echo     - 免勤词典.xlsx               editable rule table
echo     - SHA256.txt                 checksums
echo ============================================================
endlocal
'''


def main() -> int:
    path = os.path.join(ROOT, 'build_windows.bat')
    with open(path, 'w', encoding='gbk', newline='\r\n') as fh:
        fh.write(BAT.replace('\n', '\r\n'))
    print('已写出：%s（GBK 编码，CRLF 行尾）' % path)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
