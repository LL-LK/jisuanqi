@echo off
chcp 936 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set PYTHONDONTWRITEBYTECODE=1

echo ============================================================
echo   Sanzao Wanwai Attendance Statistics - Windows Build
echo   三早一晚免勤统计系统 - Windows 打包脚本
echo ============================================================
echo.

echo [1/12] Checking Python 3.13 ...
where py >nul 2>nul
if errorlevel 1 goto :try_python
py -3.13 -c "import sys" >nul 2>nul
if not errorlevel 1 (
    set PY=py -3.13
    goto :python_found
)
:try_python
where python >nul 2>nul
if errorlevel 1 goto :no_python
set PY=python
goto :python_found
:no_python
echo [ERROR] Python not found. Please install Python 3.13 and tick "Add to PATH".
exit /b 1
:python_found
%PY% -c "import sys; sys.exit(0 if sys.version_info[:2] == (3,13) else 1)"
if errorlevel 1 goto :bad_python
for /f "delims=" %%v in ('%PY% -c "import sys;print(sys.version.split()[0])"') do set PYVER=%%v
echo       Python version: %PYVER%
goto :python_ok
:bad_python
echo [ERROR] Python 3.13.x is required by this project.
exit /b 1
:python_ok

echo [2/12] Preparing build environment ...
if exist ".buildenv\Scripts\python.exe" (
    echo       reusing existing .buildenv
) else (
    %PY% -m venv .buildenv
    if errorlevel 1 goto :venv_failed
)
set VENV=%CD%\.buildenv\Scripts
if not exist "%VENV%\python.exe" goto :venv_failed
goto :venv_ok
:venv_failed
echo [ERROR] Failed to create virtual environment.
exit /b 1
:venv_ok

echo [3/12] Installing dependencies ...
"%VENV%\python.exe" -m pip install --upgrade pip --quiet
if errorlevel 1 goto :deps_failed
"%VENV%\python.exe" -m pip install -r requirements.txt --quiet
if errorlevel 1 goto :deps_failed
if not exist "build" mkdir "build"
"%VENV%\python.exe" -m pip freeze > build\requirements.lock.txt
goto :deps_ok
:deps_failed
echo [ERROR] Failed to install dependencies. Check your network.
exit /b 1
:deps_ok

echo [4/12] Syntax check (fail fast before test collection) ...
"%VENV%\python.exe" -m py_compile app.py cli.py
if errorlevel 1 goto :compile_failed
for %%f in (core\*.py tools\*.py) do "%VENV%\python.exe" -m py_compile "%%f"
if errorlevel 1 goto :compile_failed
goto :compile_ok
:compile_failed
echo [ABORTED] Syntax error in source. Fix before building.
exit /b 1
:compile_ok

echo [5/12] Test collection check (duplicate module names break everything) ...
"%VENV%\python.exe" -m pytest tests --collect-only -q
if errorlevel 1 goto :collect_failed
goto :collect_ok
:collect_failed
echo [ABORTED] Test collection failed. Check duplicate test module names.
exit /b 1
:collect_ok

echo [6/12] Running test suite - build aborts if any test fails ...
"%VENV%\python.exe" -m pytest tests -q -rxX --junitxml=build\junit.xml
if errorlevel 1 goto :tests_failed
echo [6/12] Checking xfail list is zero (V3) ...
"%VENV%\python.exe" -c "import sys,xml.etree.ElementTree as ET;r=ET.parse(r'build\junit.xml').getroot();n=sum(1 for s in r.iter('skipped') if s.get('type')=='pytest.xfail');print('XFAIL_COUNT=' + str(n));sys.exit(1 if n else 0)"
if errorlevel 1 goto :xfail_not_zero
goto :tests_ok
:tests_failed
echo [ABORTED] Tests failed. No packaging.
exit /b 1
:xfail_not_zero
echo [ABORTED] xfail list is not zero (V3). Fix the behavior AND remove the marker in the same commit.
exit /b 1
:tests_ok

echo [7/12] Privacy scan on SOURCE tree ...
"%VENV%\python.exe" -X utf8 tools\privacy_scan.py --root "%CD%" --out build\privacy_scan_source.md
if errorlevel 1 goto :scan_source_failed
goto :scan_source_ok
:scan_source_failed
echo [ABORTED] Privacy scan failed on source tree. No packaging.
exit /b 1
:scan_source_ok

echo [8/12] Building Windows executable ...
if exist "dist" rmdir /s /q "dist" >nul 2>&1
if exist "build\AttendanceStats" rmdir /s /q "build\AttendanceStats" >nul 2>&1
"%VENV%\python.exe" -m PyInstaller --noconfirm --clean attendance.spec
if errorlevel 1 goto :pyi_failed
goto :pyi_ok
:pyi_failed
echo [ERROR] PyInstaller build failed.
exit /b 1
:pyi_ok

echo [9/12] Assembling release folder ...
"%VENV%\python.exe" tools\make_release.py
if errorlevel 1 goto :release_failed
goto :release_ok
:release_failed
echo [ERROR] Failed to assemble the release folder.
exit /b 1
:release_ok

echo [10/12] Packing Windows distribution package ...
"%VENV%\python.exe" tools\make_zip.py
if errorlevel 1 goto :zip_failed
goto :zip_ok
:zip_failed
echo [ERROR] Failed to build the distribution zip.
exit /b 1
:zip_ok

echo [11/12] Privacy scan on ARTIFACT (binary included) ...
"%VENV%\python.exe" -X utf8 tools\privacy_scan.py --root "%CD%" --artifact "发布版" --out build\privacy_scan_artifact.md
if errorlevel 1 goto :scan_artifact_failed
"%VENV%\python.exe" -X utf8 tools\privacy_scan.py --root "%CD%" --artifact "发布包" --out build\privacy_scan_package.md
if errorlevel 1 goto :scan_artifact_failed
goto :scan_artifact_ok
:scan_artifact_failed
echo [ABORTED] Privacy scan failed on artifact. Do not release.
exit /b 1
:scan_artifact_ok

echo [12/12] Verifying SHA-256 manifest ...
"%VENV%\python.exe" tools\verify_release.py "发布版" "发布包"
if errorlevel 1 goto :sha_failed
goto :sha_ok
:sha_failed
echo [ERROR] SHA-256 verification failed.
exit /b 1
:sha_ok

echo.
echo ============================================================
echo   Build finished. Output folder: 发布版
echo     - 三早一晚免勤自动统计.exe   double-click to run
echo     - 免勤词典.xlsx               editable rule table
echo     - SHA256.txt                 checksums
echo   Gates passed: syntax / collect / pytest / xfail-zero / privacy(source) / privacy(artifact) / sha256
echo   Release folder: 发布版   |   Distribution package: 发布包
echo ============================================================
endlocal
exit /b 0
