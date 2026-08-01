@echo off
REM ============================================================
REM  Owner-only build script. NOT for employees.
REM  Produces:
REM    dist\SheinExtractAU.exe              (PyInstaller output)
REM    dist\SheinExtractAU-Setup-X.Y.Z.exe  (Inno Setup output)
REM
REM Prerequisites (one-time):
REM   1. Python 3.10+ on PATH (Anaconda fine)
REM   2. pip install pyinstaller
REM   3. pip install -r requirements.txt
REM   4. Inno Setup 6.x installed (so iscc.exe is on PATH or at
REM      C:\Program Files (x86)\Inno Setup 6\)
REM
REM Usage:
REM   build.bat        full build + installer
REM   build.bat exe    only PyInstaller, skip Inno Setup
REM ============================================================

setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

REM Anaconda ships tk/tcl/openssl DLLs in Library\bin. PyInstaller walks
REM dependencies via PATH — without this, DLLs aren't found and the EXE
REM crashes at runtime (e.g. ImportError: _tkinter).
for /f "delims=" %%P in ('python -c "import sys,os; print(os.path.join(sys.prefix,'Library','bin'))" 2^>nul') do set "CONDA_LIBBIN=%%P"
if defined CONDA_LIBBIN if exist "%CONDA_LIBBIN%" (
    echo [build] Prepending conda Library\bin to PATH: %CONDA_LIBBIN%
    set "PATH=%CONDA_LIBBIN%;%PATH%"
)

echo.
echo ============================================================
echo  Step 1/3: Generate key_store.py from .build_key.txt
echo ============================================================
if not exist .build_key.txt (
    echo [ERROR] .build_key.txt not found. Create it with the raw API key.
    exit /b 1
)
python make_key_store.py
if errorlevel 1 (
    echo [ERROR] make_key_store.py failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 2/3: PyInstaller - build SheinExtractAU.exe
echo ============================================================
if exist build rmdir /s /q build
if exist dist\SheinExtractAU.exe del /q dist\SheinExtractAU.exe
python -m PyInstaller pyinstaller.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] PyInstaller failed.
    exit /b 1
)

if /i "%1"=="exe" (
    echo.
    echo Skipping Inno Setup (--exe-only). dist\SheinExtractAU.exe is ready.
    exit /b 0
)

echo.
echo ============================================================
echo  Step 3/3: Inno Setup - wrap into installer
echo ============================================================
where iscc >nul 2>nul
if errorlevel 1 (
    if exist "C:\Program Files (x86)\Inno Setup 6\iscc.exe" (
        set "ISCC=C:\Program Files (x86)\Inno Setup 6\iscc.exe"
    ) else (
        echo [ERROR] iscc.exe not found. Install Inno Setup 6 from
        echo https://jrsoftware.org/isdl.php
        exit /b 1
    )
) else (
    set "ISCC=iscc"
)
"%ISCC%" installer.iss
if errorlevel 1 (
    echo [ERROR] Inno Setup compilation failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Build complete!
echo ============================================================
echo Installer: dist\SheinExtractAU-Setup-*.exe
echo Next: test locally, then `release.bat X.Y.Z` to publish.
echo ============================================================
