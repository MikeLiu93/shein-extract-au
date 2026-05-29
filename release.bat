@echo off
REM ============================================================
REM  Semi-automated release.
REM
REM Usage:
REM   release.bat 0.1.1
REM
REM Does:
REM   1. Verify clean working tree + on main + in sync with origin
REM   2. Update version.py to the given version
REM   3. Commit + push the version bump
REM   4. Run build.bat (PyInstaller + Inno Setup)
REM   5. Verify dist\SheinExtractAU-Setup-X.Y.Z.exe exists
REM   6. git tag v{version} + git push origin v{version}
REM   7. Open the GitHub "Draft a new release" page in your browser
REM      with the tag pre-selected. You manually drag the setup
REM      .exe into Assets, write release notes, click Publish.
REM ============================================================

setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

if "%1"=="" (
    echo Usage: release.bat X.Y.Z
    exit /b 1
)
set VERSION=%1

echo.
echo ============================================================
echo  Step 1/7: Pre-flight checks
echo ============================================================

git rev-parse --abbrev-ref HEAD > %TEMP%\branch.txt
set /p BRANCH=<%TEMP%\branch.txt
del %TEMP%\branch.txt
if not "%BRANCH%"=="main" (
    echo [ERROR] Not on main branch ^(on %BRANCH%^). Aborting.
    exit /b 1
)

git status --porcelain > %TEMP%\status.txt
for %%A in (%TEMP%\status.txt) do if not %%~zA==0 (
    echo [ERROR] Working tree not clean. Commit or stash first.
    del %TEMP%\status.txt
    exit /b 1
)
del %TEMP%\status.txt

git fetch origin main >nul 2>&1
git rev-list HEAD..origin/main --count > %TEMP%\behind.txt
set /p BEHIND=<%TEMP%\behind.txt
del %TEMP%\behind.txt
if not "%BEHIND%"=="0" (
    echo [ERROR] Local main is %BEHIND% commit^(s^) behind origin/main. Pull first.
    exit /b 1
)

echo OK: clean main, in sync.

echo.
echo ============================================================
echo  Step 2/7: Bump version.py to %VERSION%
echo ============================================================

python -c "import pathlib,re; p=pathlib.Path('version.py'); s=p.read_text(encoding='utf-8'); s2=re.sub(r'VERSION\s*=\s*\".*?\"', 'VERSION = \"%VERSION%\"', s); p.write_text(s2, encoding='utf-8'); print(s2.strip())"
if errorlevel 1 (
    echo [ERROR] Failed to bump version.py
    exit /b 1
)

echo.
echo ============================================================
echo  Step 3/7: Bump installer.iss MyAppVersion to %VERSION%
echo ============================================================

python -c "import pathlib,re; p=pathlib.Path('installer.iss'); s=p.read_text(encoding='utf-8'); s2=re.sub(r'#define MyAppVersion \"[^\"]*\"', '#define MyAppVersion \"%VERSION%\"', s); p.write_text(s2, encoding='utf-8'); print('installer.iss bumped')"
if errorlevel 1 (
    echo [ERROR] Failed to bump installer.iss
    exit /b 1
)

echo.
echo ============================================================
echo  Step 4/7: Commit + push version bump
echo ============================================================

git add version.py installer.iss
git commit -m "release: v%VERSION%"
git push origin main
if errorlevel 1 (
    echo [ERROR] git push failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 5/7: Build (PyInstaller + Inno Setup)
echo ============================================================

call build.bat
if errorlevel 1 (
    echo [ERROR] build.bat failed.
    exit /b 1
)

if not exist "dist\SheinExtractAU-Setup-%VERSION%.exe" (
    echo [ERROR] Expected dist\SheinExtractAU-Setup-%VERSION%.exe but it's missing.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 6/7: Tag + push tag
echo ============================================================

git tag v%VERSION%
git push origin v%VERSION%
if errorlevel 1 (
    echo [ERROR] tag push failed.
    exit /b 1
)

echo.
echo ============================================================
echo  Step 7/7: Open GitHub "Draft a new release" page
echo ============================================================
echo.
echo Drag this file into the Assets section, write release notes,
echo then click "Publish release":
echo.
echo   dist\SheinExtractAU-Setup-%VERSION%.exe
echo.

start "" "https://github.com/MikeLiu93/shein-extract-au/releases/new?tag=v%VERSION%"

echo Done. Browser opened. Finish the release in the browser.
