@echo off
setlocal
set "REPO_ROOT=%~dp0..\.."
for %%I in ("%REPO_ROOT%") do set "REPO_ROOT=%%~fI"
cd /d "%REPO_ROOT%"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] Missing virtual environment: .venv
    echo Run "install_env.bat" first.
    pause
    exit /b 1
)

set "SPEC_FILE=Lark-Formatter_v0.20_LTS.spec"
if not exist "%SPEC_FILE%" (
    echo [ERROR] Spec file not found: %SPEC_FILE%
    pause
    exit /b 1
)

tasklist /FI "IMAGENAME eq Lark-Formatter.exe" 2>nul | find /I "Lark-Formatter.exe" >nul
if not errorlevel 1 (
    echo [ERROR] Lark-Formatter.exe is running.
    echo Please close the app before packaging, then run this script again.
    pause
    exit /b 1
)

echo [1/3] Ensure PyInstaller is installed
".venv\Scripts\python.exe" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found, installing...
    ".venv\Scripts\python.exe" -m pip install pyinstaller
    if errorlevel 1 (
        echo [ERROR] Failed to install PyInstaller
        pause
        exit /b 1
    )
) else (
    echo PyInstaller already installed.
)

echo [2/3] Build release package from %SPEC_FILE%
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean "%SPEC_FILE%"
if errorlevel 1 (
    echo [ERROR] Packaging failed
    pause
    exit /b 1
)

echo [3/3] Place templates folder next to EXE
set "DIST_DIR=dist\Lark-Formatter_v0.20_LTS"
if not exist "%DIST_DIR%" (
    echo [ERROR] Dist folder not found: %DIST_DIR%
    pause
    exit /b 1
)
if exist "%DIST_DIR%\templates" rmdir /s /q "%DIST_DIR%\templates"
xcopy /E /I /Y "src\scene\presets" "%DIST_DIR%\templates" >nul
if errorlevel 1 (
    echo [ERROR] Failed to copy templates folder to dist.
    pause
    exit /b 1
)

echo [OK] Build completed.
echo Output folder: dist\Lark-Formatter_v0.20_LTS
exit /b 0
