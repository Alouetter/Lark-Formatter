@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "REPO_ROOT=%~dp0..\.."
for %%I in ("%REPO_ROOT%") do set "REPO_ROOT=%%~fI"
cd /d "%REPO_ROOT%"

set "BASE_PY_CMD="
set "BASE_PY_LABEL="

for %%V in (3.14 3.13 3.12 3.11 3.10) do (
    py -%%V --version >nul 2>&1
    if !errorlevel! == 0 (
        set "BASE_PY_CMD=py -%%V"
        set "BASE_PY_LABEL=Python Launcher (%%V)"
        goto :found_base
    )
)

python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if %errorlevel%==0 (
    set "BASE_PY_CMD=python"
    set "BASE_PY_LABEL=python"
    goto :found_base
)

for %%P in (
    "%ProgramFiles%\QGIS 3.42.1\apps\Python312\python.exe"
    "%ProgramFiles%\QGIS 3.42.1\bin\python.exe"
    "%ProgramFiles%\Blender Foundation\Blender 5.0\5.0\python\bin\python.exe"
    "%ProgramFiles%\Blender Foundation\Blender 4.4\4.4\python\bin\python.exe"
    "%ProgramFiles%\Blender Foundation\Blender 3.6\3.6\python\bin\python.exe"
) do (
    if exist %%~P (
        "%%~P" -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
        if !errorlevel! == 0 (
            set "BASE_PY_CMD=""%%~P"""
            set "BASE_PY_LABEL=Fallback embedded Python: %%~P"
            goto :found_base
        )
    )
)

:found_base
if not defined BASE_PY_CMD (
    echo [ERROR] No usable Python interpreter found.
    echo Install Python 3.10+ and run this script again.
    pause
    exit /b 1
)

echo [1/4] Base interpreter: %BASE_PY_LABEL%
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -m pip --version >nul 2>&1
    if errorlevel 1 (
        echo [WARN] Existing .venv is missing pip, recreating virtual environment...
        rmdir /s /q ".venv"
    )
)
if not exist ".venv\Scripts\python.exe" (
    %BASE_PY_CMD% -m venv ".venv"
    if errorlevel 1 (
        echo [ERROR] Failed to create .venv
        pause
        exit /b 1
    )
)

echo [2/4] Ensure pip is available
".venv\Scripts\python.exe" -m pip --version >nul 2>&1
if errorlevel 1 (
    ".venv\Scripts\python.exe" -m ensurepip --upgrade
    if errorlevel 1 (
        echo [ERROR] Failed to bootstrap pip in .venv
        pause
        exit /b 1
    )
)

echo [3/4] Upgrade pip
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 (
    echo [ERROR] Failed to upgrade pip
    pause
    exit /b 1
)

echo [4/4] Install dependencies
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)

echo [OK] Environment is ready. Run "start_app.bat" to start the app.
exit /b 0
