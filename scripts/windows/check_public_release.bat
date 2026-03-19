@echo off
setlocal
set "REPO_ROOT=%~dp0..\.."
for %%I in ("%REPO_ROOT%") do set "REPO_ROOT=%%~fI"
cd /d "%REPO_ROOT%"

if exist ".venv\Scripts\python.exe" (
    set "PYTHON=.venv\Scripts\python.exe"
) else (
    set "PYTHON=python"
)

%PYTHON% -c "import sys" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python or create .venv first.
    exit /b 1
)

%PYTHON% scripts\check_public_release.py %*
exit /b %errorlevel%
