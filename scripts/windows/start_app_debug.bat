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

echo [INFO] Starting app in debug mode (console output enabled)...
".venv\Scripts\python.exe" main.py
exit /b %errorlevel%
