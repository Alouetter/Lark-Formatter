@echo off
setlocal
set "REPO_ROOT=%~dp0..\.."
for %%I in ("%REPO_ROOT%") do set "REPO_ROOT=%%~fI"
cd /d "%REPO_ROOT%"

if exist ".venv\Scripts\pythonw.exe" (
    start "" /B ".venv\Scripts\pythonw.exe" main.py
    exit /b 0
)

echo [ERROR] Missing virtual environment: .venv
echo Run "install_env.bat" first.
pause
exit /b 1
