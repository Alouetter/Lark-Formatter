@echo off
setlocal
call "%~dp0scripts\windows\install_env.bat" %*
exit /b %errorlevel%
