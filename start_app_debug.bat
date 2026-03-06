@echo off
setlocal
call "%~dp0scripts\windows\start_app_debug.bat" %*
exit /b %errorlevel%
