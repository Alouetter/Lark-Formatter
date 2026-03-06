@echo off
setlocal
call "%~dp0scripts\windows\start_app.bat" %*
exit /b %errorlevel%
