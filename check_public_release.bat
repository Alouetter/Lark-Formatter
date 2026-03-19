@echo off
setlocal
call "%~dp0scripts\windows\check_public_release.bat" %*
exit /b %errorlevel%
