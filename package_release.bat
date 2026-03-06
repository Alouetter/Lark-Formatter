@echo off
setlocal
call "%~dp0scripts\windows\package_release.bat" %*
exit /b %errorlevel%
