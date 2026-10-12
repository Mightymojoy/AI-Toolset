@echo off
chcp 936 >nul
echo ================================
echo   ITO AI 工具集门户刷新
echo ================================
echo.
set PYTHON=C:\Users\QwQ\.workbuddy\binaries\python\envs\default\Scripts\python.exe
if not exist "%PYTHON%" set PYTHON=C:\Users\QwQ\.workbuddy\binaries\python\versions\3.13.12\python.exe
"%PYTHON%" "%~dp0refresh_portal.py"
pause
