@echo off
chcp 936 >nul
echo ================================
echo   ITO AI 5DE5B7A5B7BF E99B86 E997A8E688B7E588B7E696B0
echo ================================
echo.

set PYTHON="C:\Users\yh200\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
set SCRIPT="E:\AI工具集\build_portal.py"

if not exist %PYTHON% (
    echo [ERROR] Python not found
    pause
    exit /b 1
)

if not exist %SCRIPT% (
    echo [ERROR] build_portal.py not found
    pause
    exit /b 1
)

echo [1/2] Reading Excel and generating portal...
%PYTHON% %SCRIPT%
set RET=%errorlevel%

if %RET% neq 0 (
    echo.
    echo [ERROR] Failed to generate portal
    pause
    exit /b %RET%
)

echo.
echo [2/2] Done!
echo.
echo Portal: E:\AI工具集\index.html
echo.
echo Press any key to open the portal...
pause

start "" "E:\AI工具集\index.html"