@echo off
chcp 936 >nul
echo ================================
echo   ITO AI 5DE5B7A5B7BF E99B86 E997A8E688B7E588B7E696B0
echo ================================
echo.

set PYTHON=C:\Users\yh200\.workbuddy\binaries\python\envs\default\Scripts\python.exe
set SCRIPT=E:\AI工具集\build_portal.py
set REPO=E:\AI工具集

if not exist "%PYTHON%" (
    echo [ERROR] Python not found
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo [ERROR] build_portal.py not found
    pause
    exit /b 1
)

echo [1/3] Generating portal from Excel...
"%PYTHON%" "%SCRIPT%"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to generate portal
    pause
    exit /b %errorlevel%
)

echo.
echo [2/3] Committing to Git...
cd /d "%REPO%"
git add -A
git commit -m "auto: update toolset portal"
if %errorlevel% neq 0 (
    if %errorlevel% equ 1 (
        echo [INFO] Nothing to commit, skipping push...
    ) else (
        echo [ERROR] Git commit failed
        pause
        exit /b %errorlevel%
    )
) else (
    echo [3/3] Pushing to GitHub...
    git push
    if %errorlevel% neq 0 (
        echo [ERROR] Git push failed
        pause
        exit /b %errorlevel%
    )
    echo.
    echo GitHub sync complete!
)

echo.
echo Portal: E:\AI工具集\index.html
echo.
echo Press any key to open the portal...
pause

start "" "E:\AI工具集\index.html"
