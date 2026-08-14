@echo off
chcp 936 >nul
echo ================================
echo   ITO AI 工具集门户刷新
echo ================================
echo.

set PYTHON=C:\Users\QwQ\.workbuddy\binaries\python\versions\3.13.12\python.exe
set SCRIPT=D:\E盘文件\AI工具集\build_portal.py
set REPO=D:\E盘文件\AI工具集
set GIT=C:\Users\QwQ\.workbuddy\binaries\PortableGit\versions\1.2.0\mingw64\bin\git.exe

if not exist "%PYTHON%" (
    echo [ERROR] 未找到 Python：%PYTHON%
    echo 请确认 WorkBuddy 管理的 Python 环境存在，或把 PYTHON 改成可用的 python.exe 路径。
    pause
    exit /b 1
)

if not exist "%SCRIPT%" (
    echo [ERROR] 未找到 build_portal.py：%SCRIPT%
    echo 请确认本 bat 与仓库位于 D:\E盘文件\AI工具集\ 下。
    pause
    exit /b 1
)

echo [0/4] 同步最新仓库（git pull --ff-only）...
cd /d "%REPO%"
set GIT_TERMINAL_PROMPT=0
"%GIT%" pull --ff-only
if %errorlevel% neq 0 (
    echo [WARN] 拉取失败（可能离线或本地有冲突改动），将使用本地现有脚本继续。
) else (
    echo [OK] 已同步到最新脚本与模板。
)

echo [1/4] 从 Excel 生成门户...
"%PYTHON%" "%SCRIPT%"
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 生成门户失败（多半是缺少 openpyxl，或 xlsx 结构变动）。
    pause
    exit /b %errorlevel%
)

echo.
echo [2/4] 提交到 Git...
"%GIT%" add -A
"%GIT%" commit -m "auto: update toolset portal"
if %errorlevel% neq 0 (
    if %errorlevel% equ 1 (
        echo [INFO] 没有变化需要提交，跳过推送。
    ) else (
        echo [ERROR] Git 提交失败。
        pause
        exit /b %errorlevel%
    )
) else (
    echo [3/4] 推送到 GitHub...
    "%GIT%" push origin main
    if %errorlevel% neq 0 (
        echo [ERROR] Git 推送失败（检查网络 / 凭证）。
        pause
        exit /b %errorlevel%
    )
    echo.
    echo GitHub 同步完成！GitHub Pages 约 1-2 分钟后重建，请硬刷新（Ctrl+Shift+R）查看。
)

echo.
echo 门户文件：D:\E盘文件\AI工具集\index.html
echo.
echo 按任意键打开门户...
pause

start "" "D:\E盘文件\AI工具集\index.html"
