@echo off
chcp 65001 >nul
echo ================================
echo   ITO · AI 工具集 门户刷新工具
echo ================================
echo.

set PYTHON=C:\Users\yh200\.workbuddy\binaries\python\envs\default\Scripts\python.exe
set SCRIPT=E:\AI工具集\build_portal.py

if not exist "%PYTHON%" (
    echo [错误] 找不到 Python 环境：%PYTHON%
    pause
    exit /b 1
)

echo [1/2] 正在从 Excel 解析工具清单...
"%PYTHON%" "%SCRIPT%"

if %errorlevel% neq 0 (
    echo [错误] 生成门户失败，请检查 Excel 文件格式。
    pause
    exit /b 1
)

echo.
echo [2/2] 完成！双击 index.html 即可查看门户。
echo       路径：E:\AI工具集\index.html

echo.
echo 提示：直接按回车打开门户网页...
pause >nul
start "" "E:\AI工具集\index.html"
