@echo off
cd /d %~dp0
title 财务报表生成工具
echo ============================================
echo   财务报表生成工具 v2.0
echo ============================================
echo.
echo 正在启动服务器...
echo 浏览器将自动打开，请勿关闭此窗口。
echo 使用完毕后，关闭此窗口即可停止服务。
echo.

REM 检查 Python 是否安装
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    pause
    exit /b 1
)

REM 检查 streamlit 是否安装
python -c "import streamlit" >nul 2>nul
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖，请稍候...
    pip install -r requirements.txt -q
)

REM 启动（streamlit 默认会自动打开浏览器）
start http://localhost:8501
streamlit run app.py --server.port 8501

pause
