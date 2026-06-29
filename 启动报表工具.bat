@echo off
cd /d %~dp0
echo ============================================
echo   财务报表生成工具
echo   Financial Report Generator
echo ============================================
echo.
echo 正在启动，浏览器将自动打开...
echo 请勿关闭此窗口。
echo.
streamlit run app.py --server.port 8501
pause
