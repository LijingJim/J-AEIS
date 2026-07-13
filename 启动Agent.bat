@echo off
chcp 65001 >nul
title 公司智能助手

:: 检查是否已安装依赖
python -c "import streamlit" 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] 未检测到 streamlit，请先运行 安装依赖.bat
    pause
    exit /b 1
)

echo ========================================
echo   公司智能助手 v2.1
echo   浏览器访问: http://localhost:8501
echo   关闭此窗口即可停止服务
echo ========================================
echo.

streamlit run "%~dp0agent.py" --browser.gatherUsageStats false --server.port 8501
pause
