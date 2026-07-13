@echo off
chcp 65001 >nul
title 公司智能助手 - 安装依赖

echo ========================================
echo   公司智能助手 - 首次安装
echo ========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 未检测到 Python，请先安装 Python 3.10+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)
echo [OK] Python 已就绪
python --version
echo.

:: 升级 pip
echo [1/2] 升级 pip...
python -m pip install --upgrade pip -q

:: 安装依赖
echo [2/2] 安装项目依赖...
python -m pip install -r "%~dp0requirements.txt" -q
if %errorlevel% neq 0 (
    echo [ERROR] 依赖安装失败，请检查网络连接后重试
    pause
    exit /b 1
)
echo [OK] 依赖安装完成
echo.

:: 配置引导
echo ========================================
echo   配置 API Key
echo ========================================
echo.
echo 请编辑项目目录下的 config.py 文件，填入你的 API 信息：
echo   - API_KEY: 你的 API 密钥
echo   - BASE_URL: API 地址（默认 DeepSeek）
echo   - MODEL: 模型名称
echo.
echo 或者在启动后于侧边栏直接填写（推荐）。
echo.

echo [OK] 安装完成！双击 "启动Agent.bat" 即可启动。
echo.
echo 首次启动后，在侧边栏填写 API Key 即可使用。
pause
