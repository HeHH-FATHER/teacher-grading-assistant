@echo off
chcp 65001 >nul
title 教师作业批改助手

echo ========================================
echo   👨‍🏫 智能教师作业批改助手
echo ========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.10+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查 Streamlit
python -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    echo [提示] 正在安装依赖，请稍候...
    pip install streamlit openai python-docx PyPDF2 -q
    if %errorlevel% neq 0 (
        echo [错误] 依赖安装失败，请手动执行：
        echo pip install streamlit openai python-docx PyPDF2
        pause
        exit /b 1
    )
)

echo [启动] 正在启动服务器...
echo [提示] 浏览器将自动打开，关闭此窗口即可停止服务
echo.

:: 启动 Streamlit
start "" http://localhost:8501
python -m streamlit run app.py --server.headless true --server.port 8501

pause
