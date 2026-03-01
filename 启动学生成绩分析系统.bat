@echo off
chcp 65001 >nul
title 学生成绩分析系统

echo ========================================
echo    学生成绩分析系统
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python，请先安装Python 3.8+
    pause
    exit /b 1
)

REM 检查Streamlit是否安装
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo 正在安装依赖...
    pip install streamlit pandas numpy plotly openpyxl
)

echo 正在启动应用...
echo 应用启动后，请在浏览器中访问: http://localhost:8501
echo.
echo 按 Ctrl+C 可停止应用
echo.

REM 启动Streamlit
python -m streamlit run student_analysis_app.py --server.port 8501

pause
