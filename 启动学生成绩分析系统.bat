@echo off
chcp 65001 >nul
title Student Score Analysis System

echo ========================================
echo    Student Score Analysis System
echo ========================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python not found. Please install Python 3.8+
    pause
    exit /b 1
)

python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    pip install streamlit pandas numpy plotly openpyxl xlrd
)

echo Starting application...
echo Please visit: http://localhost:8501
echo.
echo Press Ctrl+C to stop
echo.

python -m streamlit run student_analysis_app.py --server.port 8501

pause
