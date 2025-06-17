@echo off
setlocal EnableDelayedExpansion

:: 自动请求管理员权限
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo 正在请求管理员权限...
    powershell -Command "Start-Process cmd -ArgumentList '/c cd /d \"%~dp0\" && call \"%~f0\"' -Verb RunAs"
    exit /b
)

:: 确保工作目录正确
cd /d "%~dp0"

chcp 65001 >nul
title Windows 端口管理工具

echo ====================================
echo   Windows 端口管理工具 - 一键运行
echo ====================================
echo ✅ 管理员权限已获取

:: 检查并安装Python依赖
python --version >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ 未安装Python，请先安装Python 3.7+
    pause & exit /b 1
)

python -c "import customtkinter" >nul 2>&1
if %errorLevel% neq 0 (
    echo 📦 正在安装依赖包...
    python -m pip install -q customtkinter pillow
)

echo 🚀 启动程序...
echo 当前目录: %CD%
python main.py

echo 程序已退出
pause 