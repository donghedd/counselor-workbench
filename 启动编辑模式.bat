@echo off
chcp 65001 >nul
cd /d %~dp0
title 辅导员工作台 - 编辑模式（关闭本窗口即停止服务）

set "PY=C:\Users\user\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=py"

"%PY%" -c "import sys" >nul 2>nul
if errorlevel 1 (
  echo.
  echo [错误] 没找到可用的 Python，无法启动编辑服务。
  echo 请把这句话告诉工作台助手：本机 python 路径需要重新指定。
  echo.
  pause
  exit /b 1
)

echo ============================================
echo   辅导员工作台 - 编辑模式
echo ============================================
echo   服务地址：http://127.0.0.1:8787/dashboard.html
echo   浏览器会自动打开；若没打开，请手动复制上面的地址
echo   关闭本窗口即停止服务（仅本机访问，不联网）
echo ============================================
echo.

"%PY%" server.py

echo.
echo 服务已停止。按任意键关闭本窗口。
pause >nul
