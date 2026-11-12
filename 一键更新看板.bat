@echo off
chcp 65001 >nul
cd /d %~dp0
title 辅导员工作台 - 更新看板

set "PY=C:\Users\user\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=py"

"%PY%" -c "import sys" >nul 2>nul
if errorlevel 1 (
  echo [错误] 没找到可用的 Python。请告诉工作台助手：本机 python 路径需要重新指定。
  pause
  exit /b 1
)

echo 正在用 data 目录下的 CSV 重建看板并同步数据库...
"%PY%" db_sync.py csv2db
"%PY%" build.py
echo.
echo 完成：dashboard.html / 脱敏版 / 汇总xlsx / 辅导员工作台.db 已更新
echo 按任意键关闭。
pause >nul
