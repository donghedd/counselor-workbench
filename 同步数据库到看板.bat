@echo off
chcp 65001 >nul
cd /d %~dp0
title 辅导员工作台 - 同步数据库到看板

set "PY=C:\Users\user\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=py"

"%PY%" -c "import sys" >nul 2>nul
if errorlevel 1 (
  echo [错误] 没找到可用的 Python。请告诉工作台助手：本机 python 路径需要重新指定。
  pause
  exit /b 1
)

echo 正在把 辅导员工作台.db 的内容导出为 CSV 并重建看板...
"%PY%" db_sync.py db2csv
"%PY%" build.py
echo.
echo 完成：数据库改动已同步到看板
echo 按任意键关闭。
pause >nul
