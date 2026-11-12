@echo off
chcp 65001 >nul
cd /d %~dp0
title 辅导员工作台 - 更新看板

rem 依次尝试 py 启动器 / PATH 里的 python（不写死本机路径，换机器也能用）
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py"
if not defined PY (
  where python >nul 2>nul
  if not errorlevel 1 set "PY=python"
)

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
pause
