@echo off
chcp 65001 >nul
cd /d %~dp0
title 辅导员工作台 - 编辑模式（关闭窗口即退出；本窗口仅显示日志）

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
echo   会弹出独立窗口；若没弹出，可手动在浏览器打开上面的地址
echo   关闭本窗口即停止服务（仅本机访问，不联网）
echo ============================================
echo.

"%PY%" launcher.py

echo.
pause
