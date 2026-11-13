@echo off
chcp 65001 >nul
cd /d %~dp0
title 编译安装包 - 辅导员工作台

echo ============================================
echo   编译安装包（Inno Setup）
echo ============================================
echo.

rem ---- 1. 找 ISCC.exe ----
set "ISCC="
if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not defined ISCC if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
if not defined ISCC (
  where ISCC.exe >nul 2>nul
  if not errorlevel 1 set "ISCC=ISCC.exe"
)
if not defined ISCC (
  echo [没找到 Inno Setup 6]
  echo.
  echo 请先安装： https://jrsoftware.org/isdl.php
  echo 装的时候一路下一步即可，装完重新双击本文件。
  echo.
  pause
  exit /b 1
)
echo 使用编译器: %ISCC%
echo.

rem ---- 2. 找 Python（用来刷新 dist 与读取版本号）----
set "PY="
where py >nul 2>nul
if not errorlevel 1 set "PY=py"
if not defined PY (
  where python >nul 2>nul
  if not errorlevel 1 set "PY=python"
)
if not defined PY (
  echo [没找到 Python] 无法刷新 dist 目录。请先安装 Python 3.9+。
  pause
  exit /b 1
)

rem ---- 3. 版本号以 version.py 为准 ----
for /f "usebackq delims=" %%v in (`"%PY%" -c "import version;print(version.VERSION)"`) do set "VER=%%v"
if not defined VER (
  echo [读取版本号失败] 请检查 version.py 里的 VERSION。
  pause
  exit /b 1
)
echo 版本号: v%VER%
echo.

rem ---- 4. 先刷新绿色版目录（安装包的内容来源）----
echo [1/3] 生成发布目录 dist\...
"%PY%" ..\make_release.py
if errorlevel 1 goto fail

rem ---- 5. 图标资源（缺了才生成）----
if not exist "assets\app.ico" (
  echo [2/3] 生成图标与向导图...
  "%PY%" make_assets.py
) else (
  echo [2/3] 图标资源已存在，跳过（想重画就删掉 assets\app.ico 再跑）
)

rem ---- 6. 编译 ----
echo [3/3] 编译安装包...
"%ISCC%" /DAppVersion=%VER% installer.iss
if errorlevel 1 goto fail

rem ---- 7. 另存一份 ASCII 文件名，给 GitHub Release 用 ----
set "CN=..\dist\辅导员工作台-v%VER%-安装版.exe"
set "EN=..\dist\counselor-workbench-setup-v%VER%.exe"
if exist "%CN%" copy /y "%CN%" "%EN%" >nul && echo 已生成 ASCII 名副本: counselor-workbench-setup-v%VER%.exe

echo.
echo ============================================
echo   完成！产物在 ..\dist\ 下（两个文件内容相同）：
dir /b "..\dist\*v%VER%*.exe"
echo ============================================
pause
exit /b 0

:fail
echo.
echo [失败] 编译没有通过，看上面的报错。
pause
exit /b 1
