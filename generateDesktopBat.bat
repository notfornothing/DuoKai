@echo off
setlocal

REM 纯 CMD 生成桌面启动脚本（带黑框输出）
set "BASE_DIR=%~dp0"
for /f "delims=" %%I in ("%BASE_DIR%.") do set "BASE_DIR=%%~fI"
set "DESKTOP=%USERPROFILE%\Desktop"
set "LAUNCH_BAT=%DESKTOP%\DuoKai-启动(黑框).bat"

echo 正在生成桌面启动脚本: "%LAUNCH_BAT%"

REM 用一次性重定向写入所有行，避免中文与特殊符号解析问题
(
  echo @echo off
  echo chcp 65001^>nul
  echo setlocal
  echo cd /d "%BASE_DIR%"
  echo set "PY_CMD="
  echo where py ^>nul 2^>^&1 ^&^& set "PY_CMD=py"
  echo if "%%PY_CMD%%"=="" set "PY_CMD=python"
  echo "%%PY_CMD%%" -u "%BASE_DIR%\window_manager_gui.py"
  echo echo.
  echo pause
) >"%LAUNCH_BAT%"

if exist "%LAUNCH_BAT%" (
  echo 生成完成。请在桌面双击 "DuoKai-启动(黑框).bat" 运行并查看输出。
) else (
  echo 生成失败，请检查权限或路径是否存在。
)
exit /b 0