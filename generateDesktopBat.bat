@echo off
setlocal

REM 生成桌面启动脚本，使用本脚本所在目录的配置
REM 本脚本放在仓库根目录下，双击运行即可在桌面生成启动脚本

set "BASE_DIR=%~dp0"
for /f "delims=" %%I in ("%BASE_DIR%.") do set "BASE_DIR=%%~fI"
set "DESKTOP=%USERPROFILE%\Desktop"
set "LAUNCH_BAT=%DESKTOP%\DuoKai-启动.bat"

echo 正在生成桌面启动脚本: "%LAUNCH_BAT%"

>"%LAUNCH_BAT%" echo @echo off
>>"%LAUNCH_BAT%" echo setlocal
>>"%LAUNCH_BAT%" echo set "BASE_DIR=%BASE_DIR%"
>>"%LAUNCH_BAT%" echo cd /d "%%BASE_DIR%%"
>>"%LAUNCH_BAT%" echo rem 优先使用 py，其次使用 python
>>"%LAUNCH_BAT%" echo set "PY_CMD="
>>"%LAUNCH_BAT%" echo where py ^>nul 2^>^&1 ^&^& set "PY_CMD=py"
>>"%LAUNCH_BAT%" echo if "%%PY_CMD%%"=="" set "PY_CMD=python"
>>"%LAUNCH_BAT%" echo "%%PY_CMD%%" "%BASE_DIR%\window_manager_gui.py"
>>"%LAUNCH_BAT%" echo exit /b 0

echo ✅ 生成完成！请在桌面双击 "DuoKai-启动.bat" 启动程序。
exit /b 0