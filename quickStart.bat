@echo off
chcp 65001>nul
setlocal
cd /d "D:\Project_PY\DuoKai"
set "PY_CMD="
where pyw >nul 2>&1 && set "PY_CMD=pyw"
if "%PY_CMD%"=="" where pythonw >nul 2>&1 && set "PY_CMD=pythonw"
if "%PY_CMD%"=="" where py >nul 2>&1 && set "PY_CMD=py"
if "%PY_CMD%"=="" set "PY_CMD=python"
"%PY_CMD%" "D:\Project_PY\DuoKai\window_manager_gui.py"