@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal

set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY set "PY=C:/Users/sjl/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" run.py %*
pause
