@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem ---- pick python: bundled .venv > workbuddy venv > system python ----
set "PY="
if exist "%~dp0.venv\Scripts\python.exe" set "PY=%~dp0.venv\Scripts\python.exe"
if not defined PY set "PY=C:/Users/sjl/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
if not exist "%PY%" set "PY=python"

rem ---- all Chinese UI text lives in batch_run.py (keeps this file ASCII-safe) ----
"%PY%" batch_run.py %*
exit /b %ERRORLEVEL%
