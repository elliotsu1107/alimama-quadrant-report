@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Python not found. Please install Python 3.10+ from python.org first.
  pause
  exit /b 1
)

echo Creating virtual env...
python -m venv .venv
if errorlevel 1 (
  echo [ERROR] Failed to create venv.
  pause
  exit /b 1
)

echo.
echo Installing pandas / openpyxl (about 60MB, may take a few minutes)...
echo You will see download progress below. Do not close this window.
echo.
call ".venv\Scripts\python.exe" -m pip install pandas openpyxl
if errorlevel 1 (
  echo.
  echo Install failed. Retrying with China mirror...
  call ".venv\Scripts\python.exe" -m pip install pandas openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
  if errorlevel 1 (
    echo.
    echo [ERROR] Install failed again. Check your network and run setup.bat again.
    pause
    exit /b 1
  )
)

echo.
echo ============================================================
echo   Setup complete! Now double-click batch.bat or single.bat
echo ============================================================
pause
