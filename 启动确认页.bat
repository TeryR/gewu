@echo off
rem Excel confirm page launcher: drag an .xlsx file onto this icon
cd /d "%~dp0"
if "%~1"=="" (
  echo Usage: drag an .xlsx file onto this .bat icon, then release.
  pause
  exit /b 1
)
if not exist "%~1" (
  echo File not found: %~1
  pause
  exit /b 1
)
where python >nul 2>nul
if errorlevel 1 (
  echo [ERROR] python not found in PATH.
  pause
  exit /b 1
)
echo Starting confirm page for: %~1
echo Browser will open at http://127.0.0.1:8765  (Ctrl+C to quit)
python serve.py "%~1" 8765
echo.
echo Confirm page exited.
pause
