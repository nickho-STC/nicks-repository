@echo off
cd /d "%~dp0"
set PY=py
where py >nul 2>&1 || set PY=python
echo Installing the tools this script needs (one time only)...
%PY% -m pip install --upgrade pip
%PY% -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo *** Something went wrong. Is Python installed? See README.txt step 1. ***
) else (
  echo.
  echo Done! Now double-click 2_TEST_10_PATIENTS.bat
)
pause
