@echo off
cd /d "%~dp0"
set PY=py
where py >nul 2>&1 || set PY=python
%PY% fill_ccta_intake.py --all
echo.
start "" "%~dp0output"
pause
