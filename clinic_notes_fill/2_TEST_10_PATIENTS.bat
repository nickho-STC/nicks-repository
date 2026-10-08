@echo off
cd /d "%~dp0"
set PY=py
where py >nul 2>&1 || set PY=python
%PY% fill_clinical_notes.py
echo.
start "" "%~dp0output"
pause
