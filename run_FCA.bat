@echo off
cd /d %~dp0
py -m pip install openpyxl
py build_fca_sheets.py
if errorlevel 1 (
  python -m pip install openpyxl
  python build_fca_sheets.py
)
echo.
pause
