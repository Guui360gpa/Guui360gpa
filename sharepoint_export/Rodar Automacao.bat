@echo off
REM Executa a automacao SharePoint -> Excel
cd /d "%~dp0"
python sharepoint_to_excel.py
echo.
pause
