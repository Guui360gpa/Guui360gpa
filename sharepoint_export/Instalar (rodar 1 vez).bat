@echo off
REM Instala as dependencias (rodar uma unica vez)
cd /d "%~dp0"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m playwright install msedge
echo.
echo Instalacao concluida.
pause
