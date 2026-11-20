@echo off
REM Doble clic: revisa los Excel de la carpeta INPUT que esta junto a este archivo.
REM Si copias este .bat a tu carpeta de trabajo, cambia PROGRAMA por la carpeta
REM donde descargaste Peer Check (la que contiene la carpeta "peercheck").
set "PROGRAMA=%~dp0"
set "PYTHONPATH=%PROGRAMA%"
python -m peercheck --carpeta "%~dp0."
echo.
pause
