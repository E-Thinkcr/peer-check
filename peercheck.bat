@echo off
REM Doble clic: revisa los Excel de la carpeta INPUT que está junto a este archivo.
cd /d "%~dp0"
python -m peercheck --carpeta "%~dp0."
echo.
pause
