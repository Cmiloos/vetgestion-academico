@echo off
rem Genera VetGestion.exe: UN solo archivo, en esta misma carpeta. No necesita nada al lado.
rem Los datos se guardan en %LOCALAPPDATA%\VetGestion. Si existe app\config.json (con "sheets_url"),
rem queda incluido: el .exe arranca ya conectado a Google Sheets.
cd /d "%~dp0"
python -m pip install --quiet pywebview openpyxl pyinstaller
set EXTRA=
if exist app\config.json set EXTRA=--add-data "app\config.json;."
python -m PyInstaller --noconfirm --clean --onefile --windowed ^
  --name VetGestion ^
  --icon app\icono.ico ^
  --add-data "app\ui;ui" ^
  %EXTRA% ^
  --exclude-module PIL --exclude-module tkinter --exclude-module numpy ^
  --distpath . --workpath build ^
  --paths app ^
  app\main.py
if errorlevel 1 (
  echo.
  echo ERROR: no se pudo generar el ejecutable.
  pause
  exit /b 1
)
rmdir /s /q build 2>nul
del /q VetGestion.spec 2>nul
echo.
echo Listo: VetGestion.exe
