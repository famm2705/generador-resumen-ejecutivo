@echo off
setlocal
cd /d "%~dp0"

if not exist "logs" mkdir "logs"

if not exist ".venv\Scripts\python.exe" (
  python -m venv .venv >> "logs\pmo_dashboard.log" 2>&1
  if errorlevel 1 exit /b 1
)

".venv\Scripts\python.exe" -m pip install -r requirements.txt >> "logs\pmo_dashboard.log" 2>&1
if errorlevel 1 exit /b 1

echo ==== %date% %time% ==== >> "logs\pmo_dashboard.log"
".venv\Scripts\python.exe" generar_resumen_ejecutivo.py --use-openai >> "logs\pmo_dashboard.log" 2>&1
exit /b %errorlevel%
