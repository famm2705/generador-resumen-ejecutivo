@echo off
setlocal
cd /d "%~dp0"

python -m pip install -r requirements.txt
if errorlevel 1 (
  echo No se pudieron instalar las dependencias.
  pause
  exit /b 1
)

python generar_resumen_ejecutivo.py %*
if errorlevel 1 (
  echo Error generando el dashboard PMO.
  pause
  exit /b 1
)

echo.
echo Dashboard PMO generado correctamente.
pause
