@echo off
setlocal
cd /d "%~dp0"

if "%OPENAI_API_KEY%"=="" (
  echo OPENAI_API_KEY no esta configurada en Windows.
  echo.
  echo Configurala en PowerShell con:
  echo $env:OPENAI_API_KEY="sk-..."
  echo python generar_resumen_ejecutivo.py --use-openai
  echo.
  echo O guardala de forma persistente con:
  echo setx OPENAI_API_KEY "sk-..."
  echo.
  pause
  exit /b 1
)

python -m pip install -r requirements.txt
if errorlevel 1 (
  echo No se pudieron instalar las dependencias.
  pause
  exit /b 1
)

python generar_resumen_ejecutivo.py --use-openai %*
if errorlevel 1 (
  echo Error generando el dashboard PMO con OpenAI.
  pause
  exit /b 1
)

echo.
echo Dashboard PMO generado correctamente con narrativa OpenAI.
pause
