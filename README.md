# Generador de Resumen Ejecutivo PMO

Automatiza la lectura del Excel EVM exportado desde MS Project y genera una hoja `Resumen Ejecutivo PMO` con formato tipo dashboard.

El dashboard no congela filas al hacer scroll y genera 10 graficos como imagenes PNG embebidas para que carguen de forma estable al abrir el Excel.

## Uso rapido

1. Colocar el Excel `.xlsx` y `Prompt.txt` en este folder.
2. Ejecutar `Generar_Dashboard_PMO.bat`.
3. El resultado se guarda en `outputs\*_Resumen_PMO.xlsx`.

Tambien se puede ejecutar desde terminal:

```powershell
python -m pip install -r requirements.txt
python generar_resumen_ejecutivo.py
```

## Uso con OpenAI API opcional

Los calculos EVM y el formato del dashboard son locales. La API solo se usa para mejorar la narrativa ejecutiva.

Opcion recomendada para este folder: guardar la clave en `.venv\.env`.

```powershell
python -m venv .venv
copy env.example .venv\.env
notepad .venv\.env
```

Dentro de `.venv\.env`:

```text
OPENAI_API_KEY=sk-tu_api_key_aqui
OPENAI_MODEL=gpt-5.5
```

Luego ejecutar:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe generar_resumen_ejecutivo.py --use-openai
```

Tambien se soporta `PMO_ENV_FILE` para apuntar a un `.env` fuera de OneDrive:

```powershell
$env:PMO_ENV_FILE="C:\PMO_Config\evm_dashboard.env"
.\.venv\Scripts\python.exe generar_resumen_ejecutivo.py --use-openai
```

Alternativa de sesion temporal:

```powershell
$env:OPENAI_API_KEY="sk-..."
python generar_resumen_ejecutivo.py --use-openai
```

Tambien puedes usar `Generar_Dashboard_PMO_OpenAI.bat` despues de configurar la variable `OPENAI_API_KEY` en Windows.

Opcionalmente se puede cambiar el modelo:

```powershell
$env:OPENAI_MODEL="gpt-5.5"
python generar_resumen_ejecutivo.py --use-openai
```

Si no existe `OPENAI_API_KEY`, el script usa la narrativa local y deja una nota en la hoja.

## Ejecucion programada en OneDrive

Estructura recomendada del folder sincronizado:

```text
Generador Resumen Ejecutivo\
  generar_resumen_ejecutivo.py
  requirements.txt
  Prompt.txt
  EVM_Simplificado.xlsx
  Ejecutar_Programado_PMO_OpenAI.bat
  .venv\
    .env
  logs\
```

Requisitos:

- El folder debe estar sincronizado localmente por OneDrive, no solo en la nube.
- Marcar el folder como `Always keep on this device` / `Mantener siempre en este dispositivo`.
- El Excel debe estar cerrado al momento de ejecutar para poder sobrescribir la salida. Si esta abierto, el script crea una copia con timestamp.
- `Prompt.txt` debe permanecer en el mismo folder.
- Para programacion automatica, usar `Ejecutar_Programado_PMO_OpenAI.bat`; este archivo no hace `pause` y escribe logs en `logs\pmo_dashboard.log`.

Task Scheduler:

1. Abrir `Task Scheduler`.
2. Crear una tarea nueva.
3. Trigger: diario, semanal o al horario que necesites.
4. Action:
   - Program/script: ruta completa a `Ejecutar_Programado_PMO_OpenAI.bat`
   - Start in: ruta del folder `Generador Resumen Ejecutivo`
5. Activar `Run whether user is logged on or not` solo si la maquina tiene acceso a OneDrive y archivos locales en ese contexto.
6. Revisar `logs\pmo_dashboard.log` despues de la primera ejecucion.

Formato esperado del Excel:

- Archivo `.xlsx`.
- Una hoja con tabla EVM exportada desde MS Project.
- Encabezados detectables, preferiblemente: `Nombre de tarea`, `PV`, `EV`, `AC`, `SV`, `CV`, `SPI`, `CPI`, `EAC`, `BAC`, `VAC`, `IRPC` o `TCPI`.
- Minimo recomendado: `Nombre de tarea`, `PV`, `EV`, `AC`, `BAC`.
- Campos opcionales que mejoran el reporte: `WBS`, `Responsable`, `Fecha inicio`, `Fecha fin`, `% completado`, `Float/Holgura`, `Ruta critica`, `Hitos`, `Estado`.

## Campos calculados

- SPI = EV / PV
- CPI = EV / AC
- SV = EV - PV
- CV = EV - AC
- EAC = BAC / CPI
- ETC = EAC - AC
- VAC = BAC - EAC
- TCPI = (BAC - EV) / (BAC - AC)

Cuando el Excel no trae WBS, responsable, fechas, float, ruta critica, hitos o estado, el dashboard lo reporta como limitacion de fuente en vez de inventarlo.
