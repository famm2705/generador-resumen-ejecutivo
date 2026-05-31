# Deployment Checklist - Entra + GitHub Actions + OneDrive Personal

## 1. Repo GitHub

Crear un repo privado y subir el codigo del proyecto.

Archivos que deben subirse:

```text
.github/workflows/pmo-dashboard.yml
.gitignore
cloud_onedrive_runner.py
generar_resumen_ejecutivo.py
get_onedrive_refresh_token.py
requirements.txt
README.md
GITHUB_ACTIONS_SETUP.md
DEPLOYMENT_CHECKLIST.md
Prompt.txt
env.example
```

No subir:

```text
.venv/
.env
logs/
outputs/
work/
*.xlsx
~$*
```

## 2. App en Microsoft Entra

En la app registration ya creada, verificar:

```text
Supported account types:
Accounts in any organizational directory and personal Microsoft accounts
```

Authentication:

```text
Platform:
Mobile and desktop applications

Redirect URI:
https://login.microsoftonline.com/common/oauth2/nativeclient

Allow public client flows:
Enabled, si la opcion aparece
```

API permissions:

```text
Microsoft Graph > Delegated permissions
Files.ReadWrite.All
offline_access
```

Copiar:

```text
Application (client) ID
```

Ese valor se guarda como GitHub Secret `MS_CLIENT_ID`.

## 3. OneDrive personal del owner

Crear o confirmar esta estructura:

```text
PMO EVM Reporting/
  Input/
    EVM_Actual.xlsx
  Prompt/
    Prompt.txt
  Output/
```

El Excel debe tener una hoja EVM con columnas detectables:

```text
Nombre de tarea
PV
EV
AC
BAC
```

Mejor si tambien trae:

```text
SV
CV
SPI
CPI
EAC
VAC
IRPC o TCPI
WBS
Responsable
Fecha inicio
Fecha fin
Float/Holgura
Ruta critica
Hitos
Estado
```

## 4. Obtener token del owner

En esta carpeta local:

```powershell
$env:MS_CLIENT_ID="application-client-id"
python get_onedrive_refresh_token.py
```

El owner abre la URL, pega el codigo, inicia sesion con su cuenta personal de OneDrive y acepta permisos.

El script imprime un token largo. Guardarlo como GitHub Secret:

```text
ONEDRIVE_REFRESH_TOKEN
```

No guardar ese token en OneDrive, Excel, `.env`, README ni chat.

## 5. GitHub Secrets

Repo > Settings > Secrets and variables > Actions > Secrets:

```text
MS_CLIENT_ID
ONEDRIVE_REFRESH_TOKEN
OPENAI_API_KEY
```

## 6. GitHub Variables

Repo > Settings > Secrets and variables > Actions > Variables:

```text
ONEDRIVE_INPUT_PATH=PMO EVM Reporting/Input/EVM_Actual.xlsx
ONEDRIVE_PROMPT_PATH=PMO EVM Reporting/Prompt/Prompt.txt
ONEDRIVE_OUTPUT_FOLDER=PMO EVM Reporting/Output
OPENAI_MODEL=gpt-5.5
```

Opcional:

```text
ONEDRIVE_OUTPUT_FILENAME=Resumen_PMO.xlsx
```

Si no defines `ONEDRIVE_OUTPUT_FILENAME`, cada corrida genera un archivo con timestamp.

## 7. Probar

1. GitHub > Actions.
2. Seleccionar `PMO EVM Dashboard`.
3. `Run workflow`.
4. Revisar logs.
5. Confirmar que el dashboard aparece en OneDrive `PMO EVM Reporting/Output`.
