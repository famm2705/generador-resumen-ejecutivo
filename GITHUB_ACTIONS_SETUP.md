# GitHub Actions + OneDrive Personal Setup

Este setup usa GitHub Actions como motor cloud y OneDrive personal como almacenamiento.

OneDrive personal no debe configurarse con `MS_CLIENT_SECRET` ni client credentials. Para ejecuciones programadas se usa OAuth delegado: haces login una vez, obtienes un `refresh_token` y lo guardas como GitHub Secret.

## 1. Subir este proyecto a GitHub

Crear un repositorio privado y subir estos archivos:

```text
generar_resumen_ejecutivo.py
cloud_onedrive_runner.py
get_onedrive_refresh_token.py
requirements.txt
Prompt.txt
.github/workflows/pmo-dashboard.yml
.gitignore
```

No subir `.venv`, `.env`, logs, Excel con datos sensibles, outputs generados ni API keys. El `.gitignore` ya excluye esos archivos.

## 2. Preparar OneDrive personal

Crear esta estructura en OneDrive:

```text
PMO EVM Reporting/
  Input/
    EVM_Actual.xlsx
  Output/
  Prompt/
    Prompt.txt
```

El Excel debe ser `.xlsx` y contener una hoja con columnas EVM detectables:

- `Nombre de tarea`
- `PV`
- `EV`
- `AC`
- `BAC`

Columnas recomendadas:

- `SV`, `CV`, `SPI`, `CPI`, `EAC`, `VAC`, `IRPC` o `TCPI`
- `WBS`, `Responsable`, `Fecha inicio`, `Fecha fin`, `% completado`, `Float/Holgura`, `Ruta critica`, `Hitos`, `Estado`

## 3. Crear app registration

Aunque no uses Azure compute, Microsoft Graph requiere registrar una app.

Importante: una cuenta personal `hotmail.com`, `outlook.com` o `live.com` puede mostrar el mensaje `La capacidad de crear aplicaciones fuera de un directorio esta en desuso`. Eso es esperado. El dueño del OneDrive personal no necesita crear la app.

La app debe crearse en cualquier Microsoft Entra directory disponible:

- tu directorio corporativo, si tienes permiso para crear app registrations;
- un directorio gratuito creado al registrarte en Azure;
- un tenant de Microsoft 365 Developer, si calificas.

Despues, el dueño del OneDrive personal solo inicia sesion una vez y autoriza la app.

1. Ir a `https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade`
2. New registration.
3. Supported account types: seleccionar una opcion que incluya `personal Microsoft accounts`.
4. Copiar `Application (client) ID`; sera `MS_CLIENT_ID`.
5. Authentication:
   - Agregar plataforma `Mobile and desktop applications`.
   - Agregar redirect URI `https://login.microsoftonline.com/common/oauth2/nativeclient`.
   - Habilitar `Allow public client flows` si aparece esa opcion.
6. API permissions > Microsoft Graph > Delegated permissions.
7. Agregar:
   - `Files.ReadWrite.All`
   - `offline_access`
8. No necesitas client secret para OneDrive personal.

## 4. Obtener refresh token una sola vez

En tu computadora:

```powershell
$env:MS_CLIENT_ID="client-id-de-la-app"
python get_onedrive_refresh_token.py
```

El script mostrara una URL y un codigo. Entra con la cuenta Microsoft personal donde vive OneDrive.

Al finalizar imprimira un valor largo. Guardalo como GitHub Secret:

```text
ONEDRIVE_REFRESH_TOKEN
```

No lo pongas en OneDrive ni en el repo.

Antes de ejecutar, verifica que `MS_CLIENT_ID` sea el `Application (client) ID` de la app que ya creaste en Entra.

## 5. Configurar GitHub Secrets

En el repo:

Settings > Secrets and variables > Actions > Secrets.

Crear:

```text
MS_CLIENT_ID
ONEDRIVE_REFRESH_TOKEN
OPENAI_API_KEY
```

Donde:

- `MS_CLIENT_ID`: `Application (client) ID` de la app en Entra.
- `ONEDRIVE_REFRESH_TOKEN`: token generado por `get_onedrive_refresh_token.py` despues de que el owner autoriza.
- `OPENAI_API_KEY`: key de OpenAI para mejorar la narrativa ejecutiva.

Opcional:

```text
MS_TENANT_ID=consumers
```

Si no configuras `MS_TENANT_ID`, el runner usa `consumers`.

No crear estos secrets para OneDrive personal:

```text
MS_CLIENT_SECRET
MS_USER_ID
MS_DRIVE_ID
```

## 6. Configurar GitHub Variables

Settings > Secrets and variables > Actions > Variables.

Crear:

```text
ONEDRIVE_INPUT_PATH=PMO EVM Reporting/Input/EVM_Actual.xlsx
ONEDRIVE_PROMPT_PATH=PMO EVM Reporting/Prompt/Prompt.txt
ONEDRIVE_OUTPUT_FOLDER=PMO EVM Reporting/Output
OPENAI_MODEL=gpt-5.5
```

Estas variables no son secretas; describen rutas y configuracion del workflow.

Opcional si quieres sobrescribir siempre el mismo archivo:

```text
ONEDRIVE_OUTPUT_FILENAME=Resumen_PMO.xlsx
```

Si no defines `ONEDRIVE_OUTPUT_FILENAME`, cada corrida genera un archivo con timestamp.

## 7. Probar manualmente

En GitHub:

1. Ir a Actions.
2. Seleccionar `PMO EVM Dashboard`.
3. Run workflow.
4. Revisar logs.
5. Confirmar que el archivo aparece en OneDrive `Output/`.

## 8. Programacion

El workflow ya corre de lunes a viernes:

```yaml
- cron: "0 12 * * 1-5"
```

GitHub Actions usa UTC. Ajustar el cron segun la hora deseada.

## Checklist para la llamada con el owner

Antes de la llamada:

- Tener creada la app registration en Entra.
- Tener copiado el `MS_CLIENT_ID`.
- Tener el repo privado de GitHub creado.
- Tener listo el comando `python get_onedrive_refresh_token.py`.

Durante la llamada:

1. Confirmar que el Excel y el prompt estan en OneDrive personal.
2. Confirmar los paths exactos.
3. Ejecutar:

```powershell
$env:MS_CLIENT_ID="client-id-de-la-app"
python get_onedrive_refresh_token.py
```

4. El owner entra al link de Microsoft, pega el codigo e inicia sesion.
5. El owner acepta permisos.
6. Guardar el token impreso como GitHub Secret `ONEDRIVE_REFRESH_TOKEN`.
7. Probar el workflow manualmente en GitHub Actions.

## Notas de seguridad

- Usa un repositorio privado.
- No subas Excel sensible si no quieres que viva en GitHub; el workflow puede leer el archivo desde OneDrive.
- Guarda `OPENAI_API_KEY` y `ONEDRIVE_REFRESH_TOKEN` solo como GitHub Secrets.
- Si Microsoft rota el refresh token, el workflow mostrara una nota. Si empieza a fallar por token expirado/revocado, ejecuta de nuevo `get_onedrive_refresh_token.py` y actualiza el secret.

## Referencias oficiales

- GitHub Actions schedule usa sintaxis cron POSIX en workflows.
- GitHub Actions secrets son variables protegidas que el workflow solo lee si se incluyen explicitamente.
- Microsoft Graph permite acceder a archivos de OneDrive por ruta relativa al root.
- Para OneDrive personal, Microsoft Graph usa permisos delegados como `Files.ReadWrite.All` y `offline_access`.
