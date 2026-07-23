from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from generar_resumen_ejecutivo import generate


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
GRAPH_SCOPES = "https://graph.microsoft.com/Files.ReadWrite.All offline_access"
LOCAL_INPUT = Path("work/input.xlsx")
LOCAL_OUTPUT = Path("work/output_Resumen_PMO.xlsx")
LOCAL_PROMPT = Path("Prompt.txt")
EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}


def required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise SystemExit(f"Falta configurar el secret/variable {name}.")
    return value


def graph_token() -> str:
    tenant_id = os.getenv("MS_TENANT_ID") or "consumers"
    form = {
        "client_id": required_env("MS_CLIENT_ID"),
        "refresh_token": required_env("ONEDRIVE_REFRESH_TOKEN"),
        "scope": GRAPH_SCOPES,
        "grant_type": "refresh_token",
    }
    data = urllib.parse.urlencode(form).encode("utf-8")
    request = urllib.request.Request(
        TOKEN_URL.format(tenant_id=tenant_id),
        data=data,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    payload = request_json(request)
    rotated_refresh_token = payload.get("refresh_token")
    if rotated_refresh_token and rotated_refresh_token != os.getenv("ONEDRIVE_REFRESH_TOKEN"):
        print("Microsoft devolvio un refresh token renovado. Actualiza el secret ONEDRIVE_REFRESH_TOKEN si el actual expira.")
    return payload["access_token"]


def drive_root() -> str:
    return "/me/drive"


def download_file(token: str, remote_path: str, local_path: Path) -> None:
    local_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        metadata = graph_get_json(token, f"{drive_root()}/root:/{quote_drive_path(remote_path)}")
        download_url = metadata.get("@microsoft.graph.downloadUrl")
        if not download_url:
            raise RuntimeError(f"El item existe, pero Microsoft Graph no devolvio downloadUrl. Tipo detectado: {metadata.get('folder') or metadata.get('file') or metadata.keys()}")
        request = urllib.request.Request(download_url, headers={"User-Agent": "pmo-dashboard-runner"}, method="GET")
        with urllib.request.urlopen(request, timeout=120) as response:
            local_path.write_bytes(response.read())
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            diagnostics = one_drive_path_diagnostics(token, remote_path)
            raise RuntimeError(
                "OneDrive no encontro el archivo solicitado.\n"
                f"Ruta configurada: {remote_path}\n"
                f"URL Graph: {GRAPH_BASE}{drive_root()}/root:/{quote_drive_path(remote_path)}\n\n"
                f"{diagnostics}\n"
                "Corrige ONEDRIVE_INPUT_PATH u ONEDRIVE_PROMPT_PATH en GitHub Variables."
            ) from exc
        if exc.code == 401:
            raise RuntimeError(
                "Microsoft devolvio 401 al descargar el archivo. "
                "Verifica que el refresh token sea de la misma cuenta OneDrive que contiene el archivo, "
                "y que la app tenga permisos delegados Files.ReadWrite.All y offline_access."
            ) from exc
        raise


def list_folder_children(token: str, remote_folder: str) -> list[dict]:
    response = graph_get_json(token, f"{drive_root()}/root:/{quote_drive_path(remote_folder)}:/children")
    return response.get("value", [])


def latest_excel_in_folder(token: str, remote_folder: str) -> str:
    children = list_folder_children(token, remote_folder)
    excel_files = [
        item
        for item in children
        if item.get("file")
        and not str(item.get("name", "")).startswith("~$")
        and Path(str(item.get("name", ""))).suffix.lower() in EXCEL_EXTENSIONS
    ]
    if not excel_files:
        names = ", ".join(str(item.get("name", "(sin nombre)")) for item in children[:30])
        raise RuntimeError(
            "No se encontro ningun archivo Excel valido para usar como input.\n"
            f"Folder configurado: {remote_folder}\n"
            f"Contenido detectado: {names if names else '(vacio)'}"
        )
    latest = max(excel_files, key=lambda item: item.get("lastModifiedDateTime") or item.get("createdDateTime") or "")
    return f"{remote_folder.strip().strip('/')}/{latest['name']}"


def parent_folder(remote_path: str) -> str:
    parts = [part for part in remote_path.replace("\\", "/").split("/") if part]
    if len(parts) <= 1:
        return ""
    return "/".join(parts[:-1])


def resolve_input_path(token: str) -> str:
    configured_path = required_env("ONEDRIVE_INPUT_PATH").strip().strip("/")
    use_latest = os.getenv("ONEDRIVE_INPUT_LATEST", "").lower() in {"1", "true", "yes", "y"}
    if not use_latest:
        return configured_path

    folder = (os.getenv("ONEDRIVE_INPUT_FOLDER") or parent_folder(configured_path)).strip().strip("/")
    if not folder:
        raise RuntimeError(
            "ONEDRIVE_INPUT_LATEST=true requiere ONEDRIVE_INPUT_FOLDER, "
            "o que ONEDRIVE_INPUT_PATH incluya un folder padre."
        )
    latest_path = latest_excel_in_folder(token, folder)
    print(f"Input mas reciente detectado en OneDrive: {latest_path}")
    return latest_path


def upload_file(token: str, local_path: Path, remote_path: str) -> dict:
    url = f"{GRAPH_BASE}{drive_root()}/root:/{quote_drive_path(remote_path)}:/content"
    request = urllib.request.Request(
        url,
        data=local_path.read_bytes(),
        headers={**auth_headers(token), "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        method="PUT",
    )
    return request_json(request)


def request_json(request: urllib.request.Request) -> dict:
    try:
        with urllib.request.urlopen(request, timeout=120) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {exc.code}: {details}") from exc
    return json.loads(body) if body else {}


def graph_get_json(token: str, path: str) -> dict:
    request = urllib.request.Request(f"{GRAPH_BASE}{path}", headers=auth_headers(token), method="GET")
    return request_json(request)


def one_drive_path_diagnostics(token: str, remote_path: str) -> str:
    lines = ["Diagnostico OneDrive:"]
    try:
        root = graph_get_json(token, f"{drive_root()}/root/children")
        names = [item.get("name", "(sin nombre)") for item in root.get("value", [])]
        lines.append("Root contiene: " + (", ".join(names[:30]) if names else "(vacio o no visible)"))
    except Exception as exc:
        lines.append(f"No se pudo listar root: {exc}")

    parts = [part for part in remote_path.replace("\\", "/").split("/") if part]
    for depth in range(1, min(len(parts), 3) + 1):
        folder = "/".join(parts[:depth])
        try:
            children = graph_get_json(token, f"{drive_root()}/root:/{quote_drive_path(folder)}:/children")
            names = [item.get("name", "(sin nombre)") for item in children.get("value", [])]
            lines.append(f"Contenido de '{folder}': " + (", ".join(names[:30]) if names else "(vacio o no visible)"))
        except Exception as exc:
            lines.append(f"No se pudo listar '{folder}': {exc}")
            break
    return "\n".join(lines)


def auth_headers(token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}"}


def quote_drive_path(path: str) -> str:
    parts = [part for part in path.replace("\\", "/").split("/") if part]
    return "/".join(urllib.parse.quote(part, safe="") for part in parts)


def quote_path_part(value: str) -> str:
    return urllib.parse.quote(value, safe="@.-_")


def output_remote_path() -> str:
    folder = os.getenv("ONEDRIVE_OUTPUT_FOLDER", "PMO EVM Reporting/Output").strip().strip("/")
    if os.getenv("ONEDRIVE_OUTPUT_NEXT_MONDAY", "").lower() in {"1", "true", "yes", "y"}:
        filename = f"Resumen_PMO_{next_monday_panama():%d_%m_%Y}.xlsx"
        return f"{folder}/{filename}"

    filename = os.getenv("ONEDRIVE_OUTPUT_FILENAME")
    if not filename:
        filename = f"Resumen_PMO_{datetime.utcnow():%Y%m%d_%H%M%S}Z.xlsx"
    return f"{folder}/{filename}"


def next_monday_panama() -> date:
    try:
        panama_tz = ZoneInfo("America/Panama")
    except ZoneInfoNotFoundError:
        # Panama does not use daylight saving time; UTC-5 is a safe fallback
        # for Windows/Python environments without the tzdata package installed.
        panama_tz = timezone(timedelta(hours=-5), name="America/Panama")
    today = datetime.now(panama_tz).date()
    days_until_monday = (0 - today.weekday()) % 7
    return today + timedelta(days=days_until_monday)


def main() -> int:
    token = graph_token()
    input_path = resolve_input_path(token)
    prompt_path = os.getenv("ONEDRIVE_PROMPT_PATH")

    print(f"Descargando input desde OneDrive: {input_path}")
    download_file(token, input_path, LOCAL_INPUT)

    if prompt_path:
        print(f"Descargando prompt desde OneDrive: {prompt_path}")
        download_file(token, prompt_path, LOCAL_PROMPT)
    elif not LOCAL_PROMPT.exists():
        LOCAL_PROMPT.write_text("", encoding="utf-8")

    use_openai = os.getenv("USE_OPENAI", "true").lower() in {"1", "true", "yes", "y"}
    print("Generando dashboard PMO...")
    generated_path = generate(LOCAL_INPUT, LOCAL_OUTPUT, Path.cwd(), use_openai=use_openai)

    remote_output = output_remote_path()
    print(f"Subiendo resultado a OneDrive: {remote_output}")
    upload_result = upload_file(token, generated_path, remote_output)
    web_url = upload_result.get("webUrl", "(sin webUrl en respuesta)")
    print(f"Dashboard publicado: {web_url}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
