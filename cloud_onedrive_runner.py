from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

from generar_resumen_ejecutivo import generate


GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_URL = "https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
GRAPH_SCOPES = "https://graph.microsoft.com/Files.ReadWrite.All offline_access"
LOCAL_INPUT = Path("work/input.xlsx")
LOCAL_OUTPUT = Path("work/output_Resumen_PMO.xlsx")
LOCAL_PROMPT = Path("Prompt.txt")


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
    url = f"{GRAPH_BASE}{drive_root()}/root:/{quote_drive_path(remote_path)}:/content"
    request = urllib.request.Request(url, headers=auth_headers(token), method="GET")
    with urllib.request.urlopen(request, timeout=120) as response:
        local_path.write_bytes(response.read())


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


def auth_headers(token: str) -> dict[str, str]:
    return {"Authorization": f"Bearer {token}"}


def quote_drive_path(path: str) -> str:
    parts = [part for part in path.replace("\\", "/").split("/") if part]
    return "/".join(urllib.parse.quote(part, safe="") for part in parts)


def quote_path_part(value: str) -> str:
    return urllib.parse.quote(value, safe="@.-_")


def output_remote_path() -> str:
    folder = os.getenv("ONEDRIVE_OUTPUT_FOLDER", "PMO EVM Reporting/Output").strip().strip("/")
    filename = os.getenv("ONEDRIVE_OUTPUT_FILENAME")
    if not filename:
        filename = f"Resumen_PMO_{datetime.utcnow():%Y%m%d_%H%M%S}Z.xlsx"
    return f"{folder}/{filename}"


def main() -> int:
    token = graph_token()
    input_path = required_env("ONEDRIVE_INPUT_PATH")
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
