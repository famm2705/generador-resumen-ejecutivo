from __future__ import annotations

import json
import os
import socket
import time
import urllib.error
import urllib.parse
import urllib.request


TENANT_ID = os.getenv("MS_TENANT_ID") or "consumers"
CLIENT_ID = os.getenv("MS_CLIENT_ID")
SCOPES = "https://graph.microsoft.com/Files.ReadWrite.All offline_access"
DEVICE_CODE_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/devicecode"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"


def request_json(url: str, data: dict[str, str]) -> dict:
    request = urllib.request.Request(
        url,
        data=urllib.parse.urlencode(data).encode("utf-8"),
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=25) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(details)
        except json.JSONDecodeError:
            raise RuntimeError(f"HTTP {exc.code}: {details}") from exc
        if payload.get("error") == "invalid_client" and "marked as 'mobile'" in payload.get("error_description", ""):
            message = mobile_client_error_message(payload)
            raise RuntimeError(message) from exc
        raise RuntimeError(json.dumps(payload, indent=2, ensure_ascii=False)) from exc
    except (TimeoutError, socket.timeout, urllib.error.URLError, OSError) as exc:
        raise RuntimeError(f"network_wait: no hubo respuesta de Microsoft a tiempo ({exc}).") from exc


def mobile_client_error_message(payload: dict) -> str:
    message = """
La app de Entra no esta marcada como cliente publico/mobile.

Corrige esto en Microsoft Entra:
1. App registrations > tu app > Authentication.
2. Add a platform > Mobile and desktop applications.
3. Agrega este Redirect URI:
   https://login.microsoftonline.com/common/oauth2/nativeclient
4. En Advanced settings, activa:
   Allow public client flows = Yes
5. Guarda los cambios y vuelve a ejecutar este script.

Detalle original de Microsoft:
"""
    return message + json.dumps(payload, indent=2, ensure_ascii=False)


def main() -> int:
    if not CLIENT_ID:
        raise SystemExit("Configura MS_CLIENT_ID antes de ejecutar este script.")

    device = request_json(
        DEVICE_CODE_URL,
        {
            "client_id": CLIENT_ID,
            "scope": SCOPES,
        },
    )
    print(device["message"])
    print()
    print("Esperando autorizacion...")

    deadline = time.time() + int(device.get("expires_in", 900))
    interval = int(device.get("interval", 5))
    while time.time() < deadline:
        time.sleep(interval)
        try:
            token = request_json(
                TOKEN_URL,
                {
                    "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                    "client_id": CLIENT_ID,
                    "device_code": device["device_code"],
                },
            )
        except RuntimeError as exc:
            text = str(exc)
            if "authorization_pending" in text:
                continue
            if "slow_down" in text:
                interval += 5
                continue
            if "network_wait" in text:
                print("Sin respuesta temporal de Microsoft; reintentando...")
                continue
            raise

        refresh_token = token.get("refresh_token")
        if not refresh_token:
            raise SystemExit("Microsoft no devolvio refresh_token. Verifica que el scope offline_access este permitido.")
        print()
        print("Guarda este valor como GitHub Secret ONEDRIVE_REFRESH_TOKEN:")
        print(refresh_token)
        return 0

    raise SystemExit("Tiempo agotado esperando autorizacion.")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(
            "\nProceso interrumpido. Vuelve a ejecutar el script y completa el login en el navegador antes de cerrar PowerShell."
        )
