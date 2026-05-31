from __future__ import annotations

import json
import os
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
        with urllib.request.urlopen(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        try:
            payload = json.loads(details)
        except json.JSONDecodeError:
            raise RuntimeError(f"HTTP {exc.code}: {details}") from exc
        raise RuntimeError(json.dumps(payload, indent=2, ensure_ascii=False)) from exc


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
    raise SystemExit(main())
