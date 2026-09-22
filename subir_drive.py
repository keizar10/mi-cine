#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sustituye el contenido del Excel en Google Drive (mismo archivo, mismo
enlace) usando una cuenta de servicio.

VARIABLES DE ENTORNO:
    GDRIVE_SA_JSON   contenido completo del JSON de la cuenta de servicio
    GDRIVE_FILE_ID   id del archivo en Drive (lo que va tras /file/d/ en el enlace)

USO:  python3 subir_drive.py "Movie Data Base by Keizar.xlsx"
REQUISITOS:  pip install google-auth requests
"""
import json
import os
import sys
from pathlib import Path

import requests
from google.oauth2 import service_account
from google.auth.transport.requests import Request

MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main(path):
    xlsx = Path(path)
    if not xlsx.exists():
        sys.exit(f"ERROR: no existe {xlsx}")

    sa_json = os.environ.get("GDRIVE_SA_JSON")
    file_id = os.environ.get("GDRIVE_FILE_ID")
    if not sa_json or not file_id:
        sys.exit("ERROR: faltan GDRIVE_SA_JSON o GDRIVE_FILE_ID")

    creds = service_account.Credentials.from_service_account_info(
        json.loads(sa_json), scopes=["https://www.googleapis.com/auth/drive"])
    creds.refresh(Request())
    headers = {"Authorization": f"Bearer {creds.token}"}

    # comprobar que la cuenta de servicio ve el archivo
    meta = requests.get(f"https://www.googleapis.com/drive/v3/files/{file_id}",
                        params={"fields": "name,size,modifiedTime",
                                "supportsAllDrives": "true"},
                        headers=headers, timeout=60)
    if meta.status_code == 404:
        sys.exit("ERROR: la cuenta de servicio no tiene acceso al archivo. "
                 "Compártelo con su email (client_email del JSON) como Editor.")
    meta.raise_for_status()
    print(f"Drive: {meta.json()['name']} ({meta.json().get('size')} bytes, "
          f"modificado {meta.json()['modifiedTime']})")

    # subir el nuevo contenido sobre el mismo archivo
    with open(xlsx, "rb") as f:
        r = requests.patch(
            f"https://www.googleapis.com/upload/drive/v3/files/{file_id}",
            params={"uploadType": "media", "supportsAllDrives": "true"},
            headers={**headers, "Content-Type": MIME},
            data=f, timeout=300)
    r.raise_for_status()
    print(f"OK: subido {xlsx.name} ({xlsx.stat().st_size} bytes). "
          f"El enlace sigue siendo el mismo.")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "Movie Data Base by Keizar.xlsx")
