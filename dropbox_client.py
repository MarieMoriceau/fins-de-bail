"""
dropbox_client.py — Récupère/upload des fichiers .xlsx depuis/vers Dropbox.

Utilise un refresh token permanent (pas d'expiration).
Variables d'env : DROPBOX_REFRESH_TOKEN, DROPBOX_APP_KEY, DROPBOX_APP_SECRET
"""

import json
import os
import logging

import requests

log = logging.getLogger("dropbox")

REFRESH_TOKEN = os.environ.get("DROPBOX_REFRESH_TOKEN")
APP_KEY       = os.environ.get("DROPBOX_APP_KEY")
APP_SECRET    = os.environ.get("DROPBOX_APP_SECRET")

DROPBOX_FOLDER = "/FINS DE BAIL"


def is_configured() -> bool:
    return all([REFRESH_TOKEN, APP_KEY, APP_SECRET])


def _get_access_token() -> str:
    r = requests.post("https://api.dropboxapi.com/oauth2/token", data={
        "grant_type": "refresh_token",
        "refresh_token": REFRESH_TOKEN,
        "client_id": APP_KEY,
        "client_secret": APP_SECRET,
    }, timeout=15)
    r.raise_for_status()
    log.info("Dropbox: access token obtenu")
    return r.json()["access_token"]


def fetch_xlsx_files() -> list[tuple[str, bytes]]:
    """Télécharge tous les .xlsx du dossier Dropbox."""
    if not is_configured():
        return []

    token = _get_access_token()
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    r = requests.post("https://api.dropboxapi.com/2/files/list_folder",
        headers=headers, json={"path": DROPBOX_FOLDER}, timeout=15)
    r.raise_for_status()

    entries = [e for e in r.json().get("entries", [])
               if e[".tag"] == "file" and e["name"].lower().endswith(".xlsx")]
    log.info(f"Dropbox: {len(entries)} fichier(s) .xlsx dans {DROPBOX_FOLDER}")

    results = []
    for entry in entries:
        log.info(f"  ↓ {entry['name']}")
        dl = requests.post("https://content.dropboxapi.com/2/files/download",
            headers={
                "Authorization": f"Bearer {token}",
                "Dropbox-API-Arg": json.dumps({"path": entry["path_lower"]}),
            }, timeout=120)
        dl.raise_for_status()
        results.append((entry["name"], dl.content))

    return results


def upload_file(filename: str, content: bytes):
    """Upload un fichier dans le dossier Dropbox (écrase si existant)."""
    if not is_configured():
        return
    token = _get_access_token()
    api_arg = json.dumps({
        "path": f"{DROPBOX_FOLDER}/{filename}",
        "mode": "overwrite",
        "autorename": False,
        "mute": False,
        "strict_conflict": False,
    })
    r = requests.post("https://content.dropboxapi.com/2/files/upload",
        headers={
            "Authorization": f"Bearer {token}",
            "Dropbox-API-Arg": api_arg,
            "Content-Type": "application/octet-stream",
        }, data=content, timeout=120)
    if not r.ok:
        log.error(f"Dropbox upload detail: {r.status_code} {r.text}")
    r.raise_for_status()
    log.info(f"Dropbox: ↑ {filename} uploadé")
