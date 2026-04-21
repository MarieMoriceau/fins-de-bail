"""
notion_sync.py — Push les données vers la base Notion "Baux".
Ne fait rien si NOTION_TOKEN / BAUX_DATABASE_ID absents.
"""

import os
import logging
import time
from datetime import datetime, date
from typing import Any

import requests

log = logging.getLogger("notion_sync")

NOTION_TOKEN     = os.environ.get("NOTION_TOKEN")
BAUX_DATABASE_ID = os.environ.get("BAUX_DATABASE_ID")

NOTION_VERSION = "2022-06-28"
NOTION_API     = "https://api.notion.com/v1"
PIPEDRIVE_DOMAIN = "equationsie"

TRANCHE_MAP = {
    "Fin < 6 mois":  "< 6 mois",
    "Fin 6-9 mois":  "6-9 mois",
    "Fin 9-12 mois": "9-12 mois",
}
TRANCHE_PRIO = {"< 6 mois": 0, "6-9 mois": 1, "9-12 mois": 2}

PROP = {
    "id":       "id société",
    "pipe":     "lien pipe",
    "adresse":  "adresse",
    "cp":       "Code postal",
    "ville":    "Ville",
    "entree":   "Entrée dans les lieux",
    "fin_bail": "FIn de bail",
    "mois":     "Mois restants",
    "siren":    "SIREN",
    "nego":     "Négo",
    "tranche":  "Tranche d urgence",
    "campagne": "campagne",
    "archive":  "Archivé",
}
CAMPAGNE_TAG = "Fin de bail"


def is_configured() -> bool:
    return bool(NOTION_TOKEN and BAUX_DATABASE_ID)


def _headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }


def _api(method, path, payload=None, max_retries=3):
    for attempt in range(max_retries):
        try:
            r = getattr(requests, method)(
                f"{NOTION_API}{path}", headers=_headers(),
                json=payload, timeout=30
            )
            r.raise_for_status()
            return r.json()
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 429:
                wait = float(e.response.headers.get("Retry-After", 1.5))
                log.warning(f"  Rate limited, attente {wait}s...")
                time.sleep(wait)
            else:
                raise
    raise RuntimeError(f"Notion: {max_retries} retries échouées pour {path}")


def _detect_title_property() -> str:
    db = _api("get", f"/databases/{BAUX_DATABASE_ID}")
    for name, meta in db["properties"].items():
        if meta["type"] == "title":
            return name
    raise RuntimeError("Aucune propriété Title trouvée")


def _fetch_all_notion_rows() -> dict[str, dict]:
    result = {}
    payload = {"page_size": 100}
    has_more, cursor = True, None
    while has_more:
        if cursor:
            payload["start_cursor"] = cursor
        resp = _api("post", f"/databases/{BAUX_DATABASE_ID}/query", payload)
        for page in resp["results"]:
            rich = page["properties"].get(PROP["id"], {}).get("rich_text", [])
            id_val = "".join(t["plain_text"] for t in rich).strip()
            if id_val:
                result[id_val] = page
        has_more = resp.get("has_more", False)
        cursor = resp.get("next_cursor")
    log.info(f"  Notion: {len(result)} lignes existantes")
    return result


def _to_iso(v):
    if v is None or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    try:
        return datetime.fromisoformat(str(v)).strftime("%Y-%m-%d")
    except Exception:
        return None


def _build_props(row, tranche, title_prop, existing=None):
    props = {}
    if row.get("societe"):
        props[title_prop] = {"title": [{"text": {"content": str(row["societe"])[:2000]}}]}
    if existing is None:
        props[PROP["id"]] = {"rich_text": [{"text": {"content": str(row["id"])}}]}

    props[PROP["pipe"]] = {"url": f"https://{PIPEDRIVE_DOMAIN}.pipedrive.com/organization/{row['id']}"}

    for key, prop in [("adresse", PROP["adresse"]), ("cp", PROP["cp"]),
                      ("ville", PROP["ville"]), ("siren", PROP["siren"])]:
        val = row.get(key)
        if val is not None and str(val).strip():
            props[prop] = {"rich_text": [{"text": {"content": str(val)[:2000]}}]}

    if (d := _to_iso(row.get("entree"))):
        props[PROP["entree"]] = {"date": {"start": d}}
    if (d := _to_iso(row.get("fin"))):
        props[PROP["fin_bail"]] = {"date": {"start": d}}

    mois = row.get("mois")
    try:
        props[PROP["mois"]] = {"number": int(mois)} if mois is not None else {"number": None}
    except (ValueError, TypeError):
        props[PROP["mois"]] = {"number": None}

    props[PROP["tranche"]] = {"rich_text": [{"text": {"content": tranche}}]}

    existing_tags = []
    if existing:
        existing_tags = [t["name"] for t in existing["properties"]
                         .get(PROP["campagne"], {}).get("multi_select", [])]
    tags = sorted(set(existing_tags) | {CAMPAGNE_TAG})
    props[PROP["campagne"]] = {"multi_select": [{"name": t} for t in tags]}
    props[PROP["archive"]] = {"checkbox": False}
    return props


def sync_to_notion(cat_rows: dict) -> dict:
    if not is_configured():
        return {"skipped": True}

    stats = {"created": 0, "updated": 0, "archived": 0, "errors": 0}
    title_prop = _detect_title_property()

    deduped = {}
    for sheet_name, rows in cat_rows.items():
        tranche = TRANCHE_MAP[sheet_name]
        for r in rows:
            id_str = str(r["id"])
            prev = deduped.get(id_str)
            if prev is None or TRANCHE_PRIO[tranche] < TRANCHE_PRIO[prev[1]]:
                deduped[id_str] = (r, tranche)

    log.info(f"  {len(deduped)} sociétés à synchroniser")
    existing = _fetch_all_notion_rows()

    for id_str, (row, tranche) in deduped.items():
        try:
            page = existing.get(id_str)
            props = _build_props(row, tranche, title_prop, page)
            if page is None:
                _api("post", "/pages", {"parent": {"database_id": BAUX_DATABASE_ID}, "properties": props})
                stats["created"] += 1
            else:
                _api("patch", f"/pages/{page['id']}", {"properties": props})
                stats["updated"] += 1
            time.sleep(0.35)
        except Exception as e:
            stats["errors"] += 1
            log.error(f"  ✗ ID={id_str} → {e}")

    for id_str in (set(existing) - set(deduped)):
        page = existing[id_str]
        if page["properties"].get(PROP["archive"], {}).get("checkbox", False):
            continue
        try:
            _api("patch", f"/pages/{page['id']}", {"properties": {PROP["archive"]: {"checkbox": True}}})
            stats["archived"] += 1
            time.sleep(0.35)
        except Exception as e:
            stats["errors"] += 1
            log.error(f"  ✗ Archive ID={id_str} → {e}")

    log.info(f"=== SYNC TERMINÉE === {stats}")
    return stats
