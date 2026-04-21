"""
notion_sync.py — Push les données fins de bail vers la base Notion "Baux".

Appelé automatiquement après la génération Excel.
Ne fait rien si NOTION_TOKEN n'est pas configuré (mode dégradé = Excel seul).
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

# Noms exacts des propriétés Notion (attention à la casse !)
PROP = {
    "id":       "id société",
    "pipe":     "lien pipe",
    "adresse":  "adresse",
    "cp":       "Code postal",
    "ville":    "Ville",
    "entree":   "Entrée dans les lieux",
    "fin_bail": "FIn de bail",       # ⚠ I majuscule
    "mois":     "Mois restants",
    "siren":    "SIREN",
    "nego":     "Négo",
    "tranche":  "Tranche d urgence", # ⚠ sans apostrophe
    "campagne": "campagne",
    "archive":  "archivé",
}

CAMPAGNE_TAG = "Fin de bail"


def is_configured() -> bool:
    """Renvoie True si les variables Notion sont présentes."""
    return bool(NOTION_TOKEN and BAUX_DATABASE_ID)


def _headers():
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }


def _notion_get(path: str) -> dict:
    r = requests.get(f"{NOTION_API}{path}", headers=_headers(), timeout=30)
    r.raise_for_status()
    return r.json()


def _notion_post(path: str, payload: dict) -> dict:
    r = requests.post(f"{NOTION_API}{path}", headers=_headers(), json=payload, timeout=30)
    r.raise_for_status()
    return r.json()


def _notion_patch(path: str, payload: dict) -> dict:
    r = requests.patch(f"{NOTION_API}{path}", headers=_headers(), json=payload, timeout=30)
    r.raise_for_status()
    return r.json()


def _notion_request_with_retry(method, path, payload=None, max_retries=3):
    """Wrapper avec retry sur rate-limit (429)."""
    for attempt in range(max_retries):
        try:
            if method == "GET":
                return _notion_get(path)
            elif method == "POST":
                return _notion_post(path, payload or {})
            elif method == "PATCH":
                return _notion_patch(path, payload or {})
        except requests.exceptions.HTTPError as e:
            if e.response is not None and e.response.status_code == 429:
                wait = float(e.response.headers.get("Retry-After", 1))
                log.warning(f"  Rate limited, attente {wait}s...")
                time.sleep(wait)
            else:
                raise
    raise RuntimeError(f"Notion API: {max_retries} tentatives échouées pour {path}")


def _detect_title_property() -> str:
    db = _notion_get(f"/databases/{BAUX_DATABASE_ID}")
    for name, meta in db["properties"].items():
        if meta["type"] == "title":
            return name
    raise RuntimeError("Aucune propriété Title trouvée dans la base Notion")


def _fetch_all_notion_rows() -> dict[str, dict]:
    """Récupère toutes les lignes existantes, indexées par 'id société'."""
    result: dict[str, dict] = {}
    payload: dict = {"page_size": 100}
    has_more, cursor = True, None
    while has_more:
        if cursor:
            payload["start_cursor"] = cursor
        resp = _notion_request_with_retry("POST", f"/databases/{BAUX_DATABASE_ID}/query", payload)
        for page in resp["results"]:
            id_prop = page["properties"].get(PROP["id"], {})
            rich = id_prop.get("rich_text", [])
            id_value = "".join(t["plain_text"] for t in rich).strip()
            if id_value:
                result[id_value] = page
        has_more = resp.get("has_more", False)
        cursor = resp.get("next_cursor")
    log.info(f"  Notion: {len(result)} lignes existantes")
    return result


def _to_iso_date(v: Any) -> str | None:
    if v is None or v == "":
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    try:
        return datetime.fromisoformat(str(v)).strftime("%Y-%m-%d")
    except Exception:
        return None


def _build_properties(row: dict, tranche: str, title_prop: str,
                      existing_page: dict | None = None) -> dict:
    props: dict = {}

    societe = row.get("societe")
    if societe:
        props[title_prop] = {"title": [{"text": {"content": str(societe)[:2000]}}]}

    # id société : uniquement à la création
    if existing_page is None:
        props[PROP["id"]] = {"rich_text": [{"text": {"content": str(row["id"])}}]}

    # Lien Pipedrive
    pipe_url = f"https://{PIPEDRIVE_DOMAIN}.pipedrive.com/organization/{row['id']}"
    props[PROP["pipe"]] = {"url": pipe_url}

    # Rich text : ne pas écraser par vide
    for key, prop_name in [("adresse", PROP["adresse"]), ("cp", PROP["cp"]),
                           ("ville", PROP["ville"]), ("siren", PROP["siren"])]:
        val = row.get(key)
        if val is not None and str(val).strip() != "":
            props[prop_name] = {"rich_text": [{"text": {"content": str(val)[:2000]}}]}

    # Dates : ne pas écraser par vide
    if (d := _to_iso_date(row.get("entree"))):
        props[PROP["entree"]] = {"date": {"start": d}}
    if (d := _to_iso_date(row.get("fin"))):
        props[PROP["fin_bail"]] = {"date": {"start": d}}

    # Mois restants : toujours écrasé
    mois = row.get("mois")
    try:
        props[PROP["mois"]] = {"number": int(mois)} if mois is not None else {"number": None}
    except (ValueError, TypeError):
        props[PROP["mois"]] = {"number": None}

    # Tranche : toujours écrasé
    props[PROP["tranche"]] = {"rich_text": [{"text": {"content": tranche}}]}

    # Négo : jamais écrasé (rempli à la main dans Notion)

    # Campagne : ajouter "Fin de bail" sans virer les autres tags
    existing_tags = []
    if existing_page:
        existing_tags = [
            t["name"] for t in existing_page["properties"]
                .get(PROP["campagne"], {}).get("multi_select", [])
        ]
    tags = sorted(set(existing_tags) | {CAMPAGNE_TAG})
    props[PROP["campagne"]] = {"multi_select": [{"name": t} for t in tags]}

    # Archivé = False (société vue dans l'Excel)
    props[PROP["archive"]] = {"checkbox": False}

    return props


def sync_to_notion(cat_rows: dict) -> dict:
    """
    Pousse les données vers Notion.
    cat_rows = {"Fin < 6 mois": [rows...], "Fin 6-9 mois": [...], "Fin 9-12 mois": [...]}
    Chaque row = dict avec id, societe, adresse, cp, ville, entree, fin, mois, siren.

    Retourne {"created": N, "updated": N, "archived": N, "errors": N}.
    """
    if not is_configured():
        log.warning("Notion non configuré (NOTION_TOKEN / BAUX_DATABASE_ID manquants)")
        return {"created": 0, "updated": 0, "archived": 0, "errors": 0, "skipped": True}

    stats = {"created": 0, "updated": 0, "archived": 0, "errors": 0}

    title_prop = _detect_title_property()
    log.info(f"Propriété Title Notion: '{title_prop}'")

    # Dédoublonner : garder la tranche la plus urgente
    deduped: dict[str, tuple[dict, str]] = {}
    for sheet_name, rows in cat_rows.items():
        tranche = TRANCHE_MAP[sheet_name]
        for r in rows:
            id_str = str(r["id"])
            prev = deduped.get(id_str)
            if prev is None or TRANCHE_PRIO[tranche] < TRANCHE_PRIO[prev[1]]:
                deduped[id_str] = (r, tranche)

    log.info(f"  {len(deduped)} sociétés à synchroniser")

    existing = _fetch_all_notion_rows()
    excel_ids = set(deduped.keys())
    notion_ids = set(existing.keys())

    # Create / Update
    for id_str, (row, tranche) in deduped.items():
        try:
            page = existing.get(id_str)
            props = _build_properties(row, tranche, title_prop, page)
            if page is None:
                _notion_request_with_retry("POST", "/pages", {
                    "parent": {"database_id": BAUX_DATABASE_ID},
                    "properties": props,
                })
                stats["created"] += 1
            else:
                _notion_request_with_retry("PATCH", f"/pages/{page['id']}", {"properties": props})
                stats["updated"] += 1
            # Petit throttle pour rester sous le rate limit (~3 req/s)
            time.sleep(0.35)
        except Exception as e:
            stats["errors"] += 1
            log.error(f"  ✗ ID={id_str} {row.get('societe')} → {e}")

    # Archive : lignes Notion absentes de l'Excel
    for id_str in (notion_ids - excel_ids):
        page = existing[id_str]
        if page["properties"].get(PROP["archive"], {}).get("checkbox", False):
            continue  # déjà archivé
        try:
            _notion_request_with_retry("PATCH", f"/pages/{page['id']}", {
                "properties": {PROP["archive"]: {"checkbox": True}}
            })
            stats["archived"] += 1
            time.sleep(0.35)
        except Exception as e:
            stats["errors"] += 1
            log.error(f"  ✗ Archive ID={id_str} → {e}")

    log.info(f"=== SYNC NOTION TERMINÉE === {stats}")
    return stats
