"""
app.py — Générateur fins de bail + sync Notion + mail + auto-refresh Dropbox.

Routes :
  GET  /           → page d'upload
  POST /upload     → lit headers, propose mapping
  POST /generate   → génère Excel + mail + Notion (background)
  GET  /template   → fichier template vierge
  POST /auto       → refresh mensuel automatique (Dropbox → recalcul → Notion + mail + Dropbox)

Optimisé mémoire pour Render Free (512 Mo).
"""

import gc
import json
import logging
import os
import tempfile
import threading
from datetime import date
from io import BytesIO

from flask import Flask, render_template, request, send_file, redirect, url_for, flash, jsonify

from generator import generate, read_headers, REQUIRED_FIELDS, OPTIONAL_FIELDS, DEFAULT_MAPPING
from generator import _write_call_sheet, _write_readme, CAT_META
from notion_sync import sync_to_notion, is_configured as notion_configured
from mailer import send_file as send_mail, is_configured as mail_configured
from dropbox_client import fetch_xlsx_files, upload_file as dropbox_upload, is_configured as dropbox_configured
from auto_refresh import extract_from_generated, categorize

import openpyxl

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
log = logging.getLogger("app")

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-secret-fins-de-bail-2026")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

AUTO_SECRET = os.environ.get("SYNC_SECRET", "")
AUTO_EMAIL  = os.environ.get("AUTO_EMAIL", "mmoriceau@equation-sie.com")

_temp_files = {}


# ─────────────────────────────────────────────
#  Pages web (upload manuel)
# ─────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html",
                           notion_ok=notion_configured(),
                           mail_ok=mail_configured(),
                           dropbox_ok=dropbox_configured())


@app.route("/template")
def download_template():
    return send_file("static/template_fins_de_bail.xlsx",
                     as_attachment=True,
                     download_name="template_fins_de_bail.xlsx")


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith(".xlsx"):
        flash("Fichier .xlsx attendu", "error")
        return redirect(url_for("index"))

    source_bytes = f.read()
    email = request.form.get("email", "").strip()

    try:
        headers = read_headers(source_bytes)
    except Exception as e:
        flash(f"Impossible de lire le fichier : {e}", "error")
        return redirect(url_for("index"))

    if not headers:
        flash("Fichier vide (pas de headers en ligne 1)", "error")
        return redirect(url_for("index"))

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(source_bytes)
    tmp.close()
    del source_bytes
    gc.collect()

    file_id = os.path.basename(tmp.name)
    _temp_files[file_id] = tmp.name

    # Auto-matching
    auto_map = {}
    header_lower = {col: h.lower() for col, h in headers}
    guesses = {
        "id":      ["id", "identifiant", "id pipedrive", "id société"],
        "societe": ["société", "societe", "nom", "raison sociale", "company"],
        "date":    ["date adresse retenue", "date d'entrée", "date entrée",
                    "📅 date adresse retenue", "date adresse",
                    "date entrée dans les lieux"],
        "siren":   ["siren", "siren pappers", "siren sirene", "n° siren"],
        "adresse": ["adresse", "adresse retenue", "✅ adresse retenue", "adresse complète"],
        "cp":      ["cp", "code postal", "siège cp", "zip"],
        "ville":   ["ville", "siège ville", "city"],
    }
    for field, patterns in guesses.items():
        for col, h_lower in header_lower.items():
            if h_lower in patterns:
                auto_map[field] = col
                break

    all_fields = {**REQUIRED_FIELDS, **OPTIONAL_FIELDS}
    return render_template("mapping.html",
                           headers=headers, fields=all_fields,
                           required=set(REQUIRED_FIELDS.keys()),
                           auto_map=auto_map, file_id=file_id,
                           email=email, filename=f.filename,
                           notion_ok=notion_configured(),
                           mail_ok=mail_configured())


@app.route("/generate", methods=["POST"])
def generate_route():
    file_id = request.form.get("file_id", "")
    email = request.form.get("email", "").strip()

    tmp_path = _temp_files.pop(file_id, None)
    if not tmp_path or not os.path.exists(tmp_path):
        flash("Fichier expiré — re-upload nécessaire", "error")
        return redirect(url_for("index"))

    with open(tmp_path, "rb") as f:
        source_bytes = f.read()
    os.unlink(tmp_path)

    # Mapping
    mapping = {}
    for field in list(REQUIRED_FIELDS) + list(OPTIONAL_FIELDS):
        val = request.form.get(f"map_{field}", "")
        if val:
            mapping[field] = int(val)

    missing = [REQUIRED_FIELDS[f] for f in REQUIRED_FIELDS if f not in mapping]
    if missing:
        flash(f"Champs requis manquants : {', '.join(missing)}", "error")
        return redirect(url_for("index"))

    today = date.today()

    log.info("Génération Excel...")
    xlsx_bytes, cat_rows = generate(source_bytes, mapping, today)
    del source_bytes
    gc.collect()

    counts = {k: len(v) for k, v in cat_rows.items()}
    total = sum(counts.values())
    log.info(f"  OK : {total} sociétés ({counts})")

    # Mail
    filename = f"Fins_de_bail_triennal_{today.strftime('%Y-%m-%d')}.xlsx"
    if email and mail_configured():
        log.info(f"Envoi mail à {email}...")
        try:
            send_mail(email, xlsx_bytes, filename, counts)
        except Exception as e:
            log.error(f"Mail erreur : {e}")

    # Notion en arrière-plan
    if notion_configured():
        cat_rows_copy = {k: list(v) for k, v in cat_rows.items()}
        def bg_sync():
            try:
                gc.collect()
                log.info("Sync Notion (arrière-plan)...")
                stats = sync_to_notion(cat_rows_copy)
                log.info(f"  Notion terminé : {stats}")
            except Exception as e:
                log.error(f"  Notion erreur : {e}")
        threading.Thread(target=bg_sync, daemon=True).start()

    # Téléchargement immédiat
    return send_file(
        BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


# ─────────────────────────────────────────────
#  Auto-refresh mensuel (Dropbox → recalcul)
# ─────────────────────────────────────────────

@app.route("/auto", methods=["GET", "POST"])
def auto_refresh():
    """
    Process mensuel automatique :
    1. Télécharge les fichiers générés depuis Dropbox
    2. Extrait les données, recalcule les mois restants
    3. Re-catégorise (< 6, 6-9, 9-12 mois)
    4. Génère un nouveau fichier Excel
    5. Sync Notion
    6. Envoie par mail
    7. Upload le résultat sur Dropbox
    """
    # Auth par header ou query param
    secret = request.headers.get("X-Sync-Secret", "") or request.args.get("secret", "")
    if not AUTO_SECRET or secret != AUTO_SECRET:
        return jsonify({"error": "Unauthorized"}), 401

    if not dropbox_configured():
        return jsonify({"error": "Dropbox non configuré"}), 500

    today = date.today()
    log.info(f"=== AUTO-REFRESH {today.strftime('%d/%m/%Y')} ===")

    # 1. Télécharger les fichiers depuis Dropbox
    try:
        files = fetch_xlsx_files()
    except Exception as e:
        log.error(f"Dropbox fetch échoué : {e}")
        return jsonify({"error": f"Dropbox fetch: {e}"}), 500

    if not files:
        log.warning("Aucun fichier .xlsx sur Dropbox")
        return jsonify({"error": "Aucun fichier .xlsx dans /FINS DE BAIL"}), 404

    log.info(f"  {len(files)} fichier(s) récupéré(s) depuis Dropbox")

    # 2. Extraire et recalculer
    try:
        rows = extract_from_generated(files, today)
    except Exception as e:
        log.error(f"Extraction échouée : {e}")
        return jsonify({"error": f"Extraction: {e}"}), 500

    if not rows:
        log.warning("Aucune donnée exploitable dans les fichiers")
        return jsonify({"error": "Aucune donnée exploitable"}), 404

    # 3. Catégoriser
    cat_rows = categorize(rows)
    counts = {k: len(v) for k, v in cat_rows.items()}
    total = sum(counts.values())
    log.info(f"  {total} sociétés réparties : {counts}")

    # 4. Générer le nouveau fichier Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet_name, data_list in cat_rows.items():
        _write_call_sheet(wb, sheet_name, data_list, today)
    _write_readme(wb, today, counts)

    buf = BytesIO()
    wb.save(buf)
    wb.close()
    xlsx_bytes = buf.getvalue()
    del buf
    gc.collect()

    filename = f"Fins_de_bail_triennal_{today.strftime('%Y-%m-%d')}.xlsx"
    log.info(f"  Excel généré : {filename} ({len(xlsx_bytes)} octets)")

    # 5. Sync Notion (synchrone ici, pas de téléchargement à renvoyer)
    notion_stats = {"skipped": True}
    if notion_configured():
        try:
            log.info("  Sync Notion...")
            notion_stats = sync_to_notion(cat_rows)
            log.info(f"  Notion : {notion_stats}")
        except Exception as e:
            log.error(f"  Notion erreur : {e}")
            notion_stats = {"error": str(e)}

    # 6. Envoyer par mail
    mail_ok = False
    if mail_configured() and AUTO_EMAIL:
        try:
            log.info(f"  Mail → {AUTO_EMAIL}")
            send_mail(AUTO_EMAIL, xlsx_bytes, filename, counts)
            mail_ok = True
        except Exception as e:
            log.error(f"  Mail erreur : {e}")

    # 7. Upload sur Dropbox (écrase l'ancien)
    dropbox_ok = False
    try:
        log.info(f"  Upload Dropbox → {filename}")
        dropbox_upload(filename, xlsx_bytes)
        dropbox_ok = True
    except Exception as e:
        log.error(f"  Dropbox upload erreur : {e}")

    result = {
        "status": "ok",
        "date": today.strftime("%Y-%m-%d"),
        "total": total,
        "counts": counts,
        "notion": notion_stats,
        "mail": mail_ok,
        "dropbox_upload": dropbox_ok,
    }
    log.info(f"=== AUTO-REFRESH TERMINÉ === {result}")
    return jsonify(result)


# ─────────────────────────────────────────────
#  Health check
# ─────────────────────────────────────────────

@app.route("/health")
def health():
    return jsonify({
        "status": "ok",
        "notion": notion_configured(),
        "mail": mail_configured(),
        "dropbox": dropbox_configured(),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
