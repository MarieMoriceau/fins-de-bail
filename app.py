"""
app.py — Générateur fins de bail tout-en-un.

Flow :
  1. GET  /           → page d'accueil (upload + email + lien template)
  2. POST /upload     → lit les headers du fichier, affiche la page de mapping
  3. POST /generate   → génère Excel + sync Notion + envoi mail + téléchargement
  4. GET  /template   → télécharge le fichier template vierge
"""

import json
import logging
import os
import tempfile
from datetime import date
from io import BytesIO

from flask import Flask, render_template, request, send_file, redirect, url_for, flash

from generator import generate, read_headers, REQUIRED_FIELDS, OPTIONAL_FIELDS, DEFAULT_MAPPING
from notion_sync import sync_to_notion, is_configured as notion_configured
from mailer import send_file as send_mail, is_configured as mail_configured

logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s")
log = logging.getLogger("app")

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-secret-fins-de-bail-2026")
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 Mo

# Stockage temporaire du fichier uploadé entre /upload et /generate
_temp_files = {}


@app.route("/")
def index():
    return render_template("index.html",
                           notion_ok=notion_configured(),
                           mail_ok=mail_configured())


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

    # Lire les headers
    try:
        headers = read_headers(source_bytes)
    except Exception as e:
        flash(f"Impossible de lire le fichier : {e}", "error")
        return redirect(url_for("index"))

    if not headers:
        flash("Le fichier semble vide (pas de headers en ligne 1)", "error")
        return redirect(url_for("index"))

    # Sauvegarder temporairement
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(source_bytes)
    tmp.close()
    file_id = os.path.basename(tmp.name)
    _temp_files[file_id] = tmp.name

    # Auto-matching : essayer de deviner le mapping
    auto_map = {}
    header_lower = {col: h.lower() for col, h in headers}
    guesses = {
        "id":      ["id", "identifiant", "id pipedrive", "id société"],
        "societe": ["société", "societe", "nom", "raison sociale", "company"],
        "date":    ["date adresse retenue", "date d'entrée", "date entrée", "date début bail",
                    "📅 date adresse retenue", "date adresse"],
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
                           headers=headers,
                           fields=all_fields,
                           required=set(REQUIRED_FIELDS.keys()),
                           auto_map=auto_map,
                           file_id=file_id,
                           email=email,
                           filename=f.filename,
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

    # Construire le mapping depuis le formulaire
    mapping = {}
    for field in list(REQUIRED_FIELDS) + list(OPTIONAL_FIELDS):
        val = request.form.get(f"map_{field}", "")
        if val:
            mapping[field] = int(val)

    # Vérifier les champs requis
    missing = [REQUIRED_FIELDS[f] for f in REQUIRED_FIELDS if f not in mapping]
    if missing:
        flash(f"Champs requis manquants : {', '.join(missing)}", "error")
        return redirect(url_for("index"))

    today = date.today()

    # 1. Génération Excel
    log.info("Génération Excel...")
    xlsx_bytes, cat_rows = generate(source_bytes, mapping, today)
    counts = {k: len(v) for k, v in cat_rows.items()}
    total = sum(counts.values())
    log.info(f"  OK : {total} sociétés ({counts})")

    # 2. Sync Notion
    if notion_configured():
        log.info("Sync Notion...")
        try:
            stats = sync_to_notion(cat_rows)
            log.info(f"  Notion : {stats}")
        except Exception as e:
            log.error(f"  Notion erreur : {e}")

    # 3. Envoi mail
    filename = f"Fins_de_bail_triennal_{today.strftime('%Y-%m-%d')}.xlsx"
    if email and mail_configured():
        log.info(f"Envoi mail à {email}...")
        send_mail(email, xlsx_bytes, filename, counts)

    # 4. Téléchargement
    return send_file(
        BytesIO(xlsx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
