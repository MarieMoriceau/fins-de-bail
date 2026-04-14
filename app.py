"""Mini-app Flask — génère le fichier d'appel sur demande.

Usage local :
    pip install -r requirements.txt
    python app.py
    # Ouvre http://localhost:5000

Déploiement Render : voir README.md
"""
import os
from datetime import datetime
from io import BytesIO

from flask import Flask, render_template, request, send_file, abort

from generator import generate

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 Mo max


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/generate")
def do_generate():
    f = request.files.get("source")
    if not f or not f.filename.lower().endswith(".xlsx"):
        abort(400, "Merci d'envoyer un fichier .xlsx.")

    try:
        out_bytes = generate(f.read())
    except Exception as e:
        abort(500, f"Erreur lors du traitement : {e}")

    today = datetime.now().strftime("%Y-%m-%d")
    filename = f"Fins_de_bail_triennal_{today}.xlsx"

    return send_file(
        BytesIO(out_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.get("/healthz")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
