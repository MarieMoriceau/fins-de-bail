"""
mailer.py — Envoi du fichier généré par SMTP.

Variables d'environnement :
  SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM
Ne fait rien si non configuré.
"""

import os
import logging
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

log = logging.getLogger("mailer")

SMTP_HOST = os.environ.get("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASS = os.environ.get("SMTP_PASS")
SMTP_FROM = os.environ.get("SMTP_FROM")


def is_configured() -> bool:
    return all([SMTP_HOST, SMTP_USER, SMTP_PASS, SMTP_FROM])


def send_file(to_email: str, xlsx_bytes: bytes, filename: str, counts: dict) -> bool:
    """Envoie le fichier xlsx par mail. Retourne True si succès."""
    if not is_configured():
        log.warning("SMTP non configuré — mail non envoyé")
        return False

    total = sum(counts.values())

    msg = MIMEMultipart()
    msg["From"] = SMTP_FROM
    msg["To"] = to_email
    msg["Subject"] = f"Fins de bail — {total} sociétés détectées"

    body = f"""Bonjour,

Voici le fichier de suivi des fins de bail triennales.

Résumé :
  • Fin < 6 mois : {counts.get('Fin < 6 mois', 0)} sociétés (priorité haute)
  • Fin 6-9 mois : {counts.get('Fin 6-9 mois', 0)} sociétés (priorité moyenne)
  • Fin 9-12 mois : {counts.get('Fin 9-12 mois', 0)} sociétés (à anticiper)

Le fichier est en pièce jointe.

— Générateur Fins de Bail (Render)
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))

    part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    part.set_payload(xlsx_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)
        log.info(f"Mail envoyé à {to_email}")
        return True
    except Exception as e:
        log.error(f"Échec envoi mail à {to_email} : {e}")
        return False
