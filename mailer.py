"""mailer.py — Envoi SMTP du fichier généré."""

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


def send_file(to_email, xlsx_bytes, filename, counts):
    if not is_configured():
        return False
    total = sum(counts.values())
    msg = MIMEMultipart()
    msg["From"] = SMTP_FROM
    msg["To"] = to_email
    msg["Subject"] = f"Fins de bail — {total} sociétés détectées"
    body = f"""Bonjour,

Voici le fichier de suivi des fins de bail triennales.

  • Fin < 6 mois : {counts.get('Fin < 6 mois', 0)} sociétés
  • Fin 6-9 mois : {counts.get('Fin 6-9 mois', 0)} sociétés
  • Fin 9-12 mois : {counts.get('Fin 9-12 mois', 0)} sociétés

— Générateur Fins de Bail (Render)
"""
    msg.attach(MIMEText(body, "plain", "utf-8"))
    part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    part.set_payload(xlsx_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{filename}"')
    msg.attach(part)
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as s:
            s.starttls()
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
        log.info(f"Mail envoyé à {to_email}")
        return True
    except Exception as e:
        log.error(f"Échec mail → {e}")
        return False
