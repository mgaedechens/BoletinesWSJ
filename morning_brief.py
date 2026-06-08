#!/usr/bin/env python3
"""Compilador diario de boletines WSJ: Gmail IMAP → Gmail SMTP."""

import email
import imaplib
import os
import smtplib
from datetime import date, datetime, timezone
from email.header import decode_header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from bs4 import BeautifulSoup
from dotenv import load_dotenv

load_dotenv()

GMAIL_USER = "matiasgaedechens1@gmail.com"
GMAIL_APP_PASSWORD = os.environ["GMAIL_APP_PASSWORD"]

IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Todos los newsletters de WSJ vienen de este remitente exacto.
# Hotmail los reenvía a Gmail mediante la regla configurada en Outlook.
WSJ_SENDER = "access@interactive.wsj.com"


# ---------------------------------------------------------------------------
# Helpers de decodificación
# ---------------------------------------------------------------------------

def decode_str(s: str | None) -> str:
    if not s:
        return ""
    parts = decode_header(s)
    result = []
    for part, charset in parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            result.append(str(part))
    return "".join(result)


def extract_bodies(msg) -> tuple[str | None, str | None]:
    html_body = text_body = None

    if msg.is_multipart():
        for part in msg.walk():
            if "attachment" in str(part.get("Content-Disposition", "")):
                continue
            ct = part.get_content_type()
            charset = part.get_content_charset() or "utf-8"
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            decoded = payload.decode(charset, errors="replace")
            if ct == "text/html" and html_body is None:
                html_body = decoded
            elif ct == "text/plain" and text_body is None:
                text_body = decoded
    else:
        ct = msg.get_content_type()
        charset = msg.get_content_charset() or "utf-8"
        payload = msg.get_payload(decode=True)
        if payload:
            decoded = payload.decode(charset, errors="replace")
            if ct == "text/html":
                html_body = decoded
            else:
                text_body = decoded

    return html_body, text_body


# ---------------------------------------------------------------------------
# IMAP: leer desde Gmail
# ---------------------------------------------------------------------------

def fetch_newsletters() -> list[dict]:
    today = date.today().strftime("%d-%b-%Y")  # e.g. 08-Jun-2026

    conn = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    conn.login(GMAIL_USER, GMAIL_APP_PASSWORD)

    # Intentar varias carpetas según idioma de Gmail (español vs inglés)
    selected = False
    for folder in ('"[Gmail]/All Mail"', '"[Gmail]/Todos"', "INBOX"):
        status, _ = conn.select(folder)
        if status == "OK":
            selected = True
            print(f"  Carpeta: {folder}")
            break
    if not selected:
        raise RuntimeError("No se pudo seleccionar ninguna carpeta de Gmail")

    _, data = conn.search(None, f'ON "{today}" FROM "{WSJ_SENDER}"')
    ids = data[0].split() if data[0] else []

    print(f"  Emails de {WSJ_SENDER} en Gmail hoy: {len(ids)}")

    results = []
    for msg_id in ids:
        _, raw = conn.fetch(msg_id, "(RFC822)")
        msg = email.message_from_bytes(raw[0][1])
        subject = decode_str(msg.get("Subject", ""))
        sender = decode_str(msg.get("From", ""))
        html, text = extract_bodies(msg)
        results.append({"subject": subject, "sender": sender, "html": html, "text": text})

    conn.logout()
    return results


# ---------------------------------------------------------------------------
# Compilar HTML
# ---------------------------------------------------------------------------

def extract_body_content(full_html: str) -> str:
    soup = BeautifulSoup(full_html, "lxml")
    styles = "".join(str(tag) for tag in soup.find_all("style"))
    body = soup.find("body")
    content = str(body) if body else full_html
    return styles + content


def build_compiled_html(newsletters: list[dict]) -> str:
    today_str = date.today().strftime("%d/%m/%Y")
    count = len(newsletters)
    plural = "es" if count != 1 else ""
    now_utc = datetime.now(timezone.utc).strftime("%H:%M UTC")

    sections_html = ""
    for nl in newsletters:
        if nl["html"]:
            content = extract_body_content(nl["html"])
        elif nl["text"]:
            escaped = (
                nl["text"]
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
            )
            content = f'<pre style="white-space:pre-wrap;font-family:inherit;font-size:14px;line-height:1.6;">{escaped}</pre>'
        else:
            content = '<p style="color:#94a3b8;">Sin contenido disponible.</p>'

        sections_html += f"""
    <div style="margin-bottom:32px;border:1px solid #dde1e7;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.06);">
      <div style="background:#0f172a;padding:18px 24px;">
        <h2 style="margin:0;color:#f1f5f9;font-size:16px;font-weight:600;line-height:1.4;">{nl['subject']}</h2>
        <p style="margin:6px 0 0;color:#94a3b8;font-size:12px;">De: {nl['sender']}</p>
      </div>
      <div style="padding:24px;background:#ffffff;font-size:14px;line-height:1.6;color:#1e293b;overflow-x:auto;">
        {content}
      </div>
    </div>
"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Boletines del día — {today_str}</title>
</head>
<body style="margin:0;padding:24px 16px;background:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
  <div style="max-width:800px;margin:0 auto;">

    <div style="background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);padding:32px 28px;border-radius:12px;margin-bottom:28px;">
      <h1 style="margin:0;color:#f1f5f9;font-size:24px;font-weight:700;">📰 Boletines del día</h1>
      <p style="margin:10px 0 0;color:#94a3b8;font-size:14px;">
        {today_str} &mdash; {count} boletín{plural} encontrado{plural}
      </p>
    </div>

    {sections_html}

    <p style="text-align:center;color:#94a3b8;font-size:11px;margin-top:8px;padding-bottom:8px;">
      Compilado automáticamente · {now_utc}
    </p>
  </div>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Gmail SMTP
# ---------------------------------------------------------------------------

def send_email(subject: str, html: str) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"Morning Brief <{GMAIL_USER}>"
    msg["To"] = GMAIL_USER
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as s:
        s.ehlo()
        s.starttls()
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.sendmail(GMAIL_USER, GMAIL_USER, msg.as_string())


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    print(f"[{datetime.now():%Y-%m-%d %H:%M}] Buscando boletines WSJ en Gmail…")

    newsletters = fetch_newsletters()

    if not newsletters:
        print("Sin boletines WSJ hoy. No se envía email.")
        return

    print(f"  Encontrados: {len(newsletters)}")
    for nl in newsletters:
        print(f"    · {nl['subject']}")

    today_str = date.today().strftime("%d/%m/%Y")
    subject = f"📰 Boletines del día — {today_str}"
    html = build_compiled_html(newsletters)

    print(f"  Enviando compilado a {GMAIL_USER}…")
    send_email(subject, html)
    print("✓ Email enviado correctamente.")


if __name__ == "__main__":
    main()
