"""Versand der Rechnung per E-Mail (SMTP, z.B. Gmail mit App-Passwort)."""
from __future__ import annotations

import smtplib
from email.message import EmailMessage


def senden(cfg: dict, an: list[str], betreff: str, text: str, anhaenge: list[tuple[str, bytes, str]]) -> None:
    msg = EmailMessage()
    msg["From"] = cfg.get("absender", cfg["user"])
    msg["To"] = ", ".join(an)
    msg["Subject"] = betreff
    msg.set_content(text)
    for name, data, mime in anhaenge:
        maintype, subtype = mime.split("/", 1)
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)
    with smtplib.SMTP(cfg.get("host", "smtp.gmail.com"), int(cfg.get("port", 587))) as s:
        s.starttls()
        s.login(cfg["user"], cfg["password"])
        s.send_message(msg)
