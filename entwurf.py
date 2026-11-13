"""Entwürfe speichern und wieder laden.

Jede Rechnung wird zusätzlich als JSON (Eingabedaten des Formulars) abgelegt.
Für ältere Entwürfe ohne JSON werden die Daten aus der Excel-Rechnung zurückgelesen.
"""
from __future__ import annotations

import io
import json
import re
from datetime import date, datetime


def _d(s: str | None) -> date | None:
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            pass
    return None


def zu_json(daten: dict) -> bytes:
    return json.dumps(daten, ensure_ascii=False, indent=1, default=str).encode("utf-8")


def aus_json(raw: bytes) -> dict:
    d = json.loads(raw.decode("utf-8"))
    for k in ("datum", "von", "bis"):
        d[k] = _d(d.get(k)) or date.today()
    return d


def aus_excel(raw: bytes) -> dict:
    """Liest eine von der App erzeugte Rechnungs-Excel zurück in Formulardaten."""
    from openpyxl import load_workbook

    ws = load_workbook(io.BytesIO(raw), data_only=True).active
    v = lambda ref: ws[ref].value
    kopf = {str(ws.cell(r, 1).value or "").strip(": "): ws.cell(r, 2).value for r in range(13, 23)}
    aus = str(kopf.get("Ausführung") or "")
    teile = [t.strip() for t in re.split(r"[–-]", aus) if t.strip()]
    datum = _d(str(kopf.get("Rechnungsdatum") or "")) or date.today()
    von = _d(teile[0]) if teile else datum
    bis = _d(teile[-1]) if teile else von

    gruppen: list[dict] = []
    row = 24
    leer = 0
    while row < ws.max_row + 1 and leer < 2:
        a = ws.cell(row, 1).value
        b = ws.cell(row, 2).value
        row += 1
        a = "" if a is None else str(a).strip()
        b = "" if b is None else str(b).strip()
        if not a and not b:
            leer += 1
            continue
        leer = 0
        if a and re.fullmatch(r"\d+", a):                       # Positionstitel
            m = re.search(r"\s*\((?:Mängelliste Pos\. ([^)]+)|Zusatzleistung)\)\s*$", b)
            gruppen.append({"bkp": (m.group(1).strip() if m and m.group(1) else ""),
                            "titel": b[:m.start()].strip() if m else b, "zeilen": []})
        elif a and gruppen:                                      # Zeile 1.1, 1.2 …
            menge, einheit, ansatz = ws.cell(row - 1, 3).value, ws.cell(row - 1, 4).value, ws.cell(row - 1, 5).value
            gruppen[-1]["zeilen"].append({"text": b, "menge": float(menge) if menge is not None else 1.0,
                                          "einheit": str(einheit or "pauschal"),
                                          "ansatz": float(ansatz) if ansatz is not None else None,
                                          "bezug": "", "unter": ""})
        elif gruppen and gruppen[-1]["zeilen"] and b.startswith("Bezug Mängelliste:"):
            gruppen[-1]["zeilen"][-1]["bezug"] = b.split(":", 1)[1].strip()
        elif gruppen and gruppen[-1]["zeilen"] and b.startswith("–"):
            z = gruppen[-1]["zeilen"][-1]
            z["unter"] = (z["unter"] + "\n" if z["unter"] else "") + b.lstrip("– ").strip()
        # «Total Pos. …» und Sonstiges ignorieren

    return {"version": 1, "nr": str(kopf.get("Rechnungs-Nr.") or ""), "datum": datum, "von": von, "bis": bis,
            "empfaenger": {"name": v("D8") or "", "strasse": v("D9") or "", "plz_ort": v("D10") or ""},
            "gruppen": gruppen}
