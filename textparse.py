"""Schnellerfassung: einfache Textzeilen → Rechnungszeilen (ohne AI, regelbasiert).

Beispiele (eine Leistung pro Zeile, «-» = Unterpunkt der vorherigen Zeile):
    11 Stk Innentüren à 525
    Türzargen 11 Stk à 125
    Transport pauschal 150
    Montage 3 Std à 85
    - Montage und Ausrichtung der Zargen
    - Einhängen der Türblätter
    Rinnenverbinder; 2; Stk; 18.50          (auch mit ; oder | getrennt)
"""
from __future__ import annotations

import re

EINHEIT_MAP = {
    "stk": "Stk.", "stück": "Stk.", "stueck": "Stk.", "st": "Stk.", "pce": "Stk.",
    "std": "Std.", "h": "Std.", "stunde": "Std.", "stunden": "Std.",
    "m": "m", "lm": "m", "meter": "m", "laufmeter": "m",
    "m2": "m²", "m²": "m²", "qm": "m²", "m3": "m³", "m³": "m³",
    "kg": "kg", "fahrt": "Fahrt", "fahrten": "Fahrt",
    "pauschal": "pauschal", "psch": "pauschal", "pau": "pauschal", "pausch": "pauschal",
}
_EINH = r"(stk|stück|stueck|st|pce|std|h|stunden?|laufmeter|meter|lm|m2|m²|qm|m3|m³|m|kg|fahrten|fahrt)\.?"
_ZAHL = r"\d+(?:['’]\d{3})*(?:[.,]\d+)?"
MANGEL_RE = re.compile(r"\b(\d{3}(?:\.\d)?(?:\.[abc])?-\d{2})\b")


def zahl(s: str) -> float:
    s = s.replace("'", "").replace("’", "").replace(",", ".")
    return float(s)


def _sauber(t: str) -> str:
    t = re.sub(r"\s{2,}", " ", t)
    t = re.sub(r"\s+([,.;:])", r"\1", t)
    return t.strip(" ,;:-–·").strip()


def zeile_parsen(roh: str) -> dict:
    z = {"text": "", "menge": 1.0, "einheit": "pauschal", "ansatz": None, "bezug": "", "unter": []}
    line = roh.strip()

    m = MANGEL_RE.search(line)
    if m:
        z["bezug"] = m.group(1)
        line = (line[:m.start()] + line[m.end():]).strip()
        line = re.sub(r"\(\s*\)|\[\s*\]", "", line)

    # --- getrennte Spalten: Text ; Menge ; Einheit ; Preis
    if re.search(r"[;|\t]", line):
        teile = [p.strip() for p in re.split(r"[;|\t]", line) if p.strip()]
        z["text"] = teile[0]
        rest = teile[1:]
        zahlen = [p for p in rest if re.fullmatch(_ZAHL + r"(?:\.-|\.–)?", p.replace("CHF", "").strip())]
        einh = [p for p in rest if p.lower().rstrip(".") in EINHEIT_MAP]
        if einh:
            z["einheit"] = EINHEIT_MAP[einh[0].lower().rstrip(".")]
        nums = [zahl(re.sub(r"[^\d'’.,]", "", p).rstrip(".")) for p in zahlen]
        if len(nums) >= 2:
            z["menge"], z["ansatz"] = nums[0], nums[1]
        elif len(nums) == 1:
            z["ansatz"] = nums[0]
        return z

    work = " " + line + " "
    low = work.lower()

    # --- Preis: «à 525», «a 525», «@ 525», «je 525», «zu 525», «x 525», «CHF 525», «525.-», «Fr. 525»
    preis = None
    for pat in (rf"(?:\bà|\ba|@|\bje|\bzu|\bx)\s*(?:chf|fr\.?)?\s*({_ZAHL})",
                rf"(?:chf|fr\.)\s*({_ZAHL})",
                rf"({_ZAHL})\s*(?:\.-|\.–|chf|fr\.?)(?=\s|$)"):
        mm = re.search(pat, low)
        if mm:
            preis = zahl(mm.group(1).rstrip("."))
            work = work[:mm.start()] + " " + work[mm.end():]
            low = work.lower()
            break

    # --- Menge + Einheit: «11 Stk», «3 Std», «2.5 m2»
    mm = re.search(rf"(?<![\d'’.,])({_ZAHL})\s*{_EINH}(?=\s|$|[,;])", low)
    if mm:
        z["menge"] = zahl(mm.group(1))
        z["einheit"] = EINHEIT_MAP[mm.group(2).lower()]
        work = work[:mm.start()] + " " + work[mm.end():]
        low = work.lower()

    # --- «pauschal»
    mp = re.search(r"\b(pauschal|psch|pausch)\.?\b", low)
    if mp:
        z["einheit"], z["menge"] = "pauschal", 1.0 if not mm else z["menge"]
        work = work[:mp.start()] + " " + work[mp.end():]
        low = work.lower()

    # --- kein «à»-Preis, aber eine einzelne Zahl am Ende → Preis
    if preis is None:
        mz = re.search(rf"({_ZAHL})\s*$", work.strip())
        if mz and not re.fullmatch(_ZAHL, work.strip()):
            preis = zahl(mz.group(1))
            work = work.strip()[:mz.start()]

    z["ansatz"] = preis
    z["text"] = _sauber(work)
    return z


def text_parsen(text: str) -> list[dict]:
    zeilen: list[dict] = []
    for roh in text.splitlines():
        if not roh.strip():
            continue
        s = roh.strip()
        if re.match(r"^[-–•*]\s*", s):
            unter = re.sub(r"^[-–•*]\s*", "", s).strip()
            if zeilen and unter:
                zeilen[-1]["unter"].append(unter)
            elif unter:
                zeilen.append({"text": unter, "menge": 1.0, "einheit": "pauschal", "ansatz": None,
                               "bezug": "", "unter": []})
            continue
        zeilen.append(zeile_parsen(s))
    return zeilen


if __name__ == "__main__":
    import json
    print(json.dumps(text_parsen(__doc__.split("Beispiele")[1].split('"""')[0].split("\n", 1)[1]
                                 .replace("(auch mit ; oder | getrennt)", "")), ensure_ascii=False, indent=1))


# ------------------------------------------------------------------ Nur Leistungstexte (z.B. aus ChatGPT kopiert)
_LISTE = re.compile(r"^(\s*)(?:[-–•*●▪◦]|\d{1,2}[.)]|[a-zA-Z][.)])\s+(.*)$")


def _md_weg(t: str) -> str:
    t = re.sub(r"^#{1,6}\s*", "", t.strip())
    t = re.sub(r"\*\*|__|`", "", t)
    t = re.sub(r"(?<!\w)\*(?!\s)|(?<!\s)\*(?!\w)", "", t)
    t = re.sub(r"^(?:schritt|step)\s*\d+\s*[:.)-]\s*", "", t, flags=re.I)
    t = re.sub(r"[\U0001F300-\U0001FAFF☀-➿]", "", t)     # Emojis
    return re.sub(r"\s{2,}", " ", t).strip(" :–-")


def leistungen_parsen(text: str) -> list[dict]:
    """Jeder Listenpunkt = eine Zeile (nur Text, Preis später). Eingerückte Punkte = Unterpunkte.
    Enthält der Text Listenpunkte, werden Einleitungs-/Schlusssätze und Überschriften ignoriert."""
    lines = [l.rstrip() for l in text.replace("\t", "    ").splitlines() if l.strip()]
    hat_liste = any(_LISTE.match(l) for l in lines)
    out: list[dict] = []
    basis = None
    for l in lines:
        m = _LISTE.match(l)
        if m:
            indent, inhalt = len(m.group(1)), _md_weg(m.group(2))
            if not inhalt:
                continue
            if basis is None:
                basis = indent
            if indent > basis and out:
                out[-1]["unter"].append(inhalt)
                continue
            out.append({"text": inhalt, "menge": 1.0, "einheit": "pauschal", "ansatz": None,
                        "bezug": (MANGEL_RE.search(inhalt).group(1) if MANGEL_RE.search(inhalt) else ""), "unter": []})
        elif not hat_liste:
            inhalt = _md_weg(l)
            if inhalt:
                out.append({"text": inhalt, "menge": 1.0, "einheit": "pauschal", "ansatz": None,
                            "bezug": "", "unter": []})
        # sonst: Überschrift / Einleitung / Schlusssatz → ignorieren
    for z in out:                                     # Mangel-Nr. aus dem Text entfernen
        if z["bezug"]:
            z["text"] = _sauber(re.sub(r"\(?\b" + re.escape(z["bezug"]) + r"\b\)?", "", z["text"]))
    return out
