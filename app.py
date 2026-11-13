"""Rechnungs-App Hausbau Gerlafingen – Eingabe per Handy, Ausgabe als Excel + PDF."""
from __future__ import annotations

import json
import uuid
from datetime import date, datetime, timedelta
from pathlib import Path

import streamlit as st

from invoice import Gruppe, Rechnung, Zeile, als_excel, als_pdf, chf, dateiname, dstr

st.set_page_config(page_title="Rechnung erfassen", page_icon="🧾", layout="centered")

NAVY = "#1F3A5F"
EINHEITEN = ["Stk.", "Std.", "pauschal", "m", "m²", "m³", "kg", "Fahrt"]
VORLAGEN = {
    "Material": ("Lieferung Material: ", 1.0, "Stk."),
    "Arbeit": ("Arbeit: ", 1.0, "Std."),
    "Lieferung": ("Transportkosten", 1.0, "pauschal"),
    "Pauschale": ("", 1.0, "pauschal"),
    "Fahrspesen": ("Fahrspesen", 1.0, "Fahrt"),
}
ZUSATZ = "➕ Zusatzleistung (nicht in der Mängelliste)"

_HERE = Path(__file__).parent
_JSON = next((p for p in (_HERE / "data" / "maengelliste.json", _HERE / "maengelliste.json") if p.exists()), None)
if _JSON is None:
    st.error("Datei «maengelliste.json» fehlt im Repo. Bitte direkt neben app.py hochladen.")
    st.stop()
DATA = json.loads(_JSON.read_text(encoding="utf-8"))
POS = {p["pos"]: p for p in DATA["positionen"]}
S = st.secrets
ss = st.session_state

# ------------------------------------------------------------------ Design
st.markdown(f"""
<style>
#MainMenu, footer {{visibility: hidden;}}
.block-container {{padding-top: 1.2rem; padding-bottom: 4rem; max-width: 760px;}}
.hero {{background: linear-gradient(135deg, {NAVY} 0%, #2E5C8A 100%); color: #fff;
        border-radius: 16px; padding: 18px 20px; margin-bottom: 14px;}}
.hero h1 {{color: #fff; font-size: 1.55rem; margin: 0 0 2px 0; padding: 0;}}
.hero p {{color: #dbe6f3; margin: 0; font-size: .9rem;}}
.step {{display: flex; align-items: center; gap: 10px; margin: 4px 0 8px 0;}}
.step .n {{background: {NAVY}; color: #fff; border-radius: 50%; width: 28px; height: 28px;
           display: inline-flex; align-items: center; justify-content: center; font-weight: 700; flex: none;}}
.step .t {{font-weight: 700; font-size: 1.1rem;}}
.total {{background: #EEF3F9; border: 1px solid #C9D6E6; border-radius: 14px; padding: 14px 18px; margin: 8px 0;}}
.total .lbl {{color: #4a5b70; font-size: .85rem;}}
.total .val {{color: {NAVY}; font-size: 1.8rem; font-weight: 800;}}
.total .sub {{color: #4a5b70; font-size: .85rem;}}
.chip {{display: inline-block; background: #EEF3F9; color: {NAVY}; border-radius: 999px;
        padding: 2px 10px; font-size: .8rem; margin-right: 6px;}}
div[data-testid="stBaseButton-primary"] button, button[kind="primary"],
button[data-testid="stBaseButton-primary"] {{background: {NAVY}; border-color: {NAVY};}}
div[data-testid="stExpander"] details {{border-radius: 12px;}}
.st-key-rtab [data-testid="stHorizontalBlock"] {{flex-wrap: nowrap !important; gap: .5rem !important;
        border-bottom: 1px solid #E3EAF3; padding: 6px 0;}}
.st-key-rtab [data-testid="stColumn"] {{min-width: 0 !important;}}
@media (max-width: 640px) {{
  .st-key-rtab [data-testid="stHorizontalBlock"] {{gap: .3rem !important;}}
  .st-key-rtab [data-testid="stColumn"] {{width: auto !important; flex: 0.9 1 0 !important;}}
  .st-key-rtab [data-testid="stColumn"]:nth-child(1) {{flex: 3 1 0 !important;}}
  .st-key-rtab [data-testid="stColumn"]:nth-child(2) {{flex: 1.5 1 0 !important;}}
  .st-key-rtab [data-testid="stColumn"]:nth-child(5) {{flex: 1.7 1 0 !important;}}
}}
.st-key-rtab [data-testid="stMarkdownContainer"] p, .st-key-rtab [data-testid="stMarkdown"] {{margin: 0 !important;}}
.st-key-rtab [data-testid="stMarkdownContainer"] {{margin-bottom: 0 !important;}}
.st-key-rtab [data-testid="stCheckbox"] {{margin: 0 !important; min-height: 0; display: flex; justify-content: center;}}
.st-key-rtab [data-testid="stCheckbox"] label {{padding: 0 !important; margin: 0 auto !important; width: auto !important;}}
.st-key-rtab [data-testid="stColumn"]:nth-child(3) [data-testid="stVerticalBlock"],
.st-key-rtab [data-testid="stColumn"]:nth-child(4) [data-testid="stVerticalBlock"] {{align-items: center !important;}}
.st-key-rtab [data-testid="stColumn"]:nth-child(3) [data-testid="stElementContainer"],
.st-key-rtab [data-testid="stColumn"]:nth-child(4) [data-testid="stElementContainer"] {{width: auto !important;}}
.st-key-rtab [data-testid="stColumn"]:nth-child(3) .rth, .st-key-rtab [data-testid="stColumn"]:nth-child(4) .rth {{width: 100%;}}
.st-key-rtab [data-testid="stBaseButton-tertiary"] p, .st-key-rtab button[kind="tertiary"] p {{
        white-space: normal !important; overflow: visible !important; text-overflow: clip !important;}}
.rth .sm {{display: none;}}
@media (max-width: 640px) {{
  .rth .lg {{display: none;}} .rth .sm {{display: inline;}}
  .rcell {{font-size: .82rem !important;}} .rsub {{font-size: .72rem !important;}}
  .st-key-rtab [data-testid="stBaseButton-tertiary"] {{font-size: .78rem !important;}}
}}
.st-key-rtab button[kind="tertiary"], .st-key-rtab [data-testid="stBaseButton-tertiary"] {{
        padding: 0 !important; min-height: 0 !important; color: {NAVY} !important; text-align: left;
        font-size: .85rem; text-decoration: underline; text-underline-offset: 3px; line-height: 1.25;}}
.rth {{font-size: .75rem; font-weight: 700; color: #4a5b70; text-transform: uppercase; letter-spacing: .03em;}}
.rth.r {{text-align: right;}}
.rth.c {{text-align: center;}}
.rcell {{font-size: .9rem; line-height: 1.35; padding: 2px 0;}}
.rsub {{font-size: .78rem; color: #4a5b70;}}
.rnum {{text-align: right; font-variant-numeric: tabular-nums; font-weight: 600;}}
.badge {{display: inline-block; border-radius: 999px; padding: 0 7px; font-size: .7rem; margin-left: 4px;}}
.badge.draft {{background: #FFF2CC; color: #7a5a00;}}
.badge.fin {{background: #DDF2E3; color: #1e6b37;}}
</style>
""", unsafe_allow_html=True)


def hero(sub: str) -> None:
    st.markdown(f'<div class="hero"><h1>🧾 Rechnung erfassen</h1><p>{sub}</p></div>', unsafe_allow_html=True)


def step(n: int | str, titel: str) -> None:
    st.markdown(f'<div class="step"><span class="n">{n}</span><span class="t">{titel}</span></div>',
                unsafe_allow_html=True)


# ------------------------------------------------------------------ Zugang
def login() -> bool:
    pw = S.get("APP_PASSWORT")
    if not pw or ss.get("ok"):
        return True
    hero("Firma Boniface · Hausbau Gerlafingen")
    with st.container(border=True):
        eingabe = st.text_input("Passwort", type="password")
        if st.button("Anmelden", type="primary", width="stretch"):
            if eingabe == pw:
                ss.ok = True
                st.rerun()
            st.error("Falsches Passwort.")
    return False


if not login():
    st.stop()


def github_status() -> str | None:
    """None = GitHub korrekt eingetragen, sonst Grund."""
    g = S.get("github")
    if not g:
        return "Abschnitt [github] fehlt in den Secrets."
    tok, repo = str(g.get("token", "")), str(g.get("repo", ""))
    if not tok or "HIER" in tok.upper() or tok.endswith("..."):
        return "In den Secrets steht beim token noch kein echter GitHub-Token (er beginnt mit github_pat_)."
    if not repo or "/" not in repo or "DEIN" in repo.upper() or "BENUTZERNAME" in repo.upper():
        return "In den Secrets steht bei repo noch der Platzhalter – Format: deinname/hausbau-rechnungen-daten."
    return None


@st.cache_resource
def store():
    if github_status():
        return None
    g = S["github"]
    from storage import GitHubStore
    return GitHubStore(str(g["token"]).strip(), str(g["repo"]).strip(), g.get("branch", "rechnungen"))


# ------------------------------------------------------------------ Zustand & Callbacks
ss.setdefault("gruppen", [])   # [{id, bkp, titel, zeilen:[zid], auto:{mangel_nr: zid}}]
if ss.pop("nr_neu", False):    # neue Nummer anfordern (wird unten vorgeschlagen)
    ss.pop("nr", None)
for _g in ss.gruppen:          # Sitzungen aus älteren App-Versionen nachrüsten
    _g.setdefault("auto", {})
    _g.setdefault("zeilen", [])


def neue_zeile(text="", menge=1.0, einheit="pauschal", ansatz=None, bezug="", unter="") -> str:
    zid = uuid.uuid4().hex[:8]
    ss[f"{zid}_text"], ss[f"{zid}_menge"], ss[f"{zid}_einheit"] = text, menge, einheit
    ss[f"{zid}_ansatz"], ss[f"{zid}_bezug"], ss[f"{zid}_unter"] = ansatz, bezug, unter
    return zid


def gruppe(gid: str) -> dict:
    return next(g for g in ss.gruppen if g["id"] == gid)


def cb_position_hinzufuegen() -> None:
    wahl = ss.get("pos_wahl")
    if not wahl or wahl == ZUSATZ:
        return
    bkp = wahl.split(" · ")[0]
    if any(g["bkp"] == bkp for g in ss.gruppen):
        ss.hinweis = f"Position {bkp} ist schon auf der Rechnung – unten ergänzen."
        return
    gid = uuid.uuid4().hex[:8]
    ss.gruppen.append({"id": gid, "bkp": bkp, "titel": POS[bkp]["titel"], "zeilen": [], "auto": {}})
    ss[f"{gid}_titel"] = POS[bkp]["titel"]
    ss.pos_wahl = None


def cb_zusatz_hinzufuegen() -> None:
    titel = (ss.get("zusatz_titel") or "").strip()
    if not titel:
        return
    gid = uuid.uuid4().hex[:8]
    ss.gruppen.append({"id": gid, "bkp": "", "titel": titel, "zeilen": [neue_zeile()], "auto": {}})
    ss[f"{gid}_titel"] = titel
    ss.zusatz_titel = ""
    ss.pos_wahl = None


def cb_mangel(gid: str, nr: str, text: str) -> None:
    """Checkbox Mängelpunkt: an = Zeile anlegen, aus = Zeile wieder entfernen."""
    g = gruppe(gid)
    if ss[f"{gid}_m_{nr}"]:
        zid = neue_zeile(text=text, bezug=nr)
        # in Reihenfolge der Mängelliste einsortieren
        reihenfolge = [m["nr"] for m in POS[g["bkp"]]["maengel"]]
        pos = len(g["zeilen"])
        for i, z in enumerate(g["zeilen"]):
            b = ss.get(f"{z}_bezug", "")
            if b in reihenfolge and reihenfolge.index(b) > reihenfolge.index(nr):
                pos = i
                break
        g["zeilen"].insert(pos, zid)
        g["auto"][nr] = zid
    else:
        zid = g["auto"].pop(nr, None)
        if zid in g["zeilen"]:
            g["zeilen"].remove(zid)


def cb_zeile_loeschen(gid: str, zid: str) -> None:
    g = gruppe(gid)
    if zid in g["zeilen"]:
        g["zeilen"].remove(zid)
    for nr, z in list(g["auto"].items()):
        if z == zid:
            g["auto"].pop(nr)
            ss[f"{gid}_m_{nr}"] = False


def cb_zeile_neu(gid: str, vorlage: str) -> None:
    txt, m, e = VORLAGEN[vorlage]
    gruppe(gid)["zeilen"].append(neue_zeile(txt, m, e))


def cb_schnell(gid: str) -> None:
    """Schnellerfassung: Text → Zeilen."""
    from textparse import leistungen_parsen, text_parsen
    g = gruppe(gid)
    gueltig = {m["nr"] for m in POS[g["bkp"]]["maengel"]} if g["bkp"] else set()
    parser = text_parsen if ss.get(f"{gid}_mitpreis") else leistungen_parsen
    neu = parser(ss.get(f"{gid}_schnell", ""))
    for z in neu:
        bezug = z["bezug"] if z["bezug"] in gueltig else ""
        g["zeilen"].append(neue_zeile(z["text"], z["menge"], z["einheit"], z["ansatz"], bezug, "\n".join(z["unter"])))
    ss[f"{gid}_schnell"] = ""
    ss.hinweis_schnell = (gid, len(neu))


def cb_gruppe_loeschen(gid: str) -> None:
    ss.gruppen = [g for g in ss.gruppen if g["id"] != gid]


BEHALTEN = {"ok", "jrows", "jrows_ts", "editor_v", "pdf_cache", "ansicht"}


def formular_leeren() -> None:
    for k in list(ss.keys()):
        if k not in BEHALTEN:
            del ss[k]


def formular_daten() -> dict:
    """Aktuelle Eingaben als speicherbares Dict (für späteres Weiterbearbeiten)."""
    return {"version": 1, "nr": ss.get("nr", ""), "datum": ss.get("datum", date.today()).isoformat(),
            "von": ss.get("von", date.today()).isoformat(), "bis": ss.get("bis", date.today()).isoformat(),
            "empfaenger": {"name": ss.get("empf_name", ""), "strasse": ss.get("empf_strasse", ""),
                           "plz_ort": ss.get("empf_plz_ort", "")},
            "gruppen": [{"bkp": g["bkp"], "titel": ss.get(f"{g['id']}_titel", g["titel"]),
                         "zeilen": [{"text": ss.get(f"{z}_text", ""), "menge": ss.get(f"{z}_menge"),
                                     "einheit": ss.get(f"{z}_einheit", ""), "ansatz": ss.get(f"{z}_ansatz"),
                                     "bezug": ss.get(f"{z}_bezug", ""), "unter": ss.get(f"{z}_unter", "")}
                                    for z in g["zeilen"]]} for g in ss.gruppen]}


def formular_fuellen(d: dict) -> None:
    formular_leeren()
    ss.nr = str(d.get("nr", ""))
    ss.datum, ss.von, ss.bis = d["datum"], d["von"], d["bis"]
    e = d.get("empfaenger") or {}
    ss.empf_name, ss.empf_strasse, ss.empf_plz_ort = e.get("name", ""), e.get("strasse", ""), e.get("plz_ort", "")
    ss.gruppen = []
    for g in d.get("gruppen", []):
        bkp = g.get("bkp", "") if g.get("bkp", "") in POS else ""
        gueltig = [m["nr"] for m in POS[bkp]["maengel"]] if bkp else []
        gid = uuid.uuid4().hex[:8]
        grp = {"id": gid, "bkp": bkp, "titel": g.get("titel", ""), "zeilen": [], "auto": {}}
        ss[f"{gid}_titel"] = g.get("titel", "")
        for z in g.get("zeilen", []):
            bezug = z.get("bezug", "") if z.get("bezug", "") in gueltig else ""
            menge = z.get("menge")
            zid = neue_zeile(z.get("text", ""), float(menge) if menge is not None else 1.0,
                             z.get("einheit") or "pauschal",
                             float(z["ansatz"]) if z.get("ansatz") is not None else None,
                             bezug, z.get("unter", "") or "")
            grp["zeilen"].append(zid)
            if bezug and bezug not in grp["auto"]:
                grp["auto"][bezug] = zid
        ss.gruppen.append(grp)
    ss.edit_nr = ss.nr


def cb_entwurf_laden(nr: str) -> None:
    from entwurf import aus_excel, aus_json
    rows, _ = journal()
    r0 = next((r for r in rows if r.get("rechnungs_nr") == nr), None)
    if not r0:
        ss.lade_msg = ("error", f"Rechnung {nr} nicht gefunden.")
        return
    if r0.get("final") == "ja":
        ss.lade_msg = ("error", f"Rechnung {nr} ist final und kann nicht mehr bearbeitet werden.")
        return
    try:
        raw = store().get(r0["daten"])[0] if r0.get("daten") else None
        d = aus_json(raw) if raw else aus_excel(store().get(r0["excel"])[0])
        d["nr"] = nr
        formular_fuellen(d)
        ss.lade_msg = ("success", f"Entwurf Nr. {nr} geladen – jetzt bearbeiten und mit «✅ Rechnung erstellen» speichern.")
        ss.ansicht = "rechnung"
    except Exception as ex:
        ss.lade_msg = ("error", f"Entwurf {nr} konnte nicht geladen werden: {ex}")


def cb_bearbeitung_abbrechen() -> None:
    formular_leeren()
    ss.nr_neu = True


def zeile_aus_state(zid: str) -> Zeile:
    unter = [u.strip(" -–•") for u in (ss.get(f"{zid}_unter") or "").splitlines() if u.strip(" -–•")]
    return Zeile(text=(ss.get(f"{zid}_text") or "").strip(), menge=ss.get(f"{zid}_menge"),
                 einheit=ss.get(f"{zid}_einheit") or "", ansatz=ss.get(f"{zid}_ansatz"),
                 bezug=ss.get(f"{zid}_bezug") or "", unterpunkte=unter)


def gruppen_obj() -> list[Gruppe]:
    return [Gruppe(titel=ss.get(f"{g['id']}_titel", g["titel"]), bkp=g["bkp"],
                   zeilen=[zeile_aus_state(z) for z in g["zeilen"]]) for g in ss.gruppen]


# ------------------------------------------------------------------ Journal-Hilfen
START = int(S.get("START_NR", 1024))


def jetzt() -> str:
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("Europe/Zurich")).strftime("%d.%m.%Y %H:%M")
    except Exception:
        return datetime.now().strftime("%d.%m.%Y %H:%M")


def journal(force: bool = False) -> tuple[list[dict], str | None]:
    """Journal aus GitHub, in der Sitzung zwischengespeichert (max. 60 s alt)."""
    if not store():
        return [], None
    alt = ss.get("jrows_ts", 0)
    if force or "jrows" not in ss or datetime.now().timestamp() - alt > 60:
        try:
            ss.jrows, ss.jrows_ts = store().journal()[0], datetime.now().timestamp()
        except Exception as e:
            alt_rows = ss.get("jrows")
            return (alt_rows if isinstance(alt_rows, list) else []), str(e)
    if not isinstance(ss.get("jrows"), list):
        ss.pop("jrows", None)
        return journal(force=True)
    return ss.jrows, None


def journal_setzen(rows) -> None:
    """Nach dem Schreiben: lokale Kopie sofort aktualisieren (kein Neuladen nötig)."""
    if isinstance(rows, list):
        ss.jrows, ss.jrows_ts = rows, datetime.now().timestamp()
    else:                                   # ältere storage.py liefert keine Liste → neu laden
        ss.pop("jrows", None)
        journal(force=True)
    ss.pop("journal_editor", None)
    ss.editor_v = ss.get("editor_v", 0) + 1


def naechste_nr() -> str:
    from storage import GitHubStore
    rows, _ = journal()
    return str(GitHubStore.naechste_nr_aus(rows, START))


# ------------------------------------------------------------------ Kopf + Anleitung
hero("Firma Boniface · Hausbau Gerlafingen · Grundlage: Mängelliste Architekt vom 04.07.2026")

# Eingaben behalten, auch wenn die Ansicht «Übersicht» offen ist (Streamlit löscht sonst nicht angezeigte Felder)
_FELDER = ("_text", "_menge", "_einheit", "_ansatz", "_bezug", "_unter", "_titel", "_schnell", "_mitpreis")
for _k in list(ss.keys()):
    if (_k in ("nr", "datum", "von", "bis", "empf_name", "empf_strasse", "empf_plz_ort", "zusatz_titel")
            or _k.endswith(_FELDER) or "_m_" in _k):
        ss[_k] = ss[_k]

rows, jerr = journal()
ss.setdefault("ansicht", "rechnung")
if ss.get("ansicht") is None:             # Segment abgewählt → zurück zur Rechnung
    ss.ansicht = "rechnung"
st.segmented_control(
    "Ansicht", ["rechnung", "uebersicht"], key="ansicht", label_visibility="collapsed", width="stretch",
    format_func=lambda o: (f"✏️ Entwurf {ss.edit_nr}" if ss.get("edit_nr") else "🧾 Rechnung") if o == "rechnung"
    else f"📋 Übersicht ({len(rows)})")
ANSICHT = ss.get("ansicht") or "rechnung"

if ANSICHT == "uebersicht":
    if not store():
        st.warning(f"GitHub-Ablage ist noch nicht eingerichtet: {github_status()}  \n"
                   "Korrigieren unter *Manage app → ⋮ → Settings → Secrets*, danach die Seite neu laden.")
    else:
        n_final = sum(1 for r in rows if r.get("final") == "ja")
        c1, c2 = st.columns([3, 1])
        c1.markdown(f"#### 📋 Erstellte Rechnungen\n{len(rows)} Rechnung(en) · davon **{n_final} final**")
        if c2.button("🔄", help="Übersicht neu laden", width="stretch"):
            journal(force=True)
            ss.editor_v = ss.get("editor_v", 0) + 1
            st.rerun()
        if jerr:
            st.error(f"Übersicht konnte nicht geladen werden: {jerr}")
        if not rows:
            st.caption("Noch keine Rechnungen erstellt.")
        else:
            v = ss.get("editor_v", 0)
            st.caption("**Final** abhaken bzw. **Löschen** markieren und unten **💾 Änderungen speichern**.")
            SP = [3, 1.5, 0.9, 0.9, 1.7]
            with st.container(key="rtab"):
                h = st.columns(SP, vertical_alignment="center")
                for c, (t, cls) in zip(h, [("Rechnung", ""), ("CHF", "r"), ("Final", "c"), ('<span class="lg">Löschen</span><span class="sm">🗑️</span>', "c"),
                                           ("", "")]):
                    c.markdown(f'<div class="rth {cls}">{t}</div>', unsafe_allow_html=True)
                final_neu, loeschen = {}, []
                for r in reversed(rows):
                    nr, ist_fin = r["rechnungs_nr"], r.get("final") == "ja"
                    with st.container(key=f"rz_{nr}_{v}"):
                        c = st.columns(SP, vertical_alignment="center")
                        status = (f'<span class="badge fin">final</span>' if ist_fin
                                  else '<span class="badge draft">Entwurf</span>')
                        c[0].markdown(f'<div class="rcell"><b>{nr}</b> {status}<br>'
                                      f'<span class="rsub">{r["rechnungsdatum"]} · {r["positionen"]}</span></div>',
                                      unsafe_allow_html=True)
                        try:
                            tot = chf(float(r["total_chf"]))
                        except ValueError:
                            tot = r["total_chf"]
                        c[1].markdown(f'<div class="rcell rnum">{tot}</div>', unsafe_allow_html=True)
                        # Final und Löschen schliessen sich gegenseitig aus
                        f_akt = ss.get(f"fin_{nr}_{v}", ist_fin)
                        d_akt = ss.get(f"del_{nr}_{v}", False)
                        f_neu = c[2].checkbox("Final", value=ist_fin, key=f"fin_{nr}_{v}", disabled=d_akt,
                                              label_visibility="collapsed")
                        d_neu = c[3].checkbox("Löschen", value=False, key=f"del_{nr}_{v}", disabled=f_akt,
                                              label_visibility="collapsed")
                        if f_neu and d_neu:              # Sicherheitsnetz
                            d_neu = False
                        if ist_fin:
                            c[4].markdown(f'<div class="rsub">final am<br>{r.get("final_am", "")[:10]}</div>',
                                          unsafe_allow_html=True)
                        else:
                            c[4].button("Entwurf bearbeiten", key=f"bearb_{nr}_{v}", type="tertiary",
                                        on_click=cb_entwurf_laden, args=(nr,))
                        if f_neu != ist_fin:
                            final_neu[nr] = f_neu
                        if d_neu:
                            loeschen.append(nr)
            if final_neu or loeschen:
                teile = []
                if final_neu:
                    teile.append(f"{len(final_neu)} Final-Änderung(en)")
                if loeschen:
                    teile.append(f"{len(loeschen)} zum Löschen markiert (Nr. {', '.join(loeschen)})")
                st.caption(" · ".join(teile) + " – noch nicht gespeichert.")
            ok_loeschen = True
            if loeschen:
                ok_loeschen = st.checkbox(f"Ja, Rechnung {', '.join(loeschen)} endgültig löschen (PDF + Excel)")
            if st.button("💾 Änderungen speichern", type="primary", width="stretch",
                         disabled=not (final_neu or loeschen) or not ok_loeschen):
                try:
                    neu_rows = rows
                    if final_neu:
                        neu_rows = store().final_setzen(final_neu, jetzt())
                    if loeschen:
                        neu_rows = store().loeschen(loeschen)
                    journal_setzen(neu_rows)
                    if not ss.get("gruppen") or ss.get("nr") in set(final_neu) | set(loeschen):
                        ss.nr_neu = True
                    st.rerun()
                except Exception as e:
                    st.error(f"Speichern fehlgeschlagen: {e}")

            st.markdown("##### Rechnung ansehen")
            nrs = [r["rechnungs_nr"] for r in reversed(rows)]
            wahl = st.selectbox("Rechnung", nrs, index=0, label_visibility="collapsed",
                                format_func=lambda n: next(f"Nr. {r['rechnungs_nr']} · {r['positionen']} · CHF {r['total_chf']}"
                                                           + (" · ✅ final" if r.get('final') == 'ja' else " · Entwurf")
                                                           for r in rows if r['rechnungs_nr'] == n))
            if wahl:
                r0 = next(r for r in rows if r["rechnungs_nr"] == wahl)
                cache = ss.setdefault("pdf_cache", {})
                key = (wahl, r0.get("erstellt"))
                try:
                    if key not in cache:
                        cache[key] = (store().get(r0["pdf"])[0], store().get(r0["excel"])[0])
                    pdf, xlsx = cache[key]
                    c1, c2 = st.columns(2)
                    if pdf:
                        c1.download_button("⬇️ PDF", pdf, r0["pdf"].split("/")[-1], "application/pdf",
                                           width="stretch", type="primary", key=f"dl_pdf_{wahl}")
                    if xlsx:
                        c2.download_button("⬇️ Excel", xlsx, r0["excel"].split("/")[-1],
                                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                           width="stretch", key=f"dl_x_{wahl}")
                    if r0.get("final") != "ja":
                        st.button(f"✏️ Entwurf {wahl} bearbeiten", key=f"edit_{wahl}", on_click=cb_entwurf_laden,
                                  args=(wahl,), width="stretch")
                    if pdf:
                        import base64
                        b64 = base64.b64encode(pdf).decode()
                        st.markdown(f'<iframe src="data:application/pdf;base64,{b64}" width="100%" height="620" '
                                    f'style="border:1px solid #C9D6E6;border-radius:10px;"></iframe>',
                                    unsafe_allow_html=True)
                        st.caption("Wird die Vorschau auf dem Handy nicht angezeigt: PDF herunterladen.")
                except Exception as e:
                    st.error(f"Rechnung konnte nicht geladen werden: {e}")


if ANSICHT == "rechnung":
    with st.expander("ℹ️ So funktioniert's", expanded=not ss.gruppen):
        st.markdown(
            "1. **Rechnung** – Datum und Ausführungszeitraum prüfen. Die Nummer wird automatisch vorgeschlagen.\n"
            "2. **Position wählen** – z.B. *273.3 · Innentüren* → **Hinzufügen**. "
            "Arbeiten, die nicht in der Liste stehen → *Zusatzleistung*.\n"
            "3. **Leistungen erfassen** – erledigte Mängelpunkte abhaken (Haken weg = Zeile weg). "
            "Pro Zeile **Menge, Einheit und Preis** eintragen. *Unterpunkte* erscheinen als einzelne «–»-Zeilen.\n"
            "4. ✍️ **Schneller:** ChatGPT fragen, welche Arbeiten dazugehören (z.B. *«Welche Arbeitsschritte gehören zu: "
            "Dachrinne undicht beim Zusammenschluss? Als kurze Liste.»*), Antwort kopieren und bei "
            "*Leistungen als Text einfügen* hineinkopieren → jeder Punkt wird eine Zeile.\n"
            "5. **Erstellen** – Total prüfen → **Rechnung erstellen** → PDF/Excel herunterladen oder per Mail senden.\n"
            "6. **Übersicht** – in der *📋 Übersicht* Rechnungen ansehen, als **final** abhaken oder Entwürfe löschen. "
"Offene Entwürfe mit **✏️ Bearbeiten** wieder öffnen und ändern. "
            "Jeder Entwurf behält seine Nummer; neue Rechnungen erhalten automatisch die kleinste freie Nummer. "
            "Wird ein Entwurf gelöscht, wird seine Nummer wieder frei.\n\n"
            "💡 Material getrennt von der Arbeit erfassen (z.B. *11 Stk. à 525.–* und *Montage pauschal*) – das will die Bank sehen.")


    entwuerfe = [r for r in rows if r.get("final") != "ja"]
    if m := ss.pop("lade_msg", None):
        getattr(st, m[0])(m[1])
    if ss.get("edit_nr"):
        with st.container(border=True):
            st.markdown(f"**✏️ Du bearbeitest den Entwurf Nr. {ss.edit_nr}.** "
                        "«Rechnung erstellen» überschreibt ihn mit den neuen Angaben.")
            st.button("✖️ Bearbeitung beenden – neue Rechnung", on_click=cb_bearbeitung_abbrechen, width="stretch")
    elif entwuerfe:
        with st.expander(f"✏️ Entwurf weiterbearbeiten ({len(entwuerfe)} offen)"):
            wahl_e = st.selectbox("Entwurf", [r["rechnungs_nr"] for r in reversed(entwuerfe)], index=None,
                                  placeholder="Entwurf wählen …", key="entwurf_wahl",
                                  format_func=lambda n: next(f"Nr. {r['rechnungs_nr']} · {r['positionen']} · "
                                                             f"CHF {r['total_chf']} · {r['erstellt']}"
                                                             for r in entwuerfe if r["rechnungs_nr"] == n))
            st.button("Entwurf laden", on_click=lambda: cb_entwurf_laden(ss.entwurf_wahl) if ss.get("entwurf_wahl") else None,
                      type="primary", width="stretch", disabled=not wahl_e)

    # ------------------------------------------------------------------ 1. Rechnung
    with st.container(border=True):
        step(1, "Rechnung")
        if "nr" not in ss:
            ss.nr = naechste_nr() if store() else str(START)
        if jerr and store():
            st.warning(f"Rechnungsnummer konnte nicht aus GitHub gelesen werden ({jerr}).")
        c1, c2 = st.columns(2)
        c1.text_input("Rechnungs-Nr.", key="nr", disabled=bool(ss.get("edit_nr")),
                      help="Beim Bearbeiten eines Entwurfs bleibt die Nummer fix." if ss.get("edit_nr") else None)
        for _k in ("datum", "von", "bis"):
            ss.setdefault(_k, date.today())
        datum = c2.date_input("Rechnungsdatum", key="datum", format="DD.MM.YYYY")
        c1, c2 = st.columns(2)
        von = c1.date_input("Ausgeführt von", key="von", format="DD.MM.YYYY")
        bis = c2.date_input("bis", key="bis", format="DD.MM.YYYY")
        vorhanden = next((r for r in rows if r.get("rechnungs_nr") == ss.nr.strip()), None)
        NR_FINAL = bool(vorhanden and vorhanden.get("final") == "ja")
        if NR_FINAL:
            st.error(f"Rechnung {ss.nr} ist bereits **final**. Nächste freie Nummer: **{naechste_nr()}**.")
        elif vorhanden and ss.get("edit_nr") != ss.nr.strip():
            st.error(f"Nr. {ss.nr} ist schon als **Entwurf** vergeben ({vorhanden['positionen']}, {vorhanden['erstellt']}). "
                     f"Nächste freie Nummer: **{naechste_nr()}** – oder den Entwurf oben über «✏️ Entwurf "
                     f"weiterbearbeiten» laden.")
        NR_BELEGT = bool(vorhanden and not NR_FINAL and ss.get("edit_nr") != ss.nr.strip())

    # ------------------------------------------------------------------ 2. Position
    with st.container(border=True):
        step(2, "Position wählen")
        optionen = [f"{p['pos']} · {p['titel']}" for p in DATA["positionen"]] + [ZUSATZ]
        wahl = st.selectbox("Position aus der Mängelliste", optionen, index=None,
                            placeholder="Position wählen …", key="pos_wahl")
        if wahl == ZUSATZ:
            st.text_input("Bezeichnung der Zusatzleistung", key="zusatz_titel",
                          placeholder="z.B. Zusätzliche Steckdose Garage")
            st.button("Zusatzleistung hinzufügen", on_click=cb_zusatz_hinzufuegen,
                      type="primary", width="stretch")
        elif wahl:
            p = POS[wahl.split(" · ")[0]]
            st.caption(f"{len(p['maengel'])} Mängelpunkte · Schätzung Architekt CHF {chf(p['schaetzung'])}")
            st.button("Hinzufügen", on_click=cb_position_hinzufuegen, type="primary", width="stretch")
        if msg := ss.pop("hinweis", None):
            st.info(msg)

    # ------------------------------------------------------------------ 3. Leistungen
    if ss.gruppen:
        step(3, "Leistungen erfassen")
    gobjs = gruppen_obj()
    for gi, (g, gobj) in enumerate(zip(list(ss.gruppen), gobjs), 1):
        gid = g["id"]
        kopf = f"{gi} · {g['bkp'] + ' · ' if g['bkp'] else '➕ '}{gobj.titel}   —   CHF {chf(gobj.total)}"
        with st.expander(kopf, expanded=True):
            st.text_input("Titel auf der Rechnung", key=f"{gid}_titel")

            if g["bkp"]:
                st.markdown("**Erledigte Punkte aus der Mängelliste**")
                for m in POS[g["bkp"]]["maengel"]:
                    k = f"{gid}_m_{m['nr']}"
                    ss.setdefault(k, m["nr"] in g["auto"])
                    st.checkbox(f"{m['nr']} – {m['text']}", key=k, on_change=cb_mangel, args=(gid, m["nr"], m["text"]))
                bezug_opt = [""] + [m["nr"] for m in POS[g["bkp"]]["maengel"]]

            if not g["zeilen"]:
                st.caption("Noch keine Zeilen – Mängelpunkt abhaken oder unten eine Zeile hinzufügen.")
            for zi, zid in enumerate(g["zeilen"], 1):
                with st.container(border=True):
                    z = zeile_aus_state(zid)
                    chips = f'<span class="chip">{gi}.{zi}</span>' + (f'<span class="chip">↳ {z.bezug}</span>' if z.bezug else "")
                    st.markdown(chips, unsafe_allow_html=True)
                    st.text_input("Leistung", key=f"{zid}_text")
                    a, b, c = st.columns([1, 1, 1.2])
                    a.number_input("Menge", key=f"{zid}_menge", min_value=0.0, step=1.0, format="%.2f")
                    _e = ss.get(f"{zid}_einheit")
                    b.selectbox("Einheit", EINHEITEN if _e in EINHEITEN or not _e else EINHEITEN + [_e],
                                key=f"{zid}_einheit")
                    c.number_input("Preis CHF", key=f"{zid}_ansatz", min_value=0.0, step=10.0, format="%.2f")
                    if g["bkp"]:
                        st.selectbox("Bezug Mängelliste", bezug_opt, key=f"{zid}_bezug",
                                     format_func=lambda x: x or "– kein direkter Bezug –")
                    st.text_area("Unterpunkte (eine pro Zeile, optional)", key=f"{zid}_unter", height=80,
                                 placeholder="Montage Zargen\nEinhängen Türblätter\nEntsorgung Verpackung")
                    cl, cr = st.columns([2, 1])
                    cl.markdown(f"Betrag: **CHF {chf(z.betrag)}**")
                    cr.button("🗑️ Zeile", key=f"del_{zid}", on_click=cb_zeile_loeschen, args=(gid, zid),
                              width="stretch")

            with st.expander("✍️ Leistungen als Text einfügen (z.B. aus ChatGPT)", expanded=not g["zeilen"]):
                st.caption("Liste der Arbeiten hier einfügen – jeder Punkt wird eine eigene Zeile, eingerückte Punkte "
                           "werden Unterpunkte. Überschriften und Einleitungssätze werden ignoriert. "
                           "Menge und Preis danach bei jeder Zeile eintragen.")
                st.text_area("Leistungen", key=f"{gid}_schnell", height=190, label_visibility="collapsed",
                             placeholder="- Gerüst aufstellen und sichern\n- Dachrinne reinigen\n"
                                         "- Alte Verbindung lösen\n    - Dichtmaterial entfernen\n"
                                         "- Rinnenverbinder montieren und verlöten\n- Dichtigkeitsprüfung mit Wasser")
                st.checkbox("Menge und Preis aus dem Text erkennen (z.B. «11 Stk à 525»)", key=f"{gid}_mitpreis")
                st.button("Zeilen erstellen", key=f"schnell_btn_{gid}", on_click=cb_schnell, args=(gid,),
                          type="primary", width="stretch")
                h = ss.get("hinweis_schnell")
                if h and h[0] == gid:
                    st.success(f"{h[1]} Zeile(n) erstellt – bitte oben kurz prüfen.")
                    ss.pop("hinweis_schnell", None)

            st.markdown("**Zeile hinzufügen**")
            cols = st.columns(3)
            for i, name in enumerate(VORLAGEN):
                cols[i % 3].button(f"+ {name}", key=f"add_{name}_{gid}", on_click=cb_zeile_neu, args=(gid, name),
                                   width="stretch")
            cols[2].button("🗑️ Position", key=f"delg_{gid}", on_click=cb_gruppe_loeschen, args=(gid,),
                           width="stretch")

    # ------------------------------------------------------------------ 4. Erstellen
    if ss.gruppen:
        firma = dict(S.get("firma", {}))
        empf_default = dict(S.get("empfaenger", {}))
        with st.container(border=True):
            step(4, "Prüfen & erstellen")
            with st.expander("Empfänger"):
                for _k in ("name", "strasse", "plz_ort"):
                    ss.setdefault(f"empf_{_k}", empf_default.get(_k, ""))
                empf = {"name": st.text_input("Name", key="empf_name"),
                        "strasse": st.text_input("Strasse", key="empf_strasse"),
                        "plz_ort": st.text_input("PLZ / Ort", key="empf_plz_ort")}
            r = Rechnung(nr=ss.nr.strip(), datum=datum, zahlbar_bis=datum + timedelta(days=30),
                         ausfuehrung=f"{dstr(von)} – {dstr(bis)}" if von != bis else dstr(von),
                         firma=firma, empfaenger=empf,
                         objekt=S.get("objekt", {}).get("text", ""),
                         grundlage=S.get("objekt", {}).get("grundlage", DATA["quelle"]),
                         gruppen=gruppen_obj())
            sub = (f"exkl. MWST {chf(r.total_exkl)} · MWST 8.1 % {chf(r.mwst)}" if r.mwst_pflichtig
                   else "ohne MWST (keine MWST-Nr. hinterlegt)")
            n_zeilen = sum(len(g.zeilen) for g in r.gruppen)
            st.markdown(f'<div class="total"><div class="lbl">Rechnung {r.nr} · {len(r.gruppen)} Position(en) · '
                        f'{n_zeilen} Zeile(n)</div><div class="val">CHF {chf(r.total)}</div>'
                        f'<div class="sub">{sub} · zahlbar bis {dstr(r.zahlbar_bis)}</div></div>',
                        unsafe_allow_html=True)

            fehler = []
            if not firma.get("name"):
                fehler.append("Firmendaten fehlen in den Secrets.")
            for g in r.gruppen:
                if not g.zeilen:
                    fehler.append(f"Position «{g.titel}» hat noch keine Zeilen.")
                for z in g.zeilen:
                    if not z.text:
                        fehler.append(f"Eine Zeile in «{g.titel}» hat keinen Leistungstext.")
                    if z.ansatz is None:
                        fehler.append(f"«{z.text or 'Zeile'}» hat noch keinen Preis.")
            if NR_FINAL:
                fehler.append(f"Rechnungs-Nr. {r.nr} ist final – bitte die nächste freie Nummer verwenden.")
            if NR_BELEGT:
                fehler.append(f"Rechnungs-Nr. {r.nr} ist schon ein anderer Entwurf – bitte die nächste freie Nummer verwenden.")
            for f in dict.fromkeys(fehler):
                st.warning(f)

            mail_cfg = S.get("mail")
            per_mail, mail_an = False, ""
            if mail_cfg:
                per_mail = st.toggle("Zusätzlich per E-Mail senden")
                if per_mail:
                    mail_an = st.text_input("E-Mail an (mehrere mit Komma)", mail_cfg.get("standard_empfaenger", ""))

            if st.button("✅ Rechnung erstellen", type="primary", width="stretch", disabled=bool(fehler)):
                with st.spinner("Rechnung wird erstellt …"):
                    xlsx, pdf, name = als_excel(r), als_pdf(r), dateiname(r)
                    res = {"name": name, "xlsx": xlsx, "pdf": pdf, "msgs": []}
                    if store():
                        try:
                            info = store().speichern(r.nr, name, xlsx, pdf, {
                                "rechnungs_nr": r.nr, "rechnungsdatum": dstr(r.datum),
                                "positionen": "; ".join(g.bkp or "Zusatz" for g in r.gruppen),
                                "total_chf": f"{r.total:.2f}", "erstellt": jetzt()},
                                daten=__import__("entwurf").zu_json(formular_daten()))
                            ss.edit_nr = r.nr
                            journal_setzen(info.get("rows"))
                            res["msgs"].append(("success", f"Als Entwurf in GitHub gespeichert: {info['pdf']}. "
                                                           "In der «📋 Übersicht» als final abhaken, wenn sie definitiv ist."))
                        except Exception as e:
                            txt = str(e)
                            if "401" in txt:
                                txt = "Token ungültig oder abgelaufen (401). Token in den Secrets prüfen."
                            elif "404" in txt:
                                txt = "Repo nicht gefunden (404). Repo-Name in den Secrets und Token-Berechtigung prüfen."
                            res["msgs"].append(("error", f"GitHub-Speicherung fehlgeschlagen: {txt}"))
                    else:
                        res["msgs"].append(("info", f"Nur zum Download – GitHub nicht eingerichtet: {github_status()}"))
                    if per_mail:
                        try:
                            from mailer import senden
                            an = [a.strip() for a in mail_an.split(",") if a.strip()]
                            senden(mail_cfg, an, f"Rechnung {r.nr} – {r.firma['name']}",
                                   f"Guten Tag\n\nIm Anhang die Rechnung {r.nr} über CHF {chf(r.total)}.\n\n"
                                   f"Freundliche Grüsse\n{r.firma['name']}",
                                   [(f"{name}.pdf", pdf, "application/pdf"),
                                    (f"{name}.xlsx", xlsx,
                                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
                            res["msgs"].append(("success", f"E-Mail gesendet an {', '.join(an)}"))
                        except Exception as e:
                            res["msgs"].append(("error", f"E-Mail fehlgeschlagen: {e}"))
                    ss.result = res
                    st.rerun()

        if res := ss.get("result"):
            with st.container(border=True):
                st.markdown("#### 🎉 Rechnung erstellt")
                for kind, m in res["msgs"]:
                    getattr(st, kind)(m)
                c1, c2 = st.columns(2)
                c1.download_button("⬇️ PDF", res["pdf"], f"{res['name']}.pdf", "application/pdf",
                                   width="stretch", type="primary")
                c2.download_button("⬇️ Excel", res["xlsx"], f"{res['name']}.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   width="stretch")
                st.button("Neue Rechnung beginnen", width="stretch", on_click=cb_bearbeitung_abbrechen)
