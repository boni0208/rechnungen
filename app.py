"""Rechnungs-App Hausbau Gerlafingen – Eingabe per Handy, Ausgabe als Excel + PDF."""
from __future__ import annotations

import json
import uuid
from datetime import date, timedelta
from pathlib import Path

import streamlit as st

from invoice import Gruppe, Rechnung, Zeile, als_excel, als_pdf, chf, dateiname, dstr

st.set_page_config(page_title="Rechnung erfassen", page_icon="🧾", layout="centered")

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
    st.error("Datei «maengelliste.json» fehlt im Repo. Bitte in den Ordner «data» oder direkt neben app.py hochladen.")
    st.stop()
DATA = json.loads(_JSON.read_text(encoding="utf-8"))
POS = {p["pos"]: p for p in DATA["positionen"]}
S = st.secrets
ss = st.session_state


# ------------------------------------------------------------------ Zugang
def login() -> bool:
    pw = S.get("APP_PASSWORT")
    if not pw or ss.get("ok"):
        return True
    st.title("🧾 Rechnungen Hausbau")
    eingabe = st.text_input("Passwort", type="password")
    if st.button("Anmelden", type="primary", use_container_width=True):
        if eingabe == pw:
            ss.ok = True
            st.rerun()
        st.error("Falsches Passwort.")
    return False


if not login():
    st.stop()


@st.cache_resource
def store():
    g = S.get("github")
    if not g or not g.get("token"):
        return None
    from storage import GitHubStore
    return GitHubStore(g["token"], g["repo"], g.get("branch", "rechnungen"))


# ------------------------------------------------------------------ Zustand
ss.setdefault("gruppen", [])          # [{id, bkp, titel, zeilen:[id,...]}]


def neue_zeile(text="", menge=1.0, einheit="pauschal", ansatz=None, bezug="", unter="") -> str:
    zid = uuid.uuid4().hex[:8]
    ss[f"{zid}_text"], ss[f"{zid}_menge"], ss[f"{zid}_einheit"] = text, menge, einheit
    ss[f"{zid}_ansatz"], ss[f"{zid}_bezug"], ss[f"{zid}_unter"] = ansatz, bezug, unter
    return zid


def zeile_aus_state(zid: str) -> Zeile:
    unter = [u.strip(" -–•") for u in (ss.get(f"{zid}_unter") or "").splitlines() if u.strip(" -–•")]
    return Zeile(text=(ss.get(f"{zid}_text") or "").strip(), menge=ss.get(f"{zid}_menge"),
                 einheit=ss.get(f"{zid}_einheit") or "", ansatz=ss.get(f"{zid}_ansatz"),
                 bezug=ss.get(f"{zid}_bezug") or "", unterpunkte=unter)


def gruppen_obj() -> list[Gruppe]:
    out = []
    for g in ss.gruppen:
        titel = ss.get(f"{g['id']}_titel", g["titel"])
        out.append(Gruppe(titel=titel, bkp=g["bkp"], zeilen=[zeile_aus_state(z) for z in g["zeilen"]]))
    return out


# ------------------------------------------------------------------ 1. Kopf
st.title("🧾 Rechnung erfassen")
st.caption(f"Grundlage: {DATA['quelle']}")

with st.container(border=True):
    st.subheader("1 · Rechnung")
    if "nr" not in ss:
        start = int(S.get("START_NR", 1024))
        try:
            ss.nr = str(store().naechste_nr(start)) if store() else str(start)
        except Exception as e:                       # GitHub nicht erreichbar → Startnummer
            st.warning(f"Rechnungsnummer konnte nicht aus GitHub gelesen werden ({e}).")
            ss.nr = str(start)
    st.text_input("Rechnungs-Nr.", key="nr")
    datum = st.date_input("Rechnungsdatum", value=date.today(), format="DD.MM.YYYY")
    c1, c2 = st.columns(2)
    von = c1.date_input("Ausgeführt von", value=date.today(), format="DD.MM.YYYY")
    bis = c2.date_input("bis", value=date.today(), format="DD.MM.YYYY")

# ------------------------------------------------------------------ 2. Position hinzufügen
with st.container(border=True):
    st.subheader("2 · Position aus Mängelliste wählen")
    optionen = [f"{p['pos']} · {p['titel']}" for p in DATA["positionen"]] + [ZUSATZ]
    wahl = st.selectbox("Position", optionen, index=None, placeholder="Position wählen …")
    if wahl == ZUSATZ:
        titel = st.text_input("Bezeichnung der Zusatzleistung", placeholder="z.B. Zusätzliche Steckdose Garage")
        if st.button("Zusatzleistung hinzufügen", use_container_width=True, disabled=not titel):
            ss.gruppen.append({"id": uuid.uuid4().hex[:8], "bkp": "", "titel": titel, "zeilen": [neue_zeile()]})
            st.rerun()
    elif wahl:
        p = POS[wahl.split(" · ")[0]]
        st.markdown("**Vorschlag – erledigte Punkte abhaken:**")
        gewaehlt = [m for m in p["maengel"] if st.checkbox(f"{m['nr']} – {m['text']}", key=f"chk_{m['nr']}")]
        if st.button("Position übernehmen", type="primary", use_container_width=True):
            zeilen = [neue_zeile(text=m["text"], bezug=m["nr"]) for m in gewaehlt] or [neue_zeile()]
            ss.gruppen.append({"id": uuid.uuid4().hex[:8], "bkp": p["pos"], "titel": p["titel"], "zeilen": zeilen})
            for m in p["maengel"]:
                ss.pop(f"chk_{m['nr']}", None)
            st.rerun()

# ------------------------------------------------------------------ 3. Zeilen bearbeiten
if ss.gruppen:
    st.subheader("3 · Leistungen erfassen")
    st.caption("Pro Zeile: was wurde gemacht, Menge, Einheit, Preis. Unterpunkte erscheinen als einzelne Zeilen mit «–».")
for gi, g in enumerate(list(ss.gruppen), 1):
    gobj = gruppen_obj()[gi - 1]
    label = f"{gi} · {g['bkp'] + ' ' if g['bkp'] else ''}{gobj.titel}  —  CHF {chf(gobj.total)}"
    with st.expander(label, expanded=True):
        ss.setdefault(f"{g['id']}_titel", g["titel"])
        st.text_input("Titel auf der Rechnung", key=f"{g['id']}_titel")
        bezug_opt = [""] + ([m["nr"] for m in POS[g["bkp"]]["maengel"]] if g["bkp"] else [])
        for zi, zid in enumerate(list(g["zeilen"]), 1):
            with st.container(border=True):
                st.markdown(f"**{gi}.{zi}**")
                st.text_input("Leistung", key=f"{zid}_text")
                a, b, c = st.columns([1, 1, 1.2])
                a.number_input("Menge", key=f"{zid}_menge", min_value=0.0, step=1.0, format="%.2f")
                b.selectbox("Einheit", EINHEITEN, key=f"{zid}_einheit")
                c.number_input("Ansatz CHF", key=f"{zid}_ansatz", min_value=0.0, step=10.0, format="%.2f")
                if g["bkp"]:
                    st.selectbox("Bezug Mängelliste", bezug_opt, key=f"{zid}_bezug",
                                 format_func=lambda x: x or "– kein direkter Bezug –")
                st.text_area("Unterpunkte (eine pro Zeile, optional)", key=f"{zid}_unter", height=80,
                             placeholder="Montage Zargen\nEinhängen Türblätter\nEntsorgung Verpackung")
                z = zeile_aus_state(zid)
                cl, cr = st.columns([2, 1])
                cl.markdown(f"Betrag: **CHF {chf(z.betrag)}**")
                if cr.button("🗑️ Zeile", key=f"del_{zid}", use_container_width=True):
                    g["zeilen"].remove(zid)
                    st.rerun()
        st.markdown("Zeile hinzufügen:")
        cols = st.columns(3)
        for i, (name, (txt, m, e)) in enumerate(VORLAGEN.items()):
            if cols[i % 3].button(f"+ {name}", key=f"add_{name}_{g['id']}", use_container_width=True):
                g["zeilen"].append(neue_zeile(txt, m, e))
                st.rerun()
        if cols[2].button("🗑️ Position", key=f"delg_{g['id']}", use_container_width=True):
            ss.gruppen.remove(g)
            st.rerun()

# ------------------------------------------------------------------ 4. Erstellen
if ss.gruppen:
    firma = dict(S.get("firma", {}))
    empf_default = dict(S.get("empfaenger", {}))
    with st.container(border=True):
        st.subheader("4 · Rechnung erstellen")
        with st.expander("Empfänger prüfen"):
            empf = {"name": st.text_input("Name", empf_default.get("name", "")),
                    "strasse": st.text_input("Strasse", empf_default.get("strasse", "")),
                    "plz_ort": st.text_input("PLZ / Ort", empf_default.get("plz_ort", ""))}
        r = Rechnung(nr=ss.nr.strip(), datum=datum, zahlbar_bis=datum + timedelta(days=30),
                     ausfuehrung=f"{dstr(von)} – {dstr(bis)}" if von != bis else dstr(von),
                     firma=firma, empfaenger=empf,
                     objekt=S.get("objekt", {}).get("text", ""),
                     grundlage=S.get("objekt", {}).get("grundlage", DATA["quelle"]),
                     gruppen=gruppen_obj())
        if r.mwst_pflichtig:
            st.write(f"Total exkl. MWST: CHF {chf(r.total_exkl)} · MWST 8.1 %: CHF {chf(r.mwst)}")
        st.markdown(f"### Total CHF {chf(r.total)}")
        if not r.mwst_pflichtig:
            st.caption("Ohne MWST (keine MWST-Nr. hinterlegt).")

        fehler = []
        if not firma.get("name"):
            fehler.append("Firmendaten fehlen in den Secrets.")
        for g in r.gruppen:
            for z in g.zeilen:
                if not z.text:
                    fehler.append(f"Eine Zeile in «{g.titel}» hat keinen Leistungstext.")
                if z.ansatz is None:
                    fehler.append(f"«{z.text or 'Zeile'}» hat keinen Ansatz.")
        for f in dict.fromkeys(fehler):
            st.warning(f)

        mail_cfg = S.get("mail")
        per_mail = False
        if mail_cfg:
            per_mail = st.checkbox("Zusätzlich per E-Mail senden")
            mail_an = st.text_input("E-Mail an (mehrere mit Komma)", mail_cfg.get("standard_empfaenger", ""),
                                    disabled=not per_mail)

        if st.button("✅ Rechnung erstellen", type="primary", use_container_width=True, disabled=bool(fehler)):
            with st.spinner("Rechnung wird erstellt …"):
                xlsx, pdf, name = als_excel(r), als_pdf(r), dateiname(r)
                res = {"name": name, "xlsx": xlsx, "pdf": pdf, "msgs": []}
                if store():
                    try:
                        info = store().speichern(r.nr, name, xlsx, pdf, {
                            "rechnungs_nr": r.nr, "rechnungsdatum": dstr(r.datum),
                            "positionen": "; ".join(g.bkp or "Zusatz" for g in r.gruppen),
                            "total_chf": f"{r.total:.2f}", "erstellt": date.today().isoformat()})
                        res["msgs"].append(("success", f"In GitHub gespeichert: {info['pdf']}"))
                    except Exception as e:
                        res["msgs"].append(("error", f"GitHub-Speicherung fehlgeschlagen: {e}"))
                else:
                    res["msgs"].append(("info", "GitHub ist nicht eingerichtet – nur Download."))
                if per_mail:
                    try:
                        from mailer import senden
                        an = [a.strip() for a in mail_an.split(",") if a.strip()]
                        senden(mail_cfg, an, f"Rechnung {r.nr} – {r.firma['name']}",
                               f"Guten Tag\n\nIm Anhang die Rechnung {r.nr} über CHF {chf(r.total)}.\n\nFreundliche Grüsse\n{r.firma['name']}",
                               [(f"{name}.pdf", pdf, "application/pdf"),
                                (f"{name}.xlsx", xlsx, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
                        res["msgs"].append(("success", f"E-Mail gesendet an {', '.join(an)}"))
                    except Exception as e:
                        res["msgs"].append(("error", f"E-Mail fehlgeschlagen: {e}"))
                ss.result = res

    if res := ss.get("result"):
        for kind, m in res["msgs"]:
            getattr(st, kind)(m)
        c1, c2 = st.columns(2)
        c1.download_button("⬇️ PDF", res["pdf"], f"{res['name']}.pdf", "application/pdf", use_container_width=True)
        c2.download_button("⬇️ Excel", res["xlsx"], f"{res['name']}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        if st.button("Neue Rechnung beginnen", use_container_width=True):
            for k in list(ss.keys()):
                if k != "ok":
                    del ss[k]
            st.rerun()
