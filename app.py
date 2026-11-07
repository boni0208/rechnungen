"""Rechnungs-App Hausbau Gerlafingen – Eingabe per Handy, Ausgabe als Excel + PDF."""
from __future__ import annotations

import json
import uuid
from datetime import date, timedelta
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
    if not g or not g.get("token") or "DEIN-GITHUB-NAME" in g.get("repo", "") or g["token"].startswith("GITHUB_PAT_HIER"):
        return None
    from storage import GitHubStore
    return GitHubStore(g["token"], g["repo"], g.get("branch", "rechnungen"))


# ------------------------------------------------------------------ Zustand & Callbacks
ss.setdefault("gruppen", [])   # [{id, bkp, titel, zeilen:[zid], auto:{mangel_nr: zid}}]


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


def cb_gruppe_loeschen(gid: str) -> None:
    ss.gruppen = [g for g in ss.gruppen if g["id"] != gid]


def zeile_aus_state(zid: str) -> Zeile:
    unter = [u.strip(" -–•") for u in (ss.get(f"{zid}_unter") or "").splitlines() if u.strip(" -–•")]
    return Zeile(text=(ss.get(f"{zid}_text") or "").strip(), menge=ss.get(f"{zid}_menge"),
                 einheit=ss.get(f"{zid}_einheit") or "", ansatz=ss.get(f"{zid}_ansatz"),
                 bezug=ss.get(f"{zid}_bezug") or "", unterpunkte=unter)


def gruppen_obj() -> list[Gruppe]:
    return [Gruppe(titel=ss.get(f"{g['id']}_titel", g["titel"]), bkp=g["bkp"],
                   zeilen=[zeile_aus_state(z) for z in g["zeilen"]]) for g in ss.gruppen]


# ------------------------------------------------------------------ Kopf + Anleitung
hero("Firma Boniface · Hausbau Gerlafingen · Grundlage: Mängelliste Architekt vom 04.07.2026")

with st.expander("ℹ️ So funktioniert's", expanded=not ss.gruppen):
    st.markdown("""
1. **Rechnung** – Datum und Ausführungszeitraum prüfen. Die Nummer zählt automatisch weiter.
2. **Position wählen** – z.B. *273.3 · Innentüren* → **Hinzufügen**. Arbeiten, die nicht in der Liste stehen → *Zusatzleistung*.
3. **Leistungen erfassen** – erledigte Mängelpunkte abhaken (werden als Zeile übernommen, Haken weg = Zeile weg).
   Pro Zeile **Menge, Einheit und Preis** eintragen. Mit den Knöpfen *+ Material*, *+ Arbeit* … weitere Zeilen anfügen.
   *Unterpunkte* erscheinen auf der Rechnung als einzelne «–»-Zeilen.
4. **Erstellen** – Total prüfen → **Rechnung erstellen** → PDF/Excel herunterladen oder per Mail senden.

💡 Tipp: Material getrennt von der Arbeit erfassen (z.B. *11 Stk. à 525.–* und *Montage pauschal*) – das will die Bank sehen.
""")

# ------------------------------------------------------------------ 1. Rechnung
with st.container(border=True):
    step(1, "Rechnung")
    if "nr" not in ss:
        start = int(S.get("START_NR", 1024))
        try:
            ss.nr = str(store().naechste_nr(start)) if store() else str(start)
        except Exception as e:
            st.warning(f"Rechnungsnummer konnte nicht aus GitHub gelesen werden ({e}).")
            ss.nr = str(start)
    c1, c2 = st.columns(2)
    c1.text_input("Rechnungs-Nr.", key="nr")
    datum = c2.date_input("Rechnungsdatum", value=date.today(), format="DD.MM.YYYY")
    c1, c2 = st.columns(2)
    von = c1.date_input("Ausgeführt von", value=date.today(), format="DD.MM.YYYY")
    bis = c2.date_input("bis", value=date.today(), format="DD.MM.YYYY")

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
                  type="primary", use_container_width=True)
    elif wahl:
        p = POS[wahl.split(" · ")[0]]
        st.caption(f"{len(p['maengel'])} Mängelpunkte · Schätzung Architekt CHF {chf(p['schaetzung'])}")
        st.button("Hinzufügen", on_click=cb_position_hinzufuegen, type="primary", use_container_width=True)
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
                b.selectbox("Einheit", EINHEITEN, key=f"{zid}_einheit")
                c.number_input("Preis CHF", key=f"{zid}_ansatz", min_value=0.0, step=10.0, format="%.2f")
                if g["bkp"]:
                    st.selectbox("Bezug Mängelliste", bezug_opt, key=f"{zid}_bezug",
                                 format_func=lambda x: x or "– kein direkter Bezug –")
                st.text_area("Unterpunkte (eine pro Zeile, optional)", key=f"{zid}_unter", height=80,
                             placeholder="Montage Zargen\nEinhängen Türblätter\nEntsorgung Verpackung")
                cl, cr = st.columns([2, 1])
                cl.markdown(f"Betrag: **CHF {chf(z.betrag)}**")
                cr.button("🗑️ Zeile", key=f"del_{zid}", on_click=cb_zeile_loeschen, args=(gid, zid),
                          use_container_width=True)

        st.markdown("**Zeile hinzufügen**")
        cols = st.columns(3)
        for i, name in enumerate(VORLAGEN):
            cols[i % 3].button(f"+ {name}", key=f"add_{name}_{gid}", on_click=cb_zeile_neu, args=(gid, name),
                               use_container_width=True)
        cols[2].button("🗑️ Position", key=f"delg_{gid}", on_click=cb_gruppe_loeschen, args=(gid,),
                       use_container_width=True)

# ------------------------------------------------------------------ 4. Erstellen
if ss.gruppen:
    firma = dict(S.get("firma", {}))
    empf_default = dict(S.get("empfaenger", {}))
    with st.container(border=True):
        step(4, "Prüfen & erstellen")
        with st.expander("Empfänger"):
            empf = {"name": st.text_input("Name", empf_default.get("name", "")),
                    "strasse": st.text_input("Strasse", empf_default.get("strasse", "")),
                    "plz_ort": st.text_input("PLZ / Ort", empf_default.get("plz_ort", ""))}
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
        for f in dict.fromkeys(fehler):
            st.warning(f)

        mail_cfg = S.get("mail")
        per_mail, mail_an = False, ""
        if mail_cfg:
            per_mail = st.toggle("Zusätzlich per E-Mail senden")
            if per_mail:
                mail_an = st.text_input("E-Mail an (mehrere mit Komma)", mail_cfg.get("standard_empfaenger", ""))

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
                        txt = str(e)
                        if "401" in txt:
                            txt = "Token ungültig oder abgelaufen (401). Token in den Secrets prüfen."
                        elif "404" in txt:
                            txt = "Repo nicht gefunden (404). Repo-Name in den Secrets und Token-Berechtigung prüfen."
                        res["msgs"].append(("error", f"GitHub-Speicherung fehlgeschlagen: {txt}"))
                else:
                    res["msgs"].append(("info", "GitHub ist nicht eingerichtet – Rechnung nur zum Download."))
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

    if res := ss.get("result"):
        with st.container(border=True):
            st.markdown("#### 🎉 Rechnung erstellt")
            for kind, m in res["msgs"]:
                getattr(st, kind)(m)
            c1, c2 = st.columns(2)
            c1.download_button("⬇️ PDF", res["pdf"], f"{res['name']}.pdf", "application/pdf",
                               use_container_width=True, type="primary")
            c2.download_button("⬇️ Excel", res["xlsx"], f"{res['name']}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
            if st.button("Neue Rechnung beginnen", use_container_width=True):
                for k in list(ss.keys()):
                    if k != "ok":
                        del ss[k]
                st.rerun()
