"""Rechnungsmodell und Ausgabe als Excel (A4) und PDF – Layout wie Rechnung 1023."""
from __future__ import annotations

import io
from dataclasses import dataclass, field
from datetime import date

MWST_SATZ = 0.081


# ----------------------------------------------------------------- Modell
@dataclass
class Zeile:
    text: str
    menge: float | None = None
    einheit: str = ""
    ansatz: float | None = None
    bezug: str = ""                      # Mängel-Nr., z.B. "273.3-01"
    unterpunkte: list[str] = field(default_factory=list)

    @property
    def betrag(self) -> float:
        if self.menge is None or self.ansatz is None:
            return 0.0
        return round(self.menge * self.ansatz, 2)


@dataclass
class Gruppe:
    titel: str
    bkp: str = ""                         # "" = Zusatzleistung ohne Bezug zur Mängelliste
    zeilen: list[Zeile] = field(default_factory=list)

    @property
    def total(self) -> float:
        return round(sum(z.betrag for z in self.zeilen), 2)


@dataclass
class Rechnung:
    nr: str
    datum: date
    zahlbar_bis: date
    ausfuehrung: str
    firma: dict
    empfaenger: dict
    objekt: str
    grundlage: str
    gruppen: list[Gruppe]

    @property
    def mwst_pflichtig(self) -> bool:
        return bool(self.firma.get("mwst_nr", "").strip())

    @property
    def total_exkl(self) -> float:
        return round(sum(g.total for g in self.gruppen), 2)

    @property
    def mwst(self) -> float:
        if not self.mwst_pflichtig:
            return 0.0
        return round(self.total_exkl * MWST_SATZ * 20) / 20     # auf 5 Rappen

    @property
    def total(self) -> float:
        return round(self.total_exkl + self.mwst, 2)

    @property
    def bezug_text(self) -> str:
        teile = [f"Pos. {g.bkp} {g.titel}" for g in self.gruppen if g.bkp]
        return "; ".join(teile) if teile else "–"

    def tabellenzeilen(self):
        """Flache Liste für die Ausgabe: (art, posnr, text, menge, einheit, ansatz, betrag)."""
        rows = []
        mehrere = len(self.gruppen) > 1
        for gi, g in enumerate(self.gruppen, 1):
            titel = g.titel + (f" (Mängelliste Pos. {g.bkp})" if g.bkp else " (Zusatzleistung)")
            rows.append(("titel", str(gi), titel, None, "", None, None))
            for zi, z in enumerate(g.zeilen, 1):
                rows.append(("zeile", f"{gi}.{zi}", z.text, z.menge, z.einheit, z.ansatz, z.betrag))
                for u in z.unterpunkte:
                    rows.append(("unter", "", f"– {u}", None, "", None, None))
                if z.bezug:
                    rows.append(("bezug", "", f"Bezug Mängelliste: {z.bezug}", None, "", None, None))
            if mehrere:
                rows.append(("subtotal", "", f"Total Pos. {gi}", None, "", None, g.total))
        return rows


def chf(x: float | None) -> str:
    if x is None:
        return ""
    s = f"{x:,.2f}".replace(",", "'")
    return s


def dstr(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def dateiname(r: Rechnung) -> str:
    bkps = "_".join(g.bkp for g in r.gruppen if g.bkp) or "zusatz"
    return f"Rechnung_{r.nr}_{bkps}"


# ----------------------------------------------------------------- Excel
def als_excel(r: Rechnung) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    F, NAVY = "Arial", "1F3A5F"
    CHF = '#,##0.00;-#,##0.00;"-"'
    N = lambda size=10, bold=False, color="000000", italic=False: Font(name=F, size=size, bold=bold, color=color, italic=italic)
    TOP = Alignment(vertical="top")
    LEFT = Alignment(horizontal="left", vertical="top")
    RIGHT = Alignment(horizontal="right", vertical="top")

    wb = Workbook()
    ws = wb.active
    ws.title = str(r.nr)
    for col, w in zip("ABCDEF", [19, 44, 8, 10, 12, 14]):
        ws.column_dimensions[col].width = w

    def put(ref, val, font=None, fmt=None, align=None, merge=None):
        ws[ref] = val
        ws[ref].font = font or N()
        ws[ref].alignment = align or TOP
        if fmt:
            ws[ref].number_format = fmt
        if merge:
            ws.merge_cells(merge)

    f, e = r.firma, r.empfaenger
    put("A1", f["name"], N(14, True, NAVY))
    put("A2", f["strasse"]); put("A3", f["plz_ort"])
    put("A4", f"Tel. {f['telefon']}" if f.get("telefon") else "")
    put("A5", f"UID-Nr.: {f['uid']}" + (f" · MWST-Nr.: {f['mwst_nr']}" if r.mwst_pflichtig else ""), N(9))
    put("D8", e["name"], N(10, True), merge="D8:F8")
    put("D9", e["strasse"], merge="D9:F9")
    put("D10", e["plz_ort"], merge="D10:F10")

    put("A13", "Rechnung", N(16, True, NAVY))
    kopf = [("Rechnungs-Nr.:", r.nr), ("Rechnungsdatum:", dstr(r.datum)), ("Objekt:", r.objekt),
            ("Grundlage:", r.grundlage), ("Bezug Mängelliste:", r.bezug_text),
            ("Ausführung:", r.ausfuehrung), ("Zahlbar bis:", dstr(r.zahlbar_bis))]
    for i, (k, v) in enumerate(kopf):
        row = 15 + i
        put(f"A{row}", k, N(10, True))
        put(f"B{row}", v, N(), None, Alignment(horizontal="left", vertical="top", wrap_text=True), merge=f"B{row}:F{row}")

    th = 23
    for i, h in enumerate(["Pos.-Nr.", "Leistung", "Menge", "Einheit", "Ansatz CHF", "Betrag CHF"], 1):
        c = ws.cell(th, i, h)
        c.font = N(10, True, "FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="left" if i <= 2 else ("center" if i == 4 else "right"), vertical="center")
    ws.row_dimensions[th].height = 20

    hair = Border(bottom=Side(style="hair", color="BFBFBF"))
    rows = r.tabellenzeilen()
    row = th + 1
    for idx, (art, pos, text, menge, einheit, ansatz, betrag) in enumerate(rows):
        a = ws.cell(row, 1, pos); a.number_format = "@"; a.alignment = LEFT
        b = ws.cell(row, 2, text); b.alignment = Alignment(vertical="top", wrap_text=True)
        if art == "titel":
            a.font = b.font = N(10, True)
        elif art in ("unter", "bezug"):
            a.font = N()
            b.font = N(9, color="404040", italic=(art == "bezug"))
            b.alignment = Alignment(indent=1, vertical="top", wrap_text=True)
        elif art == "subtotal":
            b.font = N(10, True); b.alignment = RIGHT
            c = ws.cell(row, 6, betrag); c.font = N(10, True); c.number_format = CHF; c.alignment = RIGHT
        else:
            a.font = b.font = N()
            vals = {3: menge, 4: einheit, 5: ansatz, 6: betrag}
            for col, v in vals.items():
                c = ws.cell(row, col, v)
                c.font = N()
                c.alignment = Alignment(horizontal="center" if col == 4 else "right", vertical="top")
                if col == 3:
                    c.number_format = "#,##0.00"
                elif col in (5, 6):
                    c.number_format = CHF
        nxt = rows[idx + 1][0] if idx + 1 < len(rows) else None
        if nxt not in ("unter", "bezug"):
            for col in range(1, 7):
                ws.cell(row, col).border = hair
        row += 1

    t = row + 1
    def totalzeile(rr, label, val, bold=False, double=False):
        put(f"E{rr}", label, N(11 if double else 10, bold), None, RIGHT)
        put(f"F{rr}", val, N(11 if double else 10, bold), CHF, RIGHT)
        if double:
            for col in "EF":
                ws[f"{col}{rr}"].border = Border(top=Side(style="thin"), bottom=Side(style="double"))
    if r.mwst_pflichtig:
        totalzeile(t, "Total exkl. MWST", r.total_exkl, True)
        totalzeile(t + 1, "MWST 8.1 %", r.mwst)
        totalzeile(t + 2, "Total CHF", r.total, True, True)
        p = t + 5
    else:
        totalzeile(t, "Total CHF", r.total, True, True)
        p = t + 3

    put(f"A{p}", "Zahlung:", N(10, True)); put(f"B{p}", "Zahlbar innert 30 Tagen netto", merge=f"B{p}:F{p}")
    put(f"A{p+1}", "IBAN:", N(10, True)); put(f"B{p+1}", f["iban"], merge=f"B{p+1}:F{p+1}")
    put(f"A{p+2}", "Bank:", N(10, True)); put(f"B{p+2}", f["bank"], merge=f"B{p+2}:F{p+2}")
    put(f"A{p+3}", "Zugunsten:", N(10, True)); put(f"B{p+3}", f["name"], merge=f"B{p+3}:F{p+3}")
    put(f"A{p+5}", "Vielen Dank für den Auftrag.", N(10, italic=True))

    ws.print_area = f"A1:F{p+6}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{th}:{th}"
    ws.page_margins.left, ws.page_margins.right = 0.8, 0.6
    ws.page_margins.top = ws.page_margins.bottom = 0.7
    ws.print_options.horizontalCentered = True
    ws.sheet_view.showGridLines = False

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ----------------------------------------------------------------- PDF
def als_pdf(r: Rechnung) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    NAVY = colors.HexColor("#1F3A5F")
    GREY = colors.HexColor("#404040")
    base = ParagraphStyle("b", fontName="Helvetica", fontSize=9.5, leading=12)
    bold = ParagraphStyle("bb", parent=base, fontName="Helvetica-Bold")
    small = ParagraphStyle("s", parent=base, fontSize=8.5, leading=10.5, textColor=GREY, leftIndent=6)
    smalli = ParagraphStyle("si", parent=small, fontName="Helvetica-Oblique")
    right = ParagraphStyle("r", parent=base, alignment=2)
    rightb = ParagraphStyle("rb", parent=bold, alignment=2)
    esc = lambda s: (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    P = lambda s, st=base: Paragraph(esc(s), st)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=20 * mm, rightMargin=16 * mm,
                            topMargin=16 * mm, bottomMargin=16 * mm,
                            title=f"Rechnung {r.nr}", author=r.firma["name"])
    W = A4[0] - 36 * mm
    story = []
    f, e = r.firma, r.empfaenger

    absender = [P(f["name"], ParagraphStyle("h", parent=bold, fontSize=14, leading=17, textColor=NAVY)),
                P(f["strasse"]), P(f["plz_ort"])]
    if f.get("telefon"):
        absender.append(P(f"Tel. {f['telefon']}"))
    absender.append(P(f"UID-Nr.: {f['uid']}" + (f" · MWST-Nr.: {f['mwst_nr']}" if r.mwst_pflichtig else ""),
                      ParagraphStyle("u", parent=base, fontSize=8.5)))
    story += absender
    story.append(Spacer(1, 10 * mm))
    emp = Table([[None, [P(e["name"], bold), P(e["strasse"]), P(e["plz_ort"])]]], colWidths=[W * 0.55, W * 0.45])
    emp.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story += [emp, Spacer(1, 12 * mm)]
    story.append(P("Rechnung", ParagraphStyle("t", parent=bold, fontSize=16, leading=20, textColor=NAVY)))
    story.append(Spacer(1, 4 * mm))
    kopf = [("Rechnungs-Nr.:", r.nr), ("Rechnungsdatum:", dstr(r.datum)), ("Objekt:", r.objekt),
            ("Grundlage:", r.grundlage), ("Bezug Mängelliste:", r.bezug_text),
            ("Ausführung:", r.ausfuehrung), ("Zahlbar bis:", dstr(r.zahlbar_bis))]
    kt = Table([[P(k, bold), P(v)] for k, v in kopf], colWidths=[34 * mm, W - 34 * mm])
    kt.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 1),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 1), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [kt, Spacer(1, 6 * mm)]

    hdr = ParagraphStyle("hd", parent=bold, textColor=colors.white)
    hdrr = ParagraphStyle("hdr", parent=hdr, alignment=2)
    hdrc = ParagraphStyle("hdc", parent=hdr, alignment=1)
    data = [[P("Pos.", hdr), P("Leistung", hdr), P("Menge", hdrr), P("Einheit", hdrc), P("Ansatz CHF", hdrr), P("Betrag CHF", hdrr)]]
    styles = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("VALIGN", (0, 0), (-1, -1), "TOP"),
              ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    rows = r.tabellenzeilen()
    center = ParagraphStyle("c", parent=base, alignment=1)
    for idx, (art, pos, text, menge, einheit, ansatz, betrag) in enumerate(rows):
        i = len(data)
        if art == "titel":
            data.append([P(pos, bold), P(text, bold), "", "", "", ""])
        elif art == "unter":
            data.append(["", P(text, small), "", "", "", ""])
        elif art == "bezug":
            data.append(["", P(text, smalli), "", "", "", ""])
        elif art == "subtotal":
            data.append(["", P(text, rightb), "", "", "", P(chf(betrag), rightb)])
        else:
            data.append([P(pos), P(text), P(chf(menge), right), P(einheit, center), P(chf(ansatz), right), P(chf(betrag), right)])
        if art in ("unter", "bezug"):
            styles.append(("TOPPADDING", (0, i), (-1, i), 0))
        nxt = rows[idx + 1][0] if idx + 1 < len(rows) else None
        if nxt not in ("unter", "bezug"):
            styles.append(("LINEBELOW", (0, i), (-1, i), 0.3, colors.HexColor("#BFBFBF")))
    cw = [16 * mm, W - 16 * mm - 15 * mm - 17 * mm - 24 * mm - 24 * mm, 15 * mm, 17 * mm, 24 * mm, 24 * mm]
    tbl = Table(data, colWidths=cw, repeatRows=1)
    tbl.setStyle(TableStyle(styles))
    story += [tbl, Spacer(1, 5 * mm)]

    tot = []
    if r.mwst_pflichtig:
        tot += [[P("Total exkl. MWST", rightb), P(chf(r.total_exkl), rightb)],
                [P("MWST 8.1 %", right), P(chf(r.mwst), right)]]
    tot.append([P("Total CHF", ParagraphStyle("tt", parent=rightb, fontSize=11)),
                P(chf(r.total), ParagraphStyle("tv", parent=rightb, fontSize=11))])
    tt = Table(tot, colWidths=[36 * mm, 26 * mm], hAlign="RIGHT")
    n = len(tot) - 1
    tt.setStyle(TableStyle([("LINEABOVE", (0, n), (-1, n), 0.6, colors.black),
                            ("LINEBELOW", (0, n), (-1, n), 1.2, colors.black),
                            ("RIGHTPADDING", (0, 0), (-1, -1), 2)]))
    story += [tt, Spacer(1, 10 * mm)]

    zahl = [("Zahlung:", "Zahlbar innert 30 Tagen netto"), ("IBAN:", f["iban"]),
            ("Bank:", f["bank"]), ("Zugunsten:", f["name"])]
    zt = Table([[P(k, bold), P(v)] for k, v in zahl], colWidths=[34 * mm, W - 34 * mm])
    zt.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 1),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    story += [zt, Spacer(1, 6 * mm), P("Vielen Dank für den Auftrag.", ParagraphStyle("d", parent=base, fontName="Helvetica-Oblique"))]

    doc.build(story)
    return buf.getvalue()
