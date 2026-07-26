#!/usr/bin/env python3
"""
Generator of the sample PDF invoices (data 100% fictional) — for the demo.
Writes 3 invoices with the same layout and different data into input/.

The invoice text itself is Polish on purpose: the point of the demo is parsing
a real-world Polish VAT invoice layout.
"""
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle

BASE = Path(__file__).resolve().parent
OUT = BASE / "input"

# DejaVu has the Polish glyphs (already on the system, comes with matplotlib)
FONT = "DejaVu"
pdfmetrics.registerFont(TTFont(FONT, "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont(FONT + "-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

INVOICES = [
    {
        "number": "FV/2026/07/041", "date": "2026-07-03", "due_date": "2026-07-17",
        "seller": ("Biuro Serwis Nowak Sp. z o.o.", "ul. Prosta 12, 00-013 Warszawa", "NIP 5252345678"),
        "buyer": ("Delta Handel S.A.", "al. Pokoju 5, 31-548 Kraków", "NIP 6772123456"),
        "line_items": [
            ("Papier ksero A4 80g (karton 5 ryz)", 14, 42.00, 23),
            ("Toner do drukarki HP 26A", 3, 289.00, 23),
            ("Segregator A4 75mm", 40, 8.90, 23),
        ],
    },
    {
        "number": "FV/2026/07/058", "date": "2026-07-11", "due_date": "2026-08-10",
        "seller": ("TransLog Kowalczyk", "ul. Magazynowa 3, 61-013 Poznań", "NIP 7811987654"),
        "buyer": ("Delta Handel S.A.", "al. Pokoju 5, 31-548 Kraków", "NIP 6772123456"),
        "line_items": [
            ("Usługa transportowa Poznań–Kraków", 2, 850.00, 23),
            ("Opłata paliwowa", 2, 95.00, 23),
        ],
    },
    {
        "number": "FV/2026/07/072", "date": "2026-07-18", "due_date": "2026-08-01",
        "seller": ("StudioNet Wiśniewska", "ul. Cicha 8/2, 50-432 Wrocław", "NIP 8992765432"),
        "buyer": ("Delta Handel S.A.", "al. Pokoju 5, 31-548 Kraków", "NIP 6772123456"),
        "line_items": [
            ("Utrzymanie strony WWW — lipiec 2026", 1, 1200.00, 23),
            ("Hosting + domena (rozliczenie roczne)", 1, 640.00, 23),
            ("Dodatkowe poprawki graficzne (rbh)", 6, 140.00, 23),
        ],
    },
]


def fmt_pln(x):
    return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " zł"


def build(invoice, path):
    doc = SimpleDocTemplate(str(path), pagesize=A4,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=16 * mm)
    H = ParagraphStyle("h", fontName=FONT + "-Bold", fontSize=15)
    N = ParagraphStyle("n", fontName=FONT, fontSize=9, leading=12)
    S = ParagraphStyle("s", fontName=FONT, fontSize=8.2, leading=11, textColor=colors.HexColor("#555555"))

    elements = [Paragraph(f"Faktura VAT {invoice['number']}", H), Spacer(1, 4),
                Paragraph(f"Data wystawienia: {invoice['date']} &nbsp;&nbsp;|&nbsp;&nbsp; "
                          f"Termin płatności: {invoice['due_date']}", S),
                Spacer(1, 10)]

    parties = Table([[
        Paragraph(f"<b>Sprzedawca</b><br/>{invoice['seller'][0]}<br/>{invoice['seller'][1]}<br/>{invoice['seller'][2]}", N),
        Paragraph(f"<b>Nabywca</b><br/>{invoice['buyer'][0]}<br/>{invoice['buyer'][1]}<br/>{invoice['buyer'][2]}", N),
    ]], colWidths=[88 * mm, 88 * mm])
    parties.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    elements += [parties, Spacer(1, 12)]

    rows = [["Lp.", "Nazwa towaru / usługi", "Ilość", "Cena netto", "VAT", "Wartość brutto"]]
    total_net = total_gross = 0.0
    for i, (name, qty, price, vat) in enumerate(invoice["line_items"], 1):
        net = qty * price
        gross = net * (1 + vat / 100)
        total_net += net
        total_gross += gross
        rows.append([str(i), name, str(qty), fmt_pln(price), f"{vat}%", fmt_pln(gross)])
    rows.append(["", "", "", "", "RAZEM:", fmt_pln(total_gross)])

    t = Table(rows, colWidths=[10 * mm, 78 * mm, 14 * mm, 26 * mm, 14 * mm, 32 * mm])
    t.setStyle(TableStyle([
        ("FONT", (0, 0), (-1, -1), FONT, 8.6),
        ("FONT", (0, 0), (-1, 0), FONT + "-Bold", 8.6),
        ("FONT", (-2, -1), (-1, -1), FONT + "-Bold", 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f8")),
        ("GRID", (0, 0), (-1, -2), 0.4, colors.HexColor("#c9ccd4")),
        ("LINEBELOW", (-2, -1), (-1, -1), 0.8, colors.HexColor("#333333")),
        ("ALIGN", (2, 1), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    elements += [t, Spacer(1, 10),
                 Paragraph(f"Do zapłaty: <b>{fmt_pln(total_gross)}</b> &nbsp;(netto: {fmt_pln(total_net)})", N),
                 Spacer(1, 6),
                 Paragraph("Dokument wygenerowany na potrzeby demonstracji — wszystkie dane są fikcyjne.", S)]
    doc.build(elements)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for invoice in INVOICES:
        name = invoice["number"].replace("/", "_") + ".pdf"
        build(invoice, OUT / name)
        print("OK", name)


if __name__ == "__main__":
    main()
