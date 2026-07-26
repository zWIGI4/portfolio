#!/usr/bin/env python3
"""
Generator przykładowych faktur PDF (dane w 100% fikcyjne) — na potrzeby demo.
Tworzy 3 faktury o tym samym układzie, różnych danych, w dane_wejsciowe/.
"""
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle

BAZA = Path(__file__).resolve().parent
OUT = BAZA / "dane_wejsciowe"

# DejaVu ma polskie znaki (jest w systemie z matplotlib)
FONT = "DejaVu"
pdfmetrics.registerFont(TTFont(FONT, "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont(FONT + "-Bold", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"))

FAKTURY = [
    {
        "nr": "FV/2026/07/041", "data": "2026-07-03", "termin": "2026-07-17",
        "sprzedawca": ("Biuro Serwis Nowak Sp. z o.o.", "ul. Prosta 12, 00-013 Warszawa", "NIP 5252345678"),
        "nabywca": ("Delta Handel S.A.", "al. Pokoju 5, 31-548 Kraków", "NIP 6772123456"),
        "pozycje": [
            ("Papier ksero A4 80g (karton 5 ryz)", 14, 42.00, 23),
            ("Toner do drukarki HP 26A", 3, 289.00, 23),
            ("Segregator A4 75mm", 40, 8.90, 23),
        ],
    },
    {
        "nr": "FV/2026/07/058", "data": "2026-07-11", "termin": "2026-08-10",
        "sprzedawca": ("TransLog Kowalczyk", "ul. Magazynowa 3, 61-013 Poznań", "NIP 7811987654"),
        "nabywca": ("Delta Handel S.A.", "al. Pokoju 5, 31-548 Kraków", "NIP 6772123456"),
        "pozycje": [
            ("Usługa transportowa Poznań–Kraków", 2, 850.00, 23),
            ("Opłata paliwowa", 2, 95.00, 23),
        ],
    },
    {
        "nr": "FV/2026/07/072", "data": "2026-07-18", "termin": "2026-08-01",
        "sprzedawca": ("StudioNet Wiśniewska", "ul. Cicha 8/2, 50-432 Wrocław", "NIP 8992765432"),
        "nabywca": ("Delta Handel S.A.", "al. Pokoju 5, 31-548 Kraków", "NIP 6772123456"),
        "pozycje": [
            ("Utrzymanie strony WWW — lipiec 2026", 1, 1200.00, 23),
            ("Hosting + domena (rozliczenie roczne)", 1, 640.00, 23),
            ("Dodatkowe poprawki graficzne (rbh)", 6, 140.00, 23),
        ],
    },
]


def zl(x):
    return f"{x:,.2f}".replace(",", " ").replace(".", ",") + " zł"


def buduj(f, sciezka):
    doc = SimpleDocTemplate(str(sciezka), pagesize=A4,
                            leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=16 * mm)
    H = ParagraphStyle("h", fontName=FONT + "-Bold", fontSize=15)
    N = ParagraphStyle("n", fontName=FONT, fontSize=9, leading=12)
    S = ParagraphStyle("s", fontName=FONT, fontSize=8.2, leading=11, textColor=colors.HexColor("#555555"))

    el = [Paragraph(f"Faktura VAT {f['nr']}", H), Spacer(1, 4),
          Paragraph(f"Data wystawienia: {f['data']} &nbsp;&nbsp;|&nbsp;&nbsp; Termin płatności: {f['termin']}", S),
          Spacer(1, 10)]

    strony = Table([[
        Paragraph(f"<b>Sprzedawca</b><br/>{f['sprzedawca'][0]}<br/>{f['sprzedawca'][1]}<br/>{f['sprzedawca'][2]}", N),
        Paragraph(f"<b>Nabywca</b><br/>{f['nabywca'][0]}<br/>{f['nabywca'][1]}<br/>{f['nabywca'][2]}", N),
    ]], colWidths=[88 * mm, 88 * mm])
    strony.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    el += [strony, Spacer(1, 12)]

    wiersze = [["Lp.", "Nazwa towaru / usługi", "Ilość", "Cena netto", "VAT", "Wartość brutto"]]
    suma_netto = suma_brutto = 0.0
    for i, (nazwa, ilosc, cena, vat) in enumerate(f["pozycje"], 1):
        netto = ilosc * cena
        brutto = netto * (1 + vat / 100)
        suma_netto += netto
        suma_brutto += brutto
        wiersze.append([str(i), nazwa, str(ilosc), zl(cena), f"{vat}%", zl(brutto)])
    wiersze.append(["", "", "", "", "RAZEM:", zl(suma_brutto)])

    t = Table(wiersze, colWidths=[10 * mm, 78 * mm, 14 * mm, 26 * mm, 14 * mm, 32 * mm])
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
    el += [t, Spacer(1, 10),
           Paragraph(f"Do zapłaty: <b>{zl(suma_brutto)}</b> &nbsp;(netto: {zl(suma_netto)})", N),
           Spacer(1, 6),
           Paragraph("Dokument wygenerowany na potrzeby demonstracji — wszystkie dane są fikcyjne.", S)]
    doc.build(el)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for f in FAKTURY:
        nazwa = f["nr"].replace("/", "_") + ".pdf"
        buduj(f, OUT / nazwa)
        print("OK", nazwa)


if __name__ == "__main__":
    main()
