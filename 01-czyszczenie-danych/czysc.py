#!/usr/bin/env python3
"""
Czyszczenie i standaryzacja bazy klientów (demo „dostarcz-i-koniec").
Wejście:  dane_wejsciowe/klienci_brudne.csv  (średnik jako separator)
Wyjście:  wyniki/klienci_czyste.xlsx         (gotowa, uporządkowana baza)
          wyniki/raport_zmian.md             (co zostało poprawione i ile)

Uruchomienie:  python3 czysc.py
Zero zależności od systemów klienta — wynik to samodzielny plik.
"""
import re
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd

BAZA = Path(__file__).resolve().parent
WEJSCIE = BAZA / "dane_wejsciowe" / "klienci_brudne.csv"
WYJSCIE_XLSX = BAZA / "wyniki" / "klienci_czyste.xlsx"
RAPORT = BAZA / "wyniki" / "raport_zmian.md"

MIESIACE = {
    "stycznia": 1, "lutego": 2, "marca": 3, "kwietnia": 4, "maja": 5, "czerwca": 6,
    "lipca": 7, "sierpnia": 8, "wrzesnia": 9, "września": 9, "pazdziernika": 10,
    "października": 10, "listopada": 11, "grudnia": 12,
}
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[a-z]{2,}$", re.I)


def parsuj_date(s):
    """Sprowadza różne formaty daty do ISO (YYYY-MM-DD). Zwraca None, jeśli się nie da."""
    if not s or not str(s).strip():
        return None
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            pass
    m = re.match(r"(\d{1,2})\s+([a-ząćęłńóśźż]+)\s+(\d{4})", s, re.I)
    if m and m.group(2).lower() in MIESIACE:
        try:
            return datetime(int(m.group(3)), MIESIACE[m.group(2).lower()], int(m.group(1))).date().isoformat()
        except ValueError:
            return None
    return None


def parsuj_kwote(s):
    """'1 234,50 zł' -> 1234.50 (float). None, jeśli brak/nie liczba."""
    if s is None or not str(s).strip():
        return None
    s = str(s).lower().replace("zł", "").replace("pln", "").replace("\xa0", " ").strip()
    s = s.replace(" ", "")
    if "," in s and "." in s:          # 1.234,50 -> 1234.50
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def norm_telefon(s):
    """Do formatu +48XXXXXXXXX. None, jeśli nie 9 cyfr krajowych."""
    if not s or not str(s).strip():
        return None
    cyfry = re.sub(r"\D", "", str(s))
    if cyfry.startswith("48"):
        cyfry = cyfry[2:]
    if len(cyfry) == 9:
        return "+48" + cyfry
    return None


def norm_nip(s):
    """10 cyfr bez separatorów. None, jeśli nie 10 cyfr."""
    if not s or not str(s).strip():
        return None
    cyfry = re.sub(r"\D", "", str(s))
    return cyfry if len(cyfry) == 10 else None


def main():
    df = pd.read_csv(WEJSCIE, sep=";", dtype=str, keep_default_na=False)
    zmiany = {"wierszy_wejscie": len(df)}

    df.columns = [c.strip() for c in df.columns]
    for c in df.columns:
        df[c] = df[c].map(lambda x: x.strip() if isinstance(x, str) else x)

    # wiersze-śmieci: brak nazwiska lub niepoprawny e-mail
    df["email_norm"] = df["email"].str.lower()
    puste = df["imie i nazwisko"].eq("") | df["email_norm"].eq("")
    zly_email = ~df["email_norm"].apply(lambda e: bool(EMAIL_RE.match(e)) if e else False)
    smieci = puste | zly_email
    zmiany["usuniete_smieci"] = int(smieci.sum())
    df = df[~smieci].copy()

    df["data zamowienia"] = df["data zamowienia"].map(parsuj_date)
    df["kwota"] = df["kwota"].map(parsuj_kwote)
    df["telefon"] = df["telefon"].map(norm_telefon)
    df["NIP"] = df["NIP"].map(norm_nip)
    df["miasto"] = df["miasto"].str.title()
    df = df.rename(columns={"email_norm": "email_czysty"})
    df["email"] = df["email_czysty"]
    df = df.drop(columns=["email_czysty"])

    # duplikaty: ten sam e-mail + data + kwota (ta sama transakcja wpisana 2x)
    przed = len(df)
    df = df.drop_duplicates(subset=["email", "data zamowienia", "kwota"], keep="first")
    zmiany["usuniete_duplikaty"] = przed - len(df)

    zmiany["braki_daty"] = int(df["data zamowienia"].isna().sum())
    zmiany["braki_kwoty"] = int(df["kwota"].isna().sum())
    zmiany["braki_telefonu"] = int(df["telefon"].isna().sum())
    zmiany["braki_nip"] = int(df["NIP"].isna().sum())
    zmiany["wierszy_wyjscie"] = len(df)

    df = df.sort_values(["data zamowienia", "imie i nazwisko"], na_position="last")
    WYJSCIE_XLSX.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(WYJSCIE_XLSX, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Klienci")
        ark = w.sheets["Klienci"]
        for kol in ark.columns:
            szer = max((len(str(c.value)) for c in kol if c.value is not None), default=10)
            ark.column_dimensions[kol[0].column_letter].width = min(szer + 2, 40)

    suma = df["kwota"].dropna().sum()
    n_kwot = int(df["kwota"].notna().sum())
    srednia = suma / n_kwot if n_kwot else 0.0

    def pln(x):  # 1234.5 -> "1 234,50"
        return f"{x:,.2f}".replace(",", " ").replace(".", ",")

    RAPORT.write_text(
        f"""# Raport czyszczenia bazy klientów

Wejście: `{WEJSCIE.name}` — **{zmiany['wierszy_wejscie']} wierszy**
Wyjście: `{WYJSCIE_XLSX.name}` — **{zmiany['wierszy_wyjscie']} czystych rekordów**

## Co zostało poprawione
- Usunięte wiersze-śmieci (brak nazwiska / niepoprawny e-mail): **{zmiany['usuniete_smieci']}**
- Usunięte duplikaty (ten sam e-mail + data + kwota): **{zmiany['usuniete_duplikaty']}**
- Daty sprowadzone do formatu ISO `RRRR-MM-DD` (z 5 różnych zapisów, w tym słownego „12 stycznia 2024")
- Kwoty sprowadzone do liczby (z „1 234,50 zł", „2200,00", „999.99" → 1234.5, 2200.0, 999.99)
- Telefony ujednolicone do `+48XXXXXXXXX`
- NIP-y do 10 cyfr bez myślników
- E-maile małymi literami, przycięte spacje; miasta z wielkiej litery

## Pozostałe braki do uzupełnienia po stronie klienta
- Brak daty: **{zmiany['braki_daty']}** | brak kwoty: **{zmiany['braki_kwoty']}** | brak telefonu: **{zmiany['braki_telefonu']}** | brak NIP: **{zmiany['braki_nip']}**

## Szybkie KPI z czystych danych
- Suma zamówień: **{pln(suma)} zł**
- Średnie zamówienie: **{pln(srednia)} zł** (z {n_kwot} rekordów z kwotą)
"""
    )
    print(f"OK: {zmiany['wierszy_wejscie']} -> {zmiany['wierszy_wyjscie']} rekordów")
    print(f"   usunięto: {zmiany['usuniete_smieci']} śmieci, {zmiany['usuniete_duplikaty']} duplikatów")
    print(f"   pliki: {WYJSCIE_XLSX.name}, {RAPORT.name}")


if __name__ == "__main__":
    sys.exit(main())
