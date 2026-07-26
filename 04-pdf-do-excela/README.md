# 04 — Faktury PDF → rejestr w Excelu

*(EN: PDF invoices → searchable Excel register — headers + all line items, no retyping.)*

![Faktury PDF → rejestr w Excelu](../docs/img/pdf.png)

**Problem:** faktury przychodzą jako PDF-y. Ktoś co miesiąc ręcznie przepisuje numery,
daty, NIP-y i kwoty do arkusza — wolno i z błędami.

**Rozwiązanie:** skrypt czyta wszystkie PDF-y z folderu i buduje jeden rejestr Excel:
arkusz **Faktury** (nr, daty, sprzedawca, NIP-y, netto/brutto) + arkusz **Pozycje**
(każda linia z każdej faktury, z filtrem).

## Wejście → wyjście
- `dane_wejsciowe/*.pdf` — 3 przykładowe faktury (dane w 100% fikcyjne, generowane przez `generuj_faktury.py`)
- `wyniki/rejestr_faktur.xlsx` — rejestr: 3 faktury, 8 pozycji, suma brutto 7 848,63 zł

## Co potrafi parser
- pola nagłówkowe regexem z tekstu (nr faktury, daty, NIP-y, kwoty „do zapłaty")
- **sprzedawca/nabywca z układu dwukolumnowego po współrzędnych** (`within_bbox` —
  klasyczny przypadek, w którym zwykłe `extract_text` skleja kolumny)
- tabela pozycji przez `extract_tables`, kwoty sprowadzone do liczb
- kwoty w Excelu są liczbami — od razu się sumują i filtrują

## Uruchomienie
```bash
python3 generuj_faktury.py   # wygeneruj przykładowe PDF-y (albo podłóż swoje)
python3 ekstrakcja.py        # zbuduj rejestr wyniki/rejestr_faktur.xlsx
```

## W realnym zleceniu
Każdy wystawca ma inny szablon — parser dopasowuję do **Twoich** faktur (2–3 przykłady
wystarczą do kalibracji). Efekt to zawsze ten sam czysty rejestr.

Technologie: Python, pdfplumber, openpyxl (+ reportlab do generowania przykładów).
