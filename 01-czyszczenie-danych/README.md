# 01 — Czyszczenie danych: brudny CSV → gotowy Excel

*(EN: messy multi-source CSV → clean, analysis-ready Excel + change report.)*

![Czyszczenie danych — przed i po](../docs/img/cleaning.png)

**Problem:** eksport z kilku źródeł to rozjechana baza — daty w 5 formatach (w tym
słowne „12 stycznia 2024"), kwoty jako „1 234,50 zł", zdublowane transakcje,
wiersze-śmieci, telefony i NIP-y w losowych zapisach. Nie da się na tym liczyć,
filtrować ani niczego importować.

**Rozwiązanie:** jeden skrypt robi porządek i oddaje dwa pliki — czysty arkusz
oraz raport z dokładną listą poprawek (co usunięto i dlaczego).

## Wejście → wyjście
- `dane_wejsciowe/klienci_brudne.csv` — 19 wierszy typowego bałaganu
- `wyniki/klienci_czyste.xlsx` — 14 czystych, posortowanych rekordów
- `wyniki/raport_zmian.md` — co poprawiono + szybkie KPI z czystych danych

## Co konkretnie robi
- daty → ISO `RRRR-MM-DD` (5 różnych zapisów, także słowne)
- kwoty → liczby (`1 234,50 zł` → `1234.5`) — od razu się sumują
- telefony → `+48XXXXXXXXX`, NIP-y → 10 cyfr bez separatorów
- e-maile małymi literami, miasta z wielkiej, przycięte spacje
- usuwa duplikaty transakcji (e-mail + data + kwota) i wiersze bez nazwiska
  / z niepoprawnym e-mailem — z raportem, ile i czego

## Uruchomienie
```bash
python3 czysc.py
```

## W realnym zleceniu
Podsyłasz swój plik (albo zanonimizowaną próbkę) + 2 zdania, jak ma wyglądać
wynik. Reguły czyszczenia dopasowuję do Twoich danych; dostajesz czysty plik,
raport zmian i skrypt — możesz odpalać sam przy kolejnych eksportach.

Technologie: Python, pandas, openpyxl.
