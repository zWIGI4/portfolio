# Raport czyszczenia bazy klientów

Wejście: `klienci_brudne.csv` — **19 wierszy**
Wyjście: `klienci_czyste.xlsx` — **14 czystych rekordów**

## Co zostało poprawione
- Usunięte wiersze-śmieci (brak nazwiska / niepoprawny e-mail): **2**
- Usunięte duplikaty (ten sam e-mail + data + kwota): **3**
- Daty sprowadzone do formatu ISO `RRRR-MM-DD` (z 5 różnych zapisów, w tym słownego „12 stycznia 2024")
- Kwoty sprowadzone do liczby (z „1 234,50 zł", „2200,00", „999.99" → 1234.5, 2200.0, 999.99)
- Telefony ujednolicone do `+48XXXXXXXXX`
- NIP-y do 10 cyfr bez myślników
- E-maile małymi literami, przycięte spacje; miasta z wielkiej litery

## Pozostałe braki do uzupełnienia po stronie klienta
- Brak daty: **0** | brak kwoty: **2** | brak telefonu: **0** | brak NIP: **6**

## Szybkie KPI z czystych danych
- Suma zamówień: **30 825,98 zł**
- Średnie zamówienie: **2 568,83 zł** (z 12 rekordów z kwotą)
