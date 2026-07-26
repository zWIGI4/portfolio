# 03 — Dashboard sprzedaży z surowego CSV

*(EN: raw transaction export → formatted Excel dashboard with KPIs and charts.)*

![Dashboard sprzedaży](../docs/img/dashboard.png)

**Problem:** masz eksport transakcji (setki wierszy), ale zero wglądu — ile
przychodu, który miesiąc dowozi, która kategoria ciągnie wynik. Ręczne tabele
przestawne co miesiąc od nowa.

**Rozwiązanie:** jeden skrypt zamienia surowy CSV w gotowy, sformatowany
dashboard Excel. Podmieniasz plik wejściowy — dostajesz świeży raport.

## Wejście → wyjście
- `dane_wejsciowe/sprzedaz.csv` — 260 fikcyjnych transakcji (2025)
- `wyniki/dashboard.xlsx` — 4 arkusze:
  - **KPI** — przychód, liczba zamówień, średnia wartość, top kategoria
  - **Wg miesiąca** — tabela + natywny wykres słupkowy Excela
  - **Wg kategorii** — tabela z udziałem % + wykres
  - **Dane** — surowe transakcje z filtrem i zamrożonym nagłówkiem

## Dlaczego to wygodne
- wykresy są **natywne excelowe** — klient może je edytować jak własne
- liczby są liczbami (formaty `# ##0`), wszystko się sumuje i filtruje
- kolejny miesiąc = podmiana CSV i jedno uruchomienie

## Uruchomienie
```bash
python3 dashboard.py
```

## W realnym zleceniu
Działa na dowolnym eksporcie tabelarycznym (sklep, subiekt, bank, CRM) — mapuję
kolumny pod Twój format i ustalamy, jakie KPI/wykresy mają być na górze.

Technologie: Python, openpyxl (formatowanie + wykresy natywne w Excelu).
