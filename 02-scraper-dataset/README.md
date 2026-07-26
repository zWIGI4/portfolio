# 02 — Web scraping: strona WWW → gotowy zbiór danych

*(EN: website listings → clean XLSX/CSV dataset, collected politely.)*

![Strona WWW → gotowy dataset](../docs/img/scraper.png)

**Problem:** dane, których potrzebujesz (oferty, produkty, ceny, katalogi), są na
stronie WWW — a Ty potrzebujesz ich w Excelu do analizy, porównania cen albo
importu. Ręczne przepisywanie to godziny i literówki.

**Rozwiązanie:** skrypt przechodzi po stronach katalogu, wyciąga pola i oddaje
gotowy zbiór w XLSX i CSV — ceny jako liczby, dane od razu do roboty.

## Źródło demo
`books.toscrape.com` — publiczny sandbox stworzony **do** nauki scrapingu (brak
ograniczeń ToS). W realnym zleceniu podmieniam źródło i zestaw pól pod klienta.

## Wejście → wyjście
- wejście: strony katalogu (domyślnie 3, `--strony N` po więcej)
- `wyniki/produkty.xlsx` i `wyniki/produkty.csv` — 60 rekordów
- pola: `tytuł, cena (liczba), waluta, ocena 1–5, dostępność, url`

## Dobre praktyki w kodzie
- nagłówek `User-Agent` i opóźnienie między żądaniami (grzeczny scraping)
- wymuszone UTF-8 i odporne parsowanie ceny (regex) — bez „Â£51.77"
- obsługa końca paginacji i błędów HTTP
- sformatowany Excel: pogrubiony nagłówek, zamrożony wiersz, dopasowane kolumny

## Uruchomienie
```bash
python3 scraper.py            # 3 strony
python3 scraper.py --strony 5 # więcej
```

## W realnym zleceniu
Podajesz adres strony i listę pól — dostajesz dataset + skrypt do ponownego
uruchamiania. Scrapuję wyłącznie źródła, które na to pozwalają (robots.txt/ToS).

Technologie: Python, requests, BeautifulSoup, openpyxl.
