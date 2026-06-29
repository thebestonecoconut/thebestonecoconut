# Konwerter etykiet BarTender (.btw) → tekst / PDF

Narzędzie `btw_convert.py` wyciąga **tekst** z plików etykiet BarTender (`.btw`)
i zapisuje go do pliku tekstowego (`.txt`) oraz do **PDF**.

## Ważne (przeczytaj)

Pliki `.btw` to dokumenty OLE2 tworzone przez program **BarTender**
(Seagull Scientific). To **format zamknięty (binarny)**. Pełne, wierne
odwzorowanie wyglądu etykiety (kody kreskowe, grafika, dokładny układ) potrafi
wygenerować **tylko sam BarTender** (lub jego dodatki). 

To narzędzie wyciąga **zawartość tekstową** etykiet — czyli napisy, składy,
kody EAN, ceny itp. — i układa je w czytelny plik tekstowy/PDF. To zwykle
wystarcza, gdy potrzebujesz treści, a nie graficznego podglądu.

> Jeśli potrzebujesz dokładnego graficznego PDF każdej etykiety, najlepiej w
> BarTenderze użyć: `Plik → Drukuj → drukarka "Microsoft Print to PDF"` albo
> eksportu/wydruku do PDF. To narzędzie jest do masowego wyciągania **tekstu**.

## Instalacja

```bash
python3 -m venv .venv
source .venv/bin/activate      # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Użycie

Jeden plik (utworzy `etykieta.txt` i `etykieta.pdf` obok źródła):

```bash
python btw_convert.py etykieta.btw
```

Cały folder (wszystkie `.btw`, także w podfolderach):

```bash
python btw_convert.py sciezka/do/folderu
```

Zapis wyników do osobnego folderu:

```bash
python btw_convert.py folder_z_etykietami/ -o wyniki/
```

Jeden wspólny PDF ze wszystkimi etykietami (po jednej na stronę):

```bash
python btw_convert.py folder_z_etykietami/ --combined-pdf wszystkie_etykiety.pdf
```

Tylko tekst (bez PDF) lub tylko PDF (bez tekstu):

```bash
python btw_convert.py etykieta.btw --txt-only
python btw_convert.py etykieta.btw --pdf-only
```

## Co dostajesz

- `*.txt` — czysty tekst, kodowanie UTF-8 (obsługa polskich znaków: ą, ć, ł, ó, ż…).
- `*.pdf` — czytelny PDF z listą tekstów z każdej etykiety.

## Uwagi

- Polskie znaki w PDF działają, jeśli w systemie jest czcionka DejaVu Sans
  (na Linuksie zwykle pakiet `fonts-dejavu`). Bez niej PDF użyje czcionki
  Helvetica i znaki spoza Latin-1 zostaną zastąpione.
- Jeśli z jakiejś etykiety nie da się wyciągnąć tekstu (np. cała etykieta to
  grafika), plik wynikowy będzie pusty lub bardzo krótki — to normalne dla
  takiego pliku.
