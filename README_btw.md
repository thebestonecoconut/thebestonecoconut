# Konwerter etykiet BarTender (.btw) → tekst (i PDF jako dodatek)

Wyciąga **tekst** z plików etykiet BarTender (`.btw`) do pliku tekstowego
(`.txt`). PDF jest tylko dodatkiem — **najważniejszy jest poprawny plik `.txt`**.

Dostępne są dwie wersje (robią to samo z tekstem):

| Wersja | Plik | Wymaga instalacji? | Kiedy używać |
|--------|------|--------------------|--------------|
| **BAT (Windows)** | `btw2pdf.bat` | **NIE** (PowerShell + Edge są w Windows 10/11) | Najprościej — klik / przeciągnij |
| Python | `btw_convert.py` | tak (`olefile`, `fpdf2`) | Gdy wolisz Pythona / inny system |

## Ważne o formacie .btw

Pliki `.btw` to dokumenty OLE2 z programu **BarTender** (Seagull Scientific) —
format zamknięty, binarny. Narzędzie wyciąga z nich **zawartość tekstową**
(nazwy, składy, kody EAN, ceny). Wierne odwzorowanie grafiki etykiety
(kody kreskowe, układ) potrafi tylko sam BarTender.

Co obsługuje ekstrakcja tekstu:

- tekst w **UTF‑16** (najczęstszy w BarTenderze) oraz **UTF‑8 / ASCII**,
- **polskie znaki** (ą, ć, ę, ł, ń, ó, ś, ź, ż) — zapis w UTF‑8,
- treść „zatopioną" w **XML** (znaczniki są usuwane, encje `&amp;` → `&`),
- odfiltrowanie typowych śmieci (kody hex, GUID-y, wewnętrzne nazwy BarTendera).

---

## Wersja BAT (Windows, bez instalacji)

1. Skopiuj plik `btw2pdf.bat` do folderu z etykietami (albo gdziekolwiek).
2. Uruchom na jeden z trzech sposobów:
   - **przeciągnij** plik `.btw` albo cały **folder** na ikonę `btw2pdf.bat`,
   - **dwuklik** — przetworzy wszystkie `.btw` w folderze, w którym leży `.bat`,
   - z wiersza poleceń:
     ```bat
     btw2pdf.bat "C:\sciezka\do\folderu"
     ```

Obok każdej etykiety powstanie plik `.txt` (oraz `.pdf`, jeśli dostępny Edge).

### Opcje

```bat
btw2pdf.bat --txt "C:\folder"     :: tylko pliki tekstowe (bez PDF)
btw2pdf.bat --raw "C:\folder"     :: WSZYSTKIE znalezione ciągi (bez filtrowania)
```

> **Weryfikacja poprawności:** jeśli masz wrażenie, że w `.txt` czegoś brakuje,
> uruchom z opcją `--raw`. Zobaczysz wtedy wszystko, co da się odczytać z pliku —
> dzięki temu można sprawdzić, że żadna treść nie ginie, i ewentualnie dostroić
> filtry.

---

## Wersja Python

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

```bash
python btw_convert.py etykieta.btw              # -> etykieta.txt + etykieta.pdf
python btw_convert.py folder/                   # cały folder (rekurencyjnie)
python btw_convert.py folder/ -o wyniki/        # zapis do innego folderu
python btw_convert.py folder/ --txt-only        # tylko tekst
python btw_convert.py folder/ --raw --txt-only  # wszystkie ciągi (weryfikacja)
python btw_convert.py folder/ --combined-pdf wszystkie.pdf
```

---

## Co dostajesz

- `*.txt` — czysty tekst, kodowanie **UTF‑8** (polskie znaki działają).
- `*.pdf` *(dodatek)* — czytelny PDF z treścią etykiety.
  W wersji BAT PDF tworzy wbudowany Edge; gdy go nie ma, zapisywany jest plik
  `.html`, który można wydrukować do PDF (Ctrl+P → „Microsoft Print to PDF").

## Uwaga

Najlepszy efekt uzyskasz, podsyłając jeden przykładowy plik `.btw` — wtedy można
porównać `.txt` z `--raw` i precyzyjnie dostroić, co ma trafiać do wyniku.
