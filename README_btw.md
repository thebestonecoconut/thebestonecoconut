# Konwerter etykiet BarTender (.btw) → tekst (.txt)

Wyciąga **tekst** z plików etykiet BarTender (`.btw`) do pliku tekstowego (`.txt`).

Dostępne są dwie wersje:

| Wersja | Plik | Wymaga instalacji? | Kiedy używać |
|--------|------|--------------------|--------------|
| **BAT (Windows)** | `btw2pdf.bat` | **NIE** (wbudowany PowerShell) | Najprościej — klik / przeciągnij |
| Python | `btw_convert.py` | tak (`olefile`) | Gdy wolisz Pythona / inny system |

## Ważne o formacie .btw

Pliki `.btw` to dokumenty OLE2 z programu **BarTender** (Seagull Scientific) —
format zamknięty, binarny. Narzędzie wyciąga z nich **zawartość tekstową**
(nazwy, opisy, składy, wartości odżywcze, dystrybutor). Wierne odwzorowanie
grafiki etykiety (kody kreskowe, układ) potrafi tylko sam BarTender.

Co obsługuje ekstrakcja tekstu:

- usuwanie osadzonego podglądu **PNG** (główne źródło śmieci),
- **rozpakowywanie** skompresowanych (zlib/deflate) fragmentów — tam jest treść,
- konwersję **RTF → czysty tekst** (dekodowanie `\uN`/`\'xx`, tabele, usuwanie formatowania),
- tekst w **UTF‑16** oraz **UTF‑8 / ASCII** (wieloliniowy),
- **polskie i niemieckie znaki** (ą, ć, ę, ł, ó, ä, ö, ü, ß…),
- treść w **XML** (znaczniki usuwane, encje `&amp;` → `&`),
- odfiltrowanie śmieci binarnych i wewnętrznych nazw obiektów BarTendera.

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

Obok każdej etykiety powstanie plik `.txt`.

### Opcje

```bat
btw2pdf.bat --raw "C:\folder"     :: WSZYSTKIE znalezione ciągi (bez filtrowania)
```

> **Weryfikacja poprawności:** jeśli masz wrażenie, że w `.txt` czegoś brakuje,
> uruchom z opcją `--raw`. Zobaczysz wtedy wszystko, co da się odczytać z pliku.

---

## Wersja Python

```bash
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

```bash
python btw_convert.py etykieta.btw              # -> etykieta.txt
python btw_convert.py folder/                   # cały folder (rekurencyjnie)
python btw_convert.py folder/ -o wyniki/        # zapis do innego folderu
python btw_convert.py folder/ --raw             # wszystkie ciągi (weryfikacja)
```

---

## Co dostajesz

- `*.txt` — czysty tekst, kodowanie **UTF‑8** (polskie i niemieckie znaki działają).

## Uwaga

Najlepszy efekt uzyskasz, podsyłając jeden przykładowy plik `.btw` — wtedy można
porównać wynik z `--raw` i precyzyjnie dostroić filtry.
