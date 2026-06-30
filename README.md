# thebestonecoconut

## Kopiowanie etykiet (.btw) wg listy produktow z Excela

Narzedzie wyciaga nazwy produktow z pliku **Excel** (`.xlsx`) lub **CSV**,
porownuje je z plikami (domyslnie `.btw` – etykiety BarTender) w wskazanym
folderze i **kopiuje dopasowane pliki** do osobnego folderu. Na koniec tworzy
**raport CSV**, ktory pokazuje:

- ktore produkty udalo sie dopasowac i skopiowac (`SKOPIOWANO`),
- ktore produkty nie maja pasujacego pliku (`BRAK_PLIKU`),
- ktore pliki w folderze nie pasowaly do zadnego produktu (`NIEDOPASOWANY_PLIK`),
- ewentualne bledy kopiowania (`BLAD_KOPIOWANIA`).

W konsoli wyniki sa kolorowane: zielone `[OK ]`, czerwone `[NIE]`.

> Skrypt **nie wymaga zainstalowanego MS Excel** – pliki `.xlsx` czyta
> bezposrednio. Dziala na zwyklym Windowsie (uzywa wbudowanego PowerShell).

### Pliki

- `kopiuj_etykiety.bat` – to klikasz / uruchamiasz.
- `kopiuj_etykiety.ps1` – wlasciwa logika (musi lezec obok pliku `.bat`).

### Najprostsze uzycie

1. Skopiuj oba pliki (`.bat` i `.ps1`) do folderu.
2. Wrzuc do tego samego folderu swoj plik `produkty.xlsx` (lub `.csv`)
   oraz pliki `.btw`.
3. Kliknij dwukrotnie **`kopiuj_etykiety.bat`**.
4. Dopasowane pliki znajdziesz w podfolderze `skopiowane\`, a obok – raport
   `raport_RRRRMMDD_GGMMSS.csv`.

Domyslnie nazwy produktow czytane sa z **kolumny A**, a **pierwszy wiersz**
traktowany jest jako naglowek (pomijany).

### Dopasowanie ROZMYTE (gdy nazwy sa podobne, nie identyczne)

Domyslnie skrypt uzywa **dopasowania rozmytego**, ktore radzi sobie z drobnymi
roznicami miedzy nazwa produktu a nazwa pliku:

- rozne separatory i odstepy (`Produkt Alfa` ↔ `produkt_alfa`, `Kawa 250g` ↔ `Kawa-250g`),
- polskie znaki / ich brak (`Maka pszenna` ↔ `maka_pszenna`),
- dopiski w nazwie pliku (`Kawa 250g` ↔ `Kawa-250g-ARABICA`),
- literowki, w tym przestawione litery (`Cukier` ↔ `Cukeir`).

Dla kazdego produktu liczone jest **podobienstwo w %** do najlepiej pasujacego
pliku. Plik zostaje skopiowany, gdy podobienstwo jest **>= progu** (domyslnie
**80%**). W konsoli i w raporcie widac uzyte podobienstwo oraz typ dopasowania
(`dokladne`, `zawiera`, `rozmyte`). Dla produktow bez dopasowania pokazywany
jest **najblizszy kandydat i jego %** - latwo wtedy ocenic, czy obnizyc prog.

Regulacja progu (6. parametr `.bat` lub `-Threshold` w PS):

- za **malo** trafien -> obniz prog, np. `65`,
- za **duzo** / mylne trafienia -> podnies prog, np. `90`,
- chcesz **tylko identyczne** nazwy -> uzyj `-ExactOnly`.

### Uzycie z parametrami (wiersz polecen)

```bat
kopiuj_etykiety.bat "C:\dane\produkty.xlsx" "C:\etykiety" "C:\wynik" btw A
```

Kolejnosc parametrow (wszystkie opcjonalne):

| Pozycja | Znaczenie                              | Domyslnie                |
|---------|----------------------------------------|--------------------------|
| 1       | plik Excel/CSV z nazwami produktow     | pierwszy `.xlsx`/`.csv` obok skryptu |
| 2       | folder z plikami `.btw`                | folder skryptu           |
| 3       | folder docelowy (kopie)                | `.\skopiowane`           |
| 4       | rozszerzenie plikow (bez kropki)       | `btw`                    |
| 5       | litera kolumny z nazwami (`A`, `B`...) | `A`                      |
| 6       | prog podobienstwa 0-100 (dopas. rozmyte) | `80`                   |

Przyklad z luzniejszym dopasowaniem (prog 65%):

```bat
kopiuj_etykiety.bat "C:\dane\produkty.xlsx" "C:\etykiety" "C:\wynik" btw A 65
```

Mozesz tez przeciagnac plik Excel na ikone `.bat`.

### Zaawansowane opcje (uruchomienie bezposrednio skryptu PowerShell)

```powershell
powershell -ExecutionPolicy Bypass -File .\kopiuj_etykiety.ps1 `
    -ExcelPath ".\produkty.xlsx" `
    -SourceFolder ".\etykiety" `
    -TargetFolder ".\skopiowane" `
    -Column "A" `
    -HasHeader $true `
    -Threshold 80 `   # prog podobienstwa dla dopasowania rozmytego (0-100)
    -ExactOnly `      # tylko dokladne dopasowanie (po normalizacji nazw)
    -Recurse          # przeszukuj rowniez podfoldery
```

- `-Threshold <0-100>` – minimalne podobienstwo do uznania za dopasowanie
  (domyslnie 80).
- `-ExactOnly` – wymusza dopasowanie **dokladne** (po normalizacji nazw:
  bez wielkosci liter, bez polskich znakow, bez spacji/podkreslen/myslnikow).
- `-Recurse` – szuka plikow rowniez w podfolderach.

### Jak liczone jest dopasowanie

Porownywana jest **nazwa pliku bez rozszerzenia** z nazwa produktu, po
normalizacji obu (male litery, usuniete polskie znaki oraz znaki typu spacja,
`_`, `-`, `.`). Nastepnie liczone jest podobienstwo:

1. **dokladne** (100%) – nazwy identyczne po normalizacji,
2. **zawiera** (90–100%) – jedna nazwa zawiera sie w drugiej (np. dopisek
   rozmiaru/wersji w nazwie pliku),
3. **rozmyte** – podobienstwo na podstawie odleglosci edycyjnej
   (Damerau-Levenshtein), ktora liczy wstawienia, usuniecia, zamiany i
   przestawienia sasiednich liter (literowki).
