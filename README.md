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

Mozesz tez przeciagnac plik Excel na ikone `.bat`.

### Zaawansowane opcje (uruchomienie bezposrednio skryptu PowerShell)

```powershell
powershell -ExecutionPolicy Bypass -File .\kopiuj_etykiety.ps1 `
    -ExcelPath ".\produkty.xlsx" `
    -SourceFolder ".\etykiety" `
    -TargetFolder ".\skopiowane" `
    -Column "A" `
    -HasHeader $true `
    -Partial `        # dopasowanie czesciowe (nazwa zawiera sie w nazwie pliku)
    -Recurse          # przeszukuj rowniez podfoldery
```

- `-Partial` – plik pasuje, gdy nazwa produktu zawiera sie w nazwie pliku
  (lub odwrotnie). Bez tej opcji wymagane jest dopasowanie **dokladne**
  (porownanie nie rozroznia wielkosci liter i ignoruje spacje na koncach).
- `-Recurse` – szuka plikow rowniez w podfolderach.

### Dopasowywanie nazw

Porownywana jest **nazwa pliku bez rozszerzenia** z nazwa produktu. Np. produkt
`Produkt Alfa` pasuje do pliku `Produkt Alfa.btw`. Porownanie ignoruje wielkosc
liter oraz spacje na poczatku/koncu.
