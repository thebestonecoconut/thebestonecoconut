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

Domyslnie skrypt uzywa **dopasowania rozmytego**, ktore radzi sobie z roznicami
miedzy nazwa produktu a nazwa pliku, np. (prawdziwe przyklady):

| Nazwa w Excelu                          | Nazwa pliku                                                  |
|-----------------------------------------|--------------------------------------------------------------|
| `Danmis Jogurt kozi jagodowy 125g`      | `Danmis - jogurt kozi jagodowy 125g.btw`                     |
| `Danmis Ser kozi termizowany wanilia 100g` | `Danmis - ser kozi termizowany waniliowy 100g K.btw`      |
| `Danmis Ser kozi twarogowy 200g`        | `Danmis - ser kozi twarogowy (pełnotłusty) 200g.btw`        |
| `Danmis Ser Kozi kanapkowy 150g`        | `Danmis - twaróg kozi kanapkowy 150g.btw`                   |
| `Danmis Ser Kozi wędzony plastry 100g`  | `Danmis N - ser kozi twardy wędzony ... w plastrach 100g (5160).btw` |

Radzi sobie m.in. z:

- myslnikami, dopiskami w nazwie pliku (` K`, kody w nawiasach `(1910)`, `(5160)`),
- innymi separatorami i odstepami, polskimi znakami / ich brakiem,
- inna kolejnoscia / zmiana slow (`Ser kozi` ↔ `twaróg kozi`, `twarogowy` ↔ `twaróg`),
- odmiana i koncowkami (`plastry` ↔ `plastrach`, `wanilia` ↔ `waniliowy`),
- literowkami, w tym przestawionymi literami (`Cukier` ↔ `Cukeir`).

**Jak to dziala:** dla kazdego produktu liczone jest **podobienstwo w %** do
najlepiej pasujacego pliku - laczone z dwoch metod: porownania calej nazwy oraz
porownania **slowo po slowie** (mniej istotne slowa, jak pojedyncze litery czy
kody liczbowe, maja mniejsza wage; nadmiarowe slowa w nazwie pliku obnizaja
wynik). Plik jest kopiowany, gdy podobienstwo jest **>= progu** (domyslnie
**75%**). Typ dopasowania w raporcie: `dokladne`, `zawiera`, `slowa`, `rozmyte`.

Dla produktow **bez dopasowania** raport i konsola pokazuja **najblizszego
kandydata i jego %** - latwo ocenic, czy to faktycznie brak pliku, czy trzeba
obnizyc prog.

Regulacja progu (6. parametr `.bat` lub `-Threshold` w PS):

- za **malo** trafien -> obniz prog, np. `65`,
- za **duzo** / mylne trafienia -> podnies prog, np. `85`,
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
| 6       | prog podobienstwa 0-100 (dopas. rozmyte) | `75`                   |

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
    -Threshold 75 `   # prog podobienstwa dla dopasowania rozmytego (0-100)
    -ExactOnly `      # tylko dokladne dopasowanie (po normalizacji nazw)
    -Recurse          # przeszukuj rowniez podfoldery
```

- `-Threshold <0-100>` – minimalne podobienstwo do uznania za dopasowanie
  (domyslnie 75).
- `-ExactOnly` – wymusza dopasowanie **dokladne** (po normalizacji nazw:
  bez wielkosci liter, bez polskich znakow, bez spacji/podkreslen/myslnikow).
- `-Recurse` – szuka plikow rowniez w podfolderach.

### Jak liczone jest dopasowanie

Porownywana jest **nazwa pliku bez rozszerzenia** z nazwa produktu. Liczone sa
dwie metody, a brany jest **lepszy** wynik:

1. **Cala nazwa** (po normalizacji: male litery, bez polskich znakow, bez spacji
   `_` `-` `.`):
   - `dokladne` (100%) – identyczne,
   - `zawiera` – jedna zawiera sie w drugiej,
   - `rozmyte` – odleglosc edycyjna Damerau-Levenshtein (wstaw/usun/zamien/
     przestaw sasiednie litery), dobra na literowki.
2. **Slowo po slowie** (`slowa`) – jaka czesc (wazona) slow produktu wystepuje w
   nazwie pliku, z tolerancja na koncowki/odmiane. Pojedyncze litery i kody
   liczbowe maja mniejsza wage, a nadmiarowe slowa w nazwie pliku obnizaja
   wynik. Ta metoda najlepiej radzi sobie z dopiskami i inna kolejnoscia slow.

Dopasowanie identyczne (100%) zawsze ma pierwszenstwo, wiec gdy istnieje plik
o dokladnie tej nazwie, nie sa dublowane kopie wariantow z dopiskami.
