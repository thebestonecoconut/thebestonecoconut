<#
.SYNOPSIS
    Wyciaga nazwy produktow z pliku Excel (lub CSV), porownuje je z plikami
    (domyslnie .btw - etykiety BarTender) w podanym folderze i kopiuje
    dopasowane pliki do folderu docelowego. Na koniec tworzy raport
    pokazujacy, ktore produkty/pliki udalo sie dopasowac i skopiowac,
    a ktorych nie.

.UWAGI
    Skrypt NIE wymaga zainstalowanego MS Excel - pliki .xlsx czyta bezposrednio
    (rozpakowuje XML). Obsluguje tez .csv.

.PARAMETRY
    Patrz blok param() ponizej. Wszystkie maja sensowne wartosci domyslne,
    wiec mozna uruchomic skrypt bez argumentow (najlepiej przez plik .bat).
#>

[CmdletBinding()]
param(
    # Sciezka do pliku Excel (.xlsx) lub CSV z nazwami produktow.
    # Domyslnie: pierwszy plik .xlsx/.csv obok skryptu.
    [string]$ExcelPath = "",

    # Folder, w ktorym szukamy plikow do skopiowania. Domyslnie: folder skryptu.
    [string]$SourceFolder = "",

    # Folder docelowy, do ktorego trafiaja dopasowane pliki.
    [string]$TargetFolder = "",

    # Rozszerzenie plikow do porownania (bez kropki). Domyslnie btw.
    [string]$Extension = "btw",

    # Litera kolumny w Excelu, w ktorej sa nazwy produktow (A, B, C...).
    [string]$Column = "A",

    # Czy pierwszy wiersz to naglowek (zostanie pominiety).
    [bool]$HasHeader = $true,

    # Tryb tylko dokladnego dopasowania (po normalizacji nazw). Domyslnie
    # wlaczone jest dopasowanie rozmyte (fuzzy), ktore radzi sobie z drobnymi
    # roznicami w nazwach.
    [switch]$ExactOnly,

    # Minimalne podobienstwo (0-100) dla dopasowania rozmytego. Im wyzej,
    # tym ostrzej (mniej trafien, mniej pomylek). Domyslnie 75.
    [int]$Threshold = 75,

    # Czy przeszukiwac podfoldery w poszukiwaniu plikow.
    [switch]$Recurse
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function Write-Info  ($m) { Write-Host $m -ForegroundColor Cyan }
function Write-Ok    ($m) { Write-Host $m -ForegroundColor Green }
function Write-Bad   ($m) { Write-Host $m -ForegroundColor Red }
function Write-Warn2 ($m) { Write-Host $m -ForegroundColor Yellow }

# ---------------------------------------------------------------------------
# Normalizacja nazw - sprowadza nazwe do porownywalnej postaci:
#  - male litery,
#  - polskie znaki -> bez ogonkow (a-z),
#  - usuniecie spacji, podkreslen, mysmikow, kropek i innych nie-alfanum.
# Dzieki temu "Produkt_Alfa-01.btw" i "produkt alfa 01" sa traktowane podobnie.
# ---------------------------------------------------------------------------
# Zamienia polskie znaki diakrytyczne na podstawowe litery ( a-z). Uzywamy
# kodow Unicode zamiast literalnych znakow, aby skrypt dzialal niezaleznie od
# kodowania pliku (Windows PowerShell 5.1 czyta .ps1 w kodowaniu ANSI, przez co
# literalne polskie znaki UTF-8 moglyby zostac zle zinterpretowane).
function Remove-PolishDiacritics {
    param([string]$s)
    if ([string]::IsNullOrEmpty($s)) { return $s }
    $pairs = @(
        @([char]0x0105, 'a'),  # a z ogonkiem
        @([char]0x0107, 'c'),  # c z kreska
        @([char]0x0119, 'e'),  # e z ogonkiem
        @([char]0x0142, 'l'),  # l z kreska
        @([char]0x0144, 'n'),  # n z kreska
        @([char]0x00F3, 'o'),  # o z kreska
        @([char]0x015B, 's'),  # s z kreska
        @([char]0x017A, 'z'),  # z z kreska
        @([char]0x017C, 'z')   # z z kropka
    )
    foreach ($p in $pairs) { $s = $s.Replace([string]$p[0], $p[1]) }
    return $s
}

function Get-NormName {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $s = $s.ToLowerInvariant()
    $s = Remove-PolishDiacritics $s
    # usun wszystko poza literami i cyframi
    $s = ($s -replace '[^a-z0-9]', '')
    return $s
}

# Rozbija nazwe na slowa (tokeny) wraz z waga kazdego slowa. Mniej istotne
# fragmenty (pojedyncze litery typu "K"/"N", kody liczbowe jak "1910") maja
# nizsza wage, dzieki czemu nie psuja dopasowania.
function Get-Tokens {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return @() }
    $s = $s.ToLowerInvariant()
    $s = Remove-PolishDiacritics $s
    $parts = $s -split '[^a-z0-9]+' | Where-Object { $_ -ne "" }
    $result = @()
    foreach ($t in $parts) {
        $w = $t.Length
        if ($t.Length -le 1) { $w = 0.3 }                 # pojedyncze litery: szum
        elseif ($t -match '^[0-9]+$') { $w = 1.0 }        # czyste liczby: zwykle kody
        $result += [pscustomobject]@{ T = $t; W = [double]$w }
    }
    return $result
}

# Dopasowanie po slowach: jaka czesc (wazona) slow produktu wystepuje w nazwie
# pliku, z lekka tolerancja na literowki/koncowki i kara za nadmiarowe slowa
# w nazwie pliku.
function Get-TokenScore {
    param($prodTokens, $fileTokens)
    if (-not $prodTokens -or $prodTokens.Count -eq 0) { return 0 }
    if (-not $fileTokens -or $fileTokens.Count -eq 0) { return 0 }

    $used = New-Object 'bool[]' ($fileTokens.Count)
    $totalW = 0.0
    $matchedW = 0.0
    foreach ($pt in $prodTokens) {
        $totalW += $pt.W
        $bestSim = 0.0
        $bestIdx = -1
        for ($i = 0; $i -lt $fileTokens.Count; $i++) {
            if ($used[$i]) { continue }
            $ft = $fileTokens[$i]
            if ($pt.T -eq $ft.T) { $sim = 1.0 }
            else {
                $dist = Get-Levenshtein $pt.T $ft.T
                $ml = [Math]::Max($pt.T.Length, $ft.T.Length)
                $sim = 1.0 - ($dist / $ml)
            }
            if ($sim -gt $bestSim) { $bestSim = $sim; $bestIdx = $i }
        }
        # prog 0.6 toleruje polskie odmiany/koncowki (plastry<->plastrach)
        if ($bestSim -ge 0.6 -and $bestIdx -ge 0) {
            $used[$bestIdx] = $true
            $matchedW += $pt.W * $bestSim
        }
    }
    # waga slow pliku, ktore nie zostaly wykorzystane (nadmiar w nazwie pliku)
    $extraW = 0.0
    for ($i = 0; $i -lt $fileTokens.Count; $i++) {
        if (-not $used[$i]) { $extraW += $fileTokens[$i].W }
    }
    $denom = $totalW + 0.3 * $extraW
    if ($denom -le 0) { return 0 }
    $score = 100.0 * $matchedW / $denom
    # ograniczamy do 99, aby dopasowanie IDENTYCZNE (100%) zawsze mialo
    # pierwszenstwo i nie dublowac kopii wariantami z dopiskami (np. " K").
    if ($score -gt 99) { $score = 99 }
    return [int][Math]::Round($score)
}

# Odleglosc edycyjna Damerau-Levenshtein (OSA): liczba operacji (wstaw, usun,
# zamien, przestaw dwie sasiednie litery) potrzebnych do przerobienia a w b.
# Dzieki obsludze przestawien typowe literowki (np. "Cukeir" <-> "Cukier")
# licza sie jako 1 edycja.
function Get-Levenshtein {
    param([string]$a, [string]$b)
    $la = $a.Length; $lb = $b.Length
    if ($la -eq 0) { return $lb }
    if ($lb -eq 0) { return $la }
    $d = New-Object 'int[,]' ($la + 1), ($lb + 1)
    for ($i = 0; $i -le $la; $i++) { $d[$i, 0] = $i }
    for ($j = 0; $j -le $lb; $j++) { $d[0, $j] = $j }
    for ($i = 1; $i -le $la; $i++) {
        for ($j = 1; $j -le $lb; $j++) {
            $cost = if ($a[$i - 1] -eq $b[$j - 1]) { 0 } else { 1 }
            $del = $d[($i - 1), $j] + 1
            $ins = $d[$i, ($j - 1)] + 1
            $sub = $d[($i - 1), ($j - 1)] + $cost
            $val = [Math]::Min([Math]::Min($del, $ins), $sub)
            if ($i -gt 1 -and $j -gt 1 -and $a[$i - 1] -eq $b[$j - 2] -and $a[$i - 2] -eq $b[$j - 1]) {
                $val = [Math]::Min($val, $d[($i - 2), ($j - 2)] + 1)
            }
            $d[$i, $j] = $val
        }
    }
    return $d[$la, $lb]
}

# Zwraca podobienstwo 0-100 oraz typ dopasowania. Laczy dwie metody i bierze
# lepszy wynik:
#  - porownanie calych nazw po normalizacji (dobre dla nazw 1:1 i literowek),
#  - dopasowanie po slowach (dobre, gdy plik ma dopiski/inna kolejnosc slow).
function Get-Similarity {
    param(
        [string]$prodNorm, [string]$fileNorm,
        $prodTokens, $fileTokens
    )
    if ([string]::IsNullOrEmpty($prodNorm) -or [string]::IsNullOrEmpty($fileNorm)) {
        return [pscustomobject]@{ Score = 0; Typ = "brak" }
    }
    if ($prodNorm -eq $fileNorm) { return [pscustomobject]@{ Score = 100; Typ = "dokladne" } }

    # 1) calosc po normalizacji (wynik <100, by identyczne mialo pierwszenstwo)
    if ($fileNorm.Contains($prodNorm) -or $prodNorm.Contains($fileNorm)) {
        $short = [Math]::Min($prodNorm.Length, $fileNorm.Length)
        $long  = [Math]::Max($prodNorm.Length, $fileNorm.Length)
        $fullScore = [Math]::Min(99, [int](90 + 10.0 * $short / $long))
        $fullTyp = "zawiera"
    }
    else {
        $dist = Get-Levenshtein $prodNorm $fileNorm
        $maxLen = [Math]::Max($prodNorm.Length, $fileNorm.Length)
        $fullScore = [Math]::Min(99, [int][Math]::Round(100.0 * (1.0 - $dist / $maxLen)))
        $fullTyp = "rozmyte"
    }

    # 2) po slowach
    $tokScore = Get-TokenScore $prodTokens $fileTokens

    if ($tokScore -ge $fullScore) {
        return [pscustomobject]@{ Score = $tokScore; Typ = "slowa" }
    }
    return [pscustomobject]@{ Score = $fullScore; Typ = $fullTyp }
}

# ---------------------------------------------------------------------------
# 1. Ustalenie sciezek domyslnych
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($SourceFolder)) { $SourceFolder = $scriptDir }
if ([string]::IsNullOrWhiteSpace($TargetFolder)) { $TargetFolder = Join-Path $scriptDir "skopiowane" }

if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    $cand = Get-ChildItem -Path $scriptDir -File |
        Where-Object { $_.Extension -match '^\.(xlsx|csv)$' } |
        Sort-Object Extension |  # .csv przed .xlsx, ale i tak bierzemy pierwszy
        Select-Object -First 1
    if ($cand) { $ExcelPath = $cand.FullName }
}

if ([string]::IsNullOrWhiteSpace($ExcelPath) -or -not (Test-Path -LiteralPath $ExcelPath)) {
    Write-Bad "Nie znaleziono pliku Excel/CSV. Podaj go parametrem -ExcelPath lub umiesc plik .xlsx/.csv obok skryptu."
    exit 1
}
if (-not (Test-Path -LiteralPath $SourceFolder)) {
    Write-Bad "Folder zrodlowy nie istnieje: $SourceFolder"
    exit 1
}

$Extension = $Extension.TrimStart('.')

Write-Info "=== Kopiowanie etykiet wg listy produktow ==="
Write-Info "Plik z produktami : $ExcelPath"
Write-Info "Folder zrodlowy   : $SourceFolder"
Write-Info "Folder docelowy   : $TargetFolder"
Write-Info "Rozszerzenie      : .$Extension"
Write-Info "Kolumna           : $Column   (naglowek: $HasHeader)"
if ($ExactOnly) {
    Write-Info "Dopasowanie       : tylko dokladne (po normalizacji nazw)"
} else {
    Write-Info "Dopasowanie       : rozmyte (prog podobienstwa: $Threshold%)"
}
Write-Host ""

# ---------------------------------------------------------------------------
# 2. Funkcja czytajaca nazwy produktow z .xlsx (bez Excela) lub .csv
# ---------------------------------------------------------------------------
function Get-ProductNames {
    param([string]$Path, [string]$Col, [bool]$Header)

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($ext -eq ".csv") {
        # Z CSV bierzemy zadana kolumne wg indeksu (A=0, B=1...).
        $idx = ([int][char]$Col.ToUpper()[0]) - 65
        $lines = Get-Content -LiteralPath $Path
        $result = New-Object System.Collections.Generic.List[string]
        $start = if ($Header) { 1 } else { 0 }
        # Wykrycie separatora: srednik (typowy w PL Excelu) lub przecinek.
        $sep = if ($lines.Count -gt 0 -and ($lines[0] -split ';').Count -ge ($lines[0] -split ',').Count) { ';' } else { ',' }
        for ($i = $start; $i -lt $lines.Count; $i++) {
            $cells = $lines[$i] -split $sep
            if ($idx -lt $cells.Count) {
                $v = $cells[$idx].Trim().Trim('"')
                if (-not [string]::IsNullOrWhiteSpace($v)) { $result.Add($v) }
            }
        }
        return $result
    }

    if ($ext -eq ".xlsx") {
        Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            function Read-Entry($z, $name) {
                $e = $z.Entries | Where-Object { $_.FullName -eq $name } | Select-Object -First 1
                if (-not $e) { return $null }
                $sr = New-Object System.IO.StreamReader($e.Open())
                try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
            }

            # Shared strings (teksty komorek przechowywane sa wspolnie).
            $shared = @()
            $ssXml = Read-Entry $zip "xl/sharedStrings.xml"
            if ($ssXml) {
                [xml]$ss = $ssXml
                foreach ($si in $ss.sst.si) {
                    # <si> moze miec wiele <t> (formatowanie tekstu) - sklejamy.
                    $txt = ""
                    if ($si.t -is [string]) { $txt = $si.t }
                    elseif ($si.t) { $txt = ($si.t | ForEach-Object { $_.'#text' }) -join "" }
                    if ($si.r) { $txt = ($si.r.t | ForEach-Object { if ($_ -is [string]) { $_ } else { $_.'#text' } }) -join "" }
                    $shared += $txt
                }
            }

            # Ustalenie pierwszego arkusza wg workbook.xml -> rels.
            $sheetTarget = "xl/worksheets/sheet1.xml"
            $wbXml = Read-Entry $zip "xl/workbook.xml"
            $relXml = Read-Entry $zip "xl/_rels/workbook.xml.rels"
            if ($wbXml -and $relXml) {
                [xml]$wb = $wbXml
                [xml]$rel = $relXml
                $firstSheet = $wb.workbook.sheets.sheet | Select-Object -First 1
                $rid = $firstSheet.id   # r:id -> w PS dostepne jako .id
                if (-not $rid) { $rid = $firstSheet.'r:id' }
                $r = $rel.Relationships.Relationship | Where-Object { $_.Id -eq $rid } | Select-Object -First 1
                if ($r) {
                    $t = $r.Target -replace '^/xl/', '' -replace '^xl/', ''
                    $sheetTarget = "xl/$t"
                }
            }

            $sheetXml = Read-Entry $zip $sheetTarget
            if (-not $sheetXml) { $sheetXml = Read-Entry $zip "xl/worksheets/sheet1.xml" }
            [xml]$sheet = $sheetXml

            $result = New-Object System.Collections.Generic.List[string]
            $rowNum = 0
            foreach ($row in $sheet.worksheet.sheetData.row) {
                $rowNum++
                if ($Header -and $rowNum -eq 1) { continue }
                foreach ($c in $row.c) {
                    $ref = $c.r            # np. "A5"
                    $colLetters = ($ref -replace '[0-9]', '')
                    if ($colLetters -ne $Col.ToUpper()) { continue }

                    $val = $null
                    if ($c.t -eq "s") {
                        # wartosc to indeks w shared strings
                        $i = [int]$c.v
                        if ($i -ge 0 -and $i -lt $shared.Count) { $val = $shared[$i] }
                    }
                    elseif ($c.t -eq "inlineStr") {
                        $val = $c.is.t
                    }
                    else {
                        $val = $c.v
                    }
                    if (-not [string]::IsNullOrWhiteSpace([string]$val)) {
                        $result.Add(([string]$val).Trim())
                    }
                }
            }
            return $result
        }
        finally { $zip.Dispose() }
    }

    throw "Nieobslugiwany typ pliku: $ext (uzyj .xlsx lub .csv)"
}

# ---------------------------------------------------------------------------
# 3. Wczytanie nazw produktow
# ---------------------------------------------------------------------------
$products = @(Get-ProductNames -Path $ExcelPath -Col $Column -Header $HasHeader)
$products = $products | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
Write-Info "Wczytano nazw produktow: $($products.Count)"
if ($products.Count -eq 0) {
    Write-Bad "Brak nazw produktow w kolumnie $Column. Sprawdz parametry -Column / -HasHeader."
    exit 1
}

# ---------------------------------------------------------------------------
# 4. Wczytanie listy plikow zrodlowych (z policzona nazwa znormalizowana)
# ---------------------------------------------------------------------------
$gciParams = @{ Path = $SourceFolder; File = $true; Filter = "*.$Extension" }
if ($Recurse) { $gciParams["Recurse"] = $true }
$files = @(Get-ChildItem @gciParams | ForEach-Object {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    [pscustomobject]@{
        File   = $_
        Name   = $_.Name
        Norm   = (Get-NormName $base)
        Tokens = (Get-Tokens $base)
    }
})
Write-Info "Znaleziono plikow .$Extension : $($files.Count)"
Write-Host ""

# ---------------------------------------------------------------------------
# 5. Dopasowanie i kopiowanie
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $TargetFolder)) {
    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
}

$report = New-Object System.Collections.Generic.List[object]
$matchedFiles = New-Object System.Collections.Generic.HashSet[string]
$okCount = 0
$noCount = 0

# Prog dopasowania: w trybie -ExactOnly liczy sie tylko 100%.
$effThreshold = if ($ExactOnly) { 100 } else { $Threshold }

foreach ($p in $products) {
    $pNorm = Get-NormName $p
    $pTokens = Get-Tokens $p

    # Policz podobienstwo do kazdego pliku i znajdz najlepsze.
    $scored = foreach ($fi in $files) {
        $sim = Get-Similarity -prodNorm $pNorm -fileNorm $fi.Norm -prodTokens $pTokens -fileTokens $fi.Tokens
        [pscustomobject]@{ FileInfo = $fi; Score = $sim.Score; Typ = $sim.Typ }
    }
    $scored = @($scored | Sort-Object -Property Score -Descending)
    $best = $scored | Select-Object -First 1
    $bestScore = if ($best) { $best.Score } else { 0 }

    if ($best -and $bestScore -ge $effThreshold) {
        # Skopiuj wszystkie pliki z najlepszym wynikiem (obsluga duplikatow nazw).
        $winners = @($scored | Where-Object { $_.Score -eq $bestScore })
        foreach ($w in $winners) {
            $f = $w.FileInfo.File
            $dest = Join-Path $TargetFolder $f.Name
            $status = "SKOPIOWANO"
            $err = ""
            try {
                Copy-Item -LiteralPath $f.FullName -Destination $dest -Force
                [void]$matchedFiles.Add($f.FullName)
            }
            catch {
                $status = "BLAD_KOPIOWANIA"
                $err = $_.Exception.Message
            }
            if ($status -eq "SKOPIOWANO") {
                Write-Ok ("[OK ] {0}  ->  {1}   ({2}%, {3})" -f $p, $f.Name, $w.Score, $w.Typ)
                $okCount++
            } else {
                Write-Bad ("[ERR] {0}  ->  {1} ({2})" -f $p, $f.Name, $err)
            }
            $report.Add([pscustomobject]@{
                Produkt      = $p
                Plik         = $f.Name
                Status       = $status
                Podobienstwo = $w.Score
                Typ          = $w.Typ
                Sciezka      = $f.FullName
                Blad         = $err
            })
        }
    }
    else {
        # Brak dopasowania - pokaz najblizszego kandydata, by ulatwic ocene.
        $cand = if ($best) { $best.FileInfo.Name } else { "" }
        $candScore = $bestScore
        if ($cand) {
            Write-Bad ("[NIE] {0}  ->  brak dopasowania (najblizej: {1}, {2}%)" -f $p, $cand, $candScore)
        } else {
            Write-Bad ("[NIE] {0}  ->  brak plikow .{1} do porownania" -f $p, $Extension)
        }
        $noCount++
        $report.Add([pscustomobject]@{
            Produkt      = $p
            Plik         = ""
            Status       = "BRAK_DOPASOWANIA"
            Podobienstwo = $candScore
            Typ          = ("najblizej: " + $cand)
            Sciezka      = ""
            Blad         = ""
        })
    }
}

# Pliki, ktore istnieja, ale nie pasowaly do zadnego produktu
$unmatched = $files | Where-Object { -not $matchedFiles.Contains($_.File.FullName) }
foreach ($fi in $unmatched) {
    $report.Add([pscustomobject]@{
        Produkt      = ""
        Plik         = $fi.Name
        Status       = "NIEDOPASOWANY_PLIK"
        Podobienstwo = ""
        Typ          = ""
        Sciezka      = $fi.File.FullName
        Blad         = ""
    })
}

# ---------------------------------------------------------------------------
# 6. Zapis raportu
# ---------------------------------------------------------------------------
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportPath = Join-Path $TargetFolder ("raport_$stamp.csv")
$report | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

# ---------------------------------------------------------------------------
# 7. Podsumowanie
# ---------------------------------------------------------------------------
Write-Host ""
Write-Info "================ PODSUMOWANIE ================"
Write-Ok  ("Skopiowano plikow            : {0}" -f $okCount)
Write-Bad ("Produktow bez pasujacego pliku: {0}" -f $noCount)
Write-Warn2 ("Pliki .$Extension niedopasowane : {0}" -f $unmatched.Count)
Write-Info ("Raport CSV                    : {0}" -f $reportPath)
Write-Info ("Folder z kopiami              : {0}" -f $TargetFolder)
Write-Host ""
