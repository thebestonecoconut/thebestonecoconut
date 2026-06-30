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
    # tym ostrzej (mniej trafien, mniej pomylek). Domyslnie 80.
    [int]$Threshold = 80,

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
#  - polskie znaki -> bez ogonkow (ą->a, ł->l ...),
#  - usuniecie spacji, podkreslen, mysmikow, kropek i innych nie-alfanum.
# Dzieki temu "Produkt_Alfa-01.btw" i "produkt alfa 01" sa traktowane podobnie.
# ---------------------------------------------------------------------------
function Get-NormName {
    param([string]$s)
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    $s = $s.ToLowerInvariant()
    $map = @{
        'ą'='a';'ć'='c';'ę'='e';'ł'='l';'ń'='n';'ó'='o';'ś'='s';'ź'='z';'ż'='z'
    }
    foreach ($k in $map.Keys) { $s = $s.Replace($k, $map[$k]) }
    # usun wszystko poza literami i cyframi
    $s = ($s -replace '[^a-z0-9]', '')
    return $s
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

# Zwraca podobienstwo 0-100 oraz typ dopasowania dla dwoch znormalizowanych nazw.
function Get-Similarity {
    param([string]$prod, [string]$file)
    if ($prod -eq "" -or $file -eq "") { return [pscustomobject]@{ Score = 0; Typ = "brak" } }
    if ($prod -eq $file) { return [pscustomobject]@{ Score = 100; Typ = "dokladne" } }
    if ($file.Contains($prod) -or $prod.Contains($file)) {
        # nazwa jednego zawiera sie w drugiej - mocne dopasowanie,
        # punktacja zalezna od tego ile "nadmiaru" jest w dluzszej nazwie
        $short = [Math]::Min($prod.Length, $file.Length)
        $long  = [Math]::Max($prod.Length, $file.Length)
        $score = [int](90 + 10.0 * $short / $long)   # 90-100
        return [pscustomobject]@{ Score = $score; Typ = "zawiera" }
    }
    $dist = Get-Levenshtein $prod $file
    $maxLen = [Math]::Max($prod.Length, $file.Length)
    $score = [int][Math]::Round(100.0 * (1.0 - $dist / $maxLen))
    return [pscustomobject]@{ Score = $score; Typ = "rozmyte" }
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
        File = $_
        Name = $_.Name
        Norm = (Get-NormName $base)
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

    # Policz podobienstwo do kazdego pliku i znajdz najlepsze.
    $scored = foreach ($fi in $files) {
        $sim = Get-Similarity -prod $pNorm -file $fi.Norm
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
