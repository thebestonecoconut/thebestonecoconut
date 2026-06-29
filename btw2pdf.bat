@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Konwerter etykiet BarTender (.btw) -^> TXT + PDF
REM ===================================================================
REM  btw2pdf.bat
REM  Wyciaga tekst z etykiet BarTender (.btw) do plikow TXT i PDF.
REM  NIE wymaga instalacji - uzywa rzeczy wbudowanych w Windows 10/11:
REM    * Windows PowerShell (TXT + odczyt .btw)
REM    * Microsoft Edge w trybie headless (generowanie PDF)
REM
REM  Uzycie:
REM    1) Przeciagnij plik .btw albo CALY FOLDER na ten plik .bat, lub
REM    2) Uruchom:  btw2pdf.bat "C:\sciezka\do\folderu"
REM    3) Sam dwuklik = przetwarza wszystkie .btw w folderze z tym .bat
REM ===================================================================

echo ================================================================
echo  Konwerter etykiet BarTender (.btw)  -^>  TXT + PDF
echo  Bez instalacji: PowerShell + Microsoft Edge (Windows 10/11)
echo ================================================================
echo.

set "BTW_DEFAULT=%~dp0"
set "TMPPS=%TEMP%\btw2pdf_%RANDOM%%RANDOM%.ps1"

REM Wyodrebnij sekcje PowerShell z tego pliku (wszystko po znaczniku ponizej)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$raw=[System.Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes('%~f0')); $m=[char]35+'PSSTART'; $i=$raw.IndexOf($m); $code=$raw.Substring($i+$m.Length); [IO.File]::WriteAllText($env:TMPPS,$code,(New-Object System.Text.UTF8Encoding($false)))"

REM Uruchom wlasciwa logike (przekazuje przeciagniete pliki/foldery)
powershell -NoProfile -ExecutionPolicy Bypass -File "%TMPPS%" %*
set "RC=%ERRORLEVEL%"
del "%TMPPS%" >nul 2>&1

echo.
echo Gotowe. Nacisnij dowolny klawisz, aby zamknac to okno.
pause >nul
exit /b %RC%

#PSSTART
param([Parameter(ValueFromRemainingArguments=$true)][string[]]$Paths)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Obsluga przelacznikow przekazanych obok sciezek (np.  --raw,  --no-pdf)
$Raw   = $false
$NoPdf = $false
if ($Paths) {
  $clean = New-Object System.Collections.Generic.List[string]
  foreach ($a in $Paths) {
    switch -regex ($a) {
      '^(?i)(--?raw|/raw)$'              { $Raw   = $true; continue }
      '^(?i)(--?no-?pdf|/nopdf|--?txt)$' { $NoPdf = $true; continue }
      default                            { $clean.Add($a) }
    }
  }
  $Paths = $clean.ToArray()
}

if (-not $Paths -or $Paths.Count -eq 0) {
  if ($env:BTW_DEFAULT) { $Paths = @($env:BTW_DEFAULT) } else { $Paths = @('.') }
}

$Latin1    = [System.Text.Encoding]::GetEncoding(28591)
$Utf8      = New-Object System.Text.UTF8Encoding($false)
$Printable = '[\x20-\x7E\xA0-\u024F\u2010-\u2027\u20AC]'
$RunRe     = [regex]($Printable + '{2,}')
# UTF-16LE: znak (mlodszy bajt 0x20-0xFF) + starszy bajt 0x00-0x04
$Utf16Re   = [regex]'(?:[\x20-\xFF][\x00-\x04]){2,}'
# UTF-8: ASCII drukowalne oraz sekwencje 2- i 3-bajtowe (m.in. polskie znaki, EUR, mysliniki)
$Utf8Re    = [regex]'(?:[\x20-\x7E]|[\xC2-\xDF][\x80-\xBF]|[\xE0-\xEF][\x80-\xBF]{2}){3,}'
$TagRe     = [regex]'<[^<>]{0,500}>'
$SplitRe   = [regex]'\s{2,}|[\r\n\t]+'
$NoiseRe   = [regex]'^(?:[0-9A-Fa-f]{16,}|[\W_]+|(?:[A-Za-z]\d*){1,2})$'
$Keywords  = @('BarTender','Seagull','btObject','btField','Microsoft','xmlns','schemas','FontName')

function Test-Content([string]$s) {
  if ($s.Length -lt 2) { return $false }
  if ($NoiseRe.IsMatch($s)) { return $false }
  $low = $s.ToLower()
  foreach ($k in $Keywords) { if ($low.Contains($k.ToLower())) { return $false } }
  if (-not [regex]::IsMatch($s, '[0-9A-Za-z\u00A0-\u024F]')) { return $false }
  # Krotkie tokeny bez spacji z "kodowymi" znakami to zwykle smieci z danych binarnych
  if ($s.Length -lt 8 -and $s -notmatch '\s' -and $s -match '[()&*<>|{}\[\]^~`\\=;]') { return $false }
  return $true
}

# Dekoduje encje XML i (poza trybem raw) usuwa znaczniki, zachowujac tresc.
function Expand-Region([string]$text, [bool]$raw, $runs) {
  if (-not $raw) {
    $t = $text -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' `
               -replace '&quot;','"' -replace '&apos;',"'" -replace '&#39;',"'"
    $t = $TagRe.Replace($t, '  ')          # usun znaczniki, zostaw tresc
    foreach ($piece in $SplitRe.Split($t)) {
      foreach ($r in $RunRe.Matches($piece)) { $runs.Add($r.Value.Trim()) }
    }
  } else {
    foreach ($r in $RunRe.Matches($text)) { $runs.Add($r.Value.Trim()) }
  }
}

# Wyciaga tekst UTF-16LE i UTF-8 z danego ciagu (kazdy znak = 1 bajt 0-255).
function Add-Strings([string]$s, [bool]$raw, $runs) {
  foreach ($m in $Utf16Re.Matches($s)) {
    $mb = $Latin1.GetBytes($m.Value)
    Expand-Region ([System.Text.Encoding]::Unicode.GetString($mb)) $raw $runs
  }
  foreach ($m in $Utf8Re.Matches($s)) {
    $mb = $Latin1.GetBytes($m.Value)
    Expand-Region ($Utf8.GetString($mb)) $raw $runs
  }
}

# Probuje rozpakowac surowy strumien DEFLATE zaczynajacy sie pod $start.
function Try-Inflate([byte[]]$bytes, [int]$start) {
  try {
    $in  = New-Object System.IO.MemoryStream
    $in.Write($bytes, $start, $bytes.Length - $start)
    $in.Position = 0
    $def = New-Object System.IO.Compression.DeflateStream($in, [System.IO.Compression.CompressionMode]::Decompress)
    $out = New-Object System.IO.MemoryStream
    $buf = New-Object byte[] 16384
    while (($n = $def.Read($buf, 0, $buf.Length)) -gt 0) { $out.Write($buf, 0, $n) }
    $def.Dispose()
    return $out.ToArray()
  } catch { return $null }
}

# Adler-32 (suma kontrolna konczaca strumien zlib) jako 4 bajty big-endian.
function Get-Adler32Bytes([byte[]]$data) {
  $a = 1; $b = 0; $M = 65521
  foreach ($x in $data) { $a = ($a + $x) % $M; $b = ($b + $a) % $M }
  return [byte[]]@(
    [byte](($b -shr 8) -band 0xFF), [byte]($b -band 0xFF),
    [byte](($a -shr 8) -band 0xFF), [byte]($a -band 0xFF)
  )
}

function IndexOf-Bytes([byte[]]$hay, [byte[]]$needle, [int]$from) {
  for ($k = $from; $k -le $hay.Length - $needle.Length; $k++) {
    $ok = $true
    for ($j = 0; $j -lt $needle.Length; $j++) {
      if ($hay[$k + $j] -ne $needle[$j]) { $ok = $false; break }
    }
    if ($ok) { return $k }
  }
  return -1
}

function Get-BtwLines([string]$File, [bool]$raw) {
  $bytes = [System.IO.File]::ReadAllBytes($File)
  $blob  = $Latin1.GetString($bytes)   # 1 znak = 1 bajt (0-255)

  # Usun osadzone obrazy PNG (podglad etykiety) - to one generuja wiekszosc smieci.
  $blob = [regex]::Replace($blob, "\x89PNG\r\n\x1A\n.*?IEND.{4}", " ",
                           [System.Text.RegularExpressions.RegexOptions]::Singleline)

  $runs = New-Object System.Collections.Generic.List[string]
  $work = $Latin1.GetBytes($blob)

  # 1) Rozpakuj skompresowane (zlib/deflate) fragmenty - tam siedzi tresc etykiety.
  #    Po udanym rozpakowaniu "wygaszamy" skompresowany obszar (do sumy Adler-32),
  #    zeby jego surowe bajty nie trafily do wyniku jako smieci.
  $attempts = 0
  for ($i = 0; $i -lt $work.Length - 2; $i++) {
    if ($work[$i] -ne 0x78) { continue }
    $h = ([int]$work[$i] -shl 8) -bor [int]$work[$i+1]
    if (($h % 31) -ne 0) { continue }          # naglowek zlib spelnia regule mod 31
    if ($attempts -ge 60000) { break }
    $attempts++
    $inf = Try-Inflate $work ($i + 2)
    if ($inf -and $inf.Length -gt 16) {
      Add-Strings ($Latin1.GetString($inf)) $raw $runs
      $adler = Get-Adler32Bytes $inf
      $p = IndexOf-Bytes $work $adler ($i + 2)
      $end = if ($p -ge 0) { $p + 4 } else { $i + 2 }
      for ($j = $i; $j -lt $end -and $j -lt $work.Length; $j++) { $work[$j] = 0 }
      if ($end -gt $i) { $i = $end - 1 }
    }
  }

  # 2) Czysty tekst (UTF-16 / UTF-8) z pozostalej czesci (m.in. metadane naglowka).
  Add-Strings ($Latin1.GetString($work)) $raw $runs

  $seen  = New-Object System.Collections.Generic.HashSet[string]
  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($run in $runs) {
    $c = $run.Trim()
    if ($c.Length -lt 2) { continue }
    if (-not $raw -and -not (Test-Content $c)) { continue }
    if ($seen.Add($c)) { $lines.Add($c) }
  }
  return ,$lines.ToArray()
}

function Convert-HtmlEscape([string]$s) {
  return $s.Replace('&','&amp;').Replace('<','&lt;').Replace('>','&gt;')
}

function Build-Html([string]$name, [string[]]$lines) {
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.AppendLine('<!DOCTYPE html><html lang="pl"><head><meta charset="utf-8">')
  [void]$sb.AppendLine('<style>')
  [void]$sb.AppendLine('@page { size: A4; margin: 18mm; }')
  [void]$sb.AppendLine('body { font-family: "Segoe UI", Arial, sans-serif; font-size: 12pt; color:#111; }')
  [void]$sb.AppendLine('h1 { font-size: 15pt; border-bottom: 2px solid #333; padding-bottom:6px; }')
  [void]$sb.AppendLine('ul { line-height: 1.6; } li { margin: 2px 0; }')
  [void]$sb.AppendLine('.empty { color:#999; font-style:italic; }')
  [void]$sb.AppendLine('</style></head><body>')
  [void]$sb.AppendLine('<h1>Etykieta: ' + (Convert-HtmlEscape $name) + '</h1>')
  if ($lines.Count -eq 0) {
    [void]$sb.AppendLine('<p class="empty">(nie znaleziono tekstu w tym pliku)</p>')
  } else {
    [void]$sb.AppendLine('<ul>')
    foreach ($l in $lines) { [void]$sb.AppendLine('<li>' + (Convert-HtmlEscape $l) + '</li>') }
    [void]$sb.AppendLine('</ul>')
  }
  [void]$sb.AppendLine('</body></html>')
  return $sb.ToString()
}

function Find-Edge {
  $bases = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:LOCALAPPDATA) | Where-Object { $_ }
  foreach ($b in $bases) {
    $c = Join-Path $b 'Microsoft\Edge\Application\msedge.exe'
    if (Test-Path -LiteralPath $c) { return $c }
  }
  return $null
}

$EdgePath = Find-Edge
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

function Convert-HtmlToPdf([string]$html, [string]$pdfPath) {
  if (-not $EdgePath) { return $false }
  $tmpHtml = Join-Path $env:TEMP ('btw_' + [guid]::NewGuid().ToString('N') + '.html')
  $profile = Join-Path $env:TEMP ('btwedge_' + [guid]::NewGuid().ToString('N'))
  [IO.File]::WriteAllText($tmpHtml, $html, $Utf8NoBom)
  $url = 'file:///' + ($tmpHtml -replace '\\','/')
  try {
    & $EdgePath "--headless" "--disable-gpu" "--user-data-dir=$profile" "--print-to-pdf-no-header" "--print-to-pdf=$pdfPath" "$url" 2>$null | Out-Null
  } catch { }
  Remove-Item -LiteralPath $tmpHtml -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath $profile -Recurse -Force -ErrorAction SilentlyContinue
  return (Test-Path -LiteralPath $pdfPath)
}

# --- glowna petla ---------------------------------------------------
$total = 0
$noEdgeWarned = $false

foreach ($p in $Paths) {
  $files = @()
  if (Test-Path -LiteralPath $p -PathType Container) {
    $files = Get-ChildItem -LiteralPath $p -Recurse -Filter *.btw -File
  } elseif (Test-Path -LiteralPath $p) {
    $item = Get-Item -LiteralPath $p
    if ($item.Extension -ieq '.btw') { $files = @($item) }
  } else {
    Write-Host ("  ! Pominieto (brak sciezki): {0}" -f $p) -ForegroundColor Yellow
    continue
  }

  if ($files.Count -eq 0) {
    Write-Host ("  Brak plikow .btw w: {0}" -f $p) -ForegroundColor Yellow
    continue
  }

  foreach ($f in $files) {
    try {
      $lines = Get-BtwLines $f.FullName $Raw
    } catch {
      Write-Host ("  ! Blad przy {0}: {1}" -f $f.Name, $_.Exception.Message) -ForegroundColor Red
      continue
    }

    $txt = [IO.Path]::ChangeExtension($f.FullName, '.txt')
    $tryb = if ($Raw) { 'tryb RAW - bez filtrowania' } else { 'tekst oczyszczony' }
    $header = "# Etykieta: $($f.Name)`r`n# Wyodrebniony tekst ($($lines.Count) pozycji, $tryb)`r`n`r`n"
    [IO.File]::WriteAllText($txt, $header + ($lines -join "`r`n") + "`r`n", $Utf8NoBom)
    Write-Host ("  [TXT] {0}  ({1} linii)" -f (Split-Path $txt -Leaf), $lines.Count) -ForegroundColor Green

    if (-not $NoPdf) {
      $pdf  = [IO.Path]::ChangeExtension($f.FullName, '.pdf')
      $html = Build-Html $f.Name $lines
      if (Convert-HtmlToPdf $html $pdf) {
        Write-Host ("  [PDF] {0}" -f (Split-Path $pdf -Leaf)) -ForegroundColor Green
      } else {
        $htmlOut = [IO.Path]::ChangeExtension($f.FullName, '.html')
        [IO.File]::WriteAllText($htmlOut, $html, $Utf8NoBom)
        if (-not $noEdgeWarned) {
          Write-Host "  ! Nie udalo sie uzyc Edge do PDF (PDF to dodatek). Zapisuje HTML." -ForegroundColor Yellow
          Write-Host "    Otworz plik .html i wcisnij Ctrl+P -> 'Microsoft Print to PDF'." -ForegroundColor Yellow
          $noEdgeWarned = $true
        }
        Write-Host ("  [HTML] {0}" -f (Split-Path $htmlOut -Leaf)) -ForegroundColor Yellow
      }
    }
    $total++
  }
}

Write-Host ""
Write-Host ("Przetworzono plikow: {0}" -f $total) -ForegroundColor Cyan
if (-not $Raw) {
  Write-Host "Wskazowka: jesli w .txt czegos brakuje, uruchom z opcja --raw (wszystkie ciagi)." -ForegroundColor DarkGray
}
