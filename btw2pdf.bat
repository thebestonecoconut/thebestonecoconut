@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Konwerter etykiet BarTender (.btw) -^> TXT
REM ===================================================================
REM  btw2pdf.bat
REM  Wyciaga tekst z etykiet BarTender (.btw) do plikow TXT.
REM  NIE wymaga instalacji - uzywa wbudowanego w Windows PowerShell.
REM
REM  Uzycie:
REM    1) Przeciagnij plik .btw albo CALY FOLDER na ten plik .bat, lub
REM    2) Uruchom:  btw2pdf.bat "C:\sciezka\do\folderu"
REM    3) Sam dwuklik = przetwarza wszystkie .btw w folderze z tym .bat
REM ===================================================================

echo ================================================================
echo  Konwerter etykiet BarTender (.btw)  -^>  TXT
echo  Bez instalacji: uzywa wbudowanego PowerShell
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

# Obsluga przelacznikow przekazanych obok sciezek (np.  --raw)
$Raw = $false
if ($Paths) {
  $clean = New-Object System.Collections.Generic.List[string]
  foreach ($a in $Paths) {
    switch -regex ($a) {
      '^(?i)(--?raw|/raw)$'              { $Raw = $true; continue }
      '^(?i)(--?no-?pdf|/nopdf|--?txt)$' { continue }   # akceptowane, ignorowane
      default                            { $clean.Add($a) }
    }
  }
  $Paths = $clean.ToArray()
}

if (-not $Paths -or $Paths.Count -eq 0) {
  if ($env:BTW_DEFAULT) { $Paths = @($env:BTW_DEFAULT) } else { $Paths = @('.') }
}

$Latin1    = [System.Text.Encoding]::GetEncoding(28591)
$Cp1252    = [System.Text.Encoding]::GetEncoding(1252)
$Utf8      = New-Object System.Text.UTF8Encoding($false)
$Printable = '[\x20-\x7E\xA0-\u017F\u2010-\u2027\u20AC]'
$RunRe     = [regex]($Printable + '{2,}')
# UTF-16LE: znak (tab/CR/LF lub 0x20-0xFF) + starszy bajt 0x00-0x04 (wieloliniowy tekst)
$Utf16Re   = [regex]'(?:[\x09\x0A\x0D\x20-\xFF][\x00-\x04]){2,}'
# UTF-8: ASCII drukowalne (+ tab/CR/LF) oraz sekwencje 2- i 3-bajtowe
$Utf8Re    = [regex]'(?:[\x09\x0A\x0D\x20-\x7E]|[\xC2-\xDF][\x80-\xBF]|[\xE0-\xEF][\x80-\xBF]{2}){3,}'
$TagRe     = [regex]'<[^<>]{0,500}>'
$SplitRe   = [regex]'\s{2,}|[\r\n\t]+'
$NoiseRe   = [regex]'^(?:[0-9A-Fa-f]{16,}|[\W_]+|(?:[A-Za-z]\d*){1,2})$'
$Keywords  = @('BarTender','Seagull','btObject','btField','Microsoft','xmlns','schemas','FontName',
  'DataSource','FormatData','RichTextData','BackgroundData','TextData','DialogData',
  'ControlStringData','LineControlData','NumberFormatData','PrintJobFieldDs','BackgroundRFIDData',
  'ScreenDs','ScriptEvent','PredefinedStocksPage','StatusPage','DesignTemplatePage','Status Page',
  'Print Quantity','PrinterCopies','SerializedCount','BatchCount','FormatID','NICELbl',
  'Functions and Subs','OnProcessData','OnPostSerialize','Box Options','Dialog Control',
  'Line Control','Root.Folder','Word Processor','Default Paragraph Font',
  'Szablon','Warstwa','Pasek magnetyczny','Tekst procesora tekstu','Kontrolka przycisku',
  'Linia separatora','Formularz','Antena RFID','Obraz tła','Kolor tła','Numery seryjne',
  'Użyj ustawień drukarki','Wprowadź dan','Przykładowy tekst','Wspólne podprogramy',
  'odwołania innym','Picture.bmp','Bar Tender','Format File')

# --- Konwersja RTF -> czysty tekst ---
$RtfRe = New-Object System.Text.RegularExpressions.Regex(
  "\\([a-zA-Z]{1,32})(-?\d{1,10})?[ ]?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|[\r\n]+|(.)",
  [System.Text.RegularExpressions.RegexOptions]::Singleline)
$RtfDest = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($d in @('fonttbl','colortbl','stylesheet','info','pict','object','objdata','listtable',
  'listoverridetable','list','listlevel','listoverride','filetbl','revtbl','rsidtbl','generator',
  'themedata','datastore','latentstyles','defchp','defpap','pgptbl','panose','falt','fldinst',
  'xmlnstbl','wgrffmtfilter','listpicture','blipuid')) { [void]$RtfDest.Add($d) }
$RtfSpecial = @{
  'par'="`n"; 'sect'="`n"; 'page'="`n"; 'line'="`n"; 'tab'="`t"; 'cell'=' | '; 'row'="`n";
  'nestcell'=' | '; 'nestrow'="`n"; 'emdash'=[char]0x2014; 'endash'=[char]0x2013;
  'bullet'=[char]0x2022; 'lquote'=[char]0x2018; 'rquote'=[char]0x2019;
  'ldblquote'=[char]0x201C; 'rdblquote'=[char]0x201D; 'emspace'=' '; 'enspace'=' '; 'qmspace'=' '
}

function ConvertFrom-Rtf([string]$text) {
  $stack = New-Object System.Collections.Generic.Stack[object]
  $ignorable = $false; $ucskip = 1; $curskip = 0
  $sb = New-Object System.Text.StringBuilder
  foreach ($m in $RtfRe.Matches($text)) {
    $word=$m.Groups[1]; $arg=$m.Groups[2]; $hex=$m.Groups[3]
    $char=$m.Groups[4]; $brace=$m.Groups[5]; $tchar=$m.Groups[6]
    if ($brace.Success) {
      if ($brace.Value -eq '{') { $stack.Push(@($ucskip,$ignorable)) }
      elseif ($stack.Count -gt 0) { $st=$stack.Pop(); $ucskip=$st[0]; $ignorable=$st[1] }
    } elseif ($char.Success) {
      $c=$char.Value
      if ($c -eq '~') { if (-not $ignorable) { [void]$sb.Append([char]0x00A0) } }
      elseif ($c -eq '{' -or $c -eq '}' -or $c -eq '\') { if (-not $ignorable) { [void]$sb.Append($c) } }
      elseif ($c -eq '*') { $ignorable=$true }
    } elseif ($word.Success) {
      $w=$word.Value; $curskip=0
      if ($RtfDest.Contains($w)) { $ignorable=$true }
      elseif ($ignorable) { }
      elseif ($RtfSpecial.ContainsKey($w)) { [void]$sb.Append([string]$RtfSpecial[$w]) }
      elseif ($w -eq 'uc') { $ucskip = if ($arg.Success) { [int]$arg.Value } else { 1 } }
      elseif ($w -eq 'u') {
        $cval=[int]$arg.Value; if ($cval -lt 0) { $cval += 0x10000 }
        if (-not $ignorable) { [void]$sb.Append([char]$cval) }
        $curskip=$ucskip
      }
    } elseif ($hex.Success) {
      if ($curskip -gt 0) { $curskip-- }
      elseif (-not $ignorable) { [void]$sb.Append($Cp1252.GetString([byte[]]@([Convert]::ToInt32($hex.Value,16)))) }
    } elseif ($tchar.Success) {
      if ($curskip -gt 0) { $curskip-- }
      elseif (-not $ignorable) { [void]$sb.Append($tchar.Value) }
    }
  }
  return $sb.ToString()
}

function Test-Content([string]$s) {
  if ($s.Length -lt 2) { return $false }
  if ($NoiseRe.IsMatch($s)) { return $false }
  $low = $s.ToLower()
  foreach ($k in $Keywords) { if ($low.Contains($k.ToLower())) { return $false } }
  # musi zawierac litere lub cyfre (dowolny alfabet/jezyk)
  if (-not [regex]::IsMatch($s, '[\p{L}\p{Nd}]')) { return $false }
  # Krotkie tokeny bez spacji z "kodowymi" znakami to zwykle smieci z danych binarnych
  if ($s.Length -lt 9 -and $s -notmatch '\s' -and $s -match '[()&*<>|{}\[\]^~`\\=;%#$@+]') { return $false }
  return $true
}

# Wykrywa RTF i konwertuje na czysty tekst; inaczej dekoduje encje XML i usuwa
# znaczniki (zostawiajac tresc). W trybie raw zwraca wszystko bez zmian.
function Expand-Region([string]$text, [bool]$raw, $runs) {
  if ($raw) {
    foreach ($r in $RunRe.Matches($text)) { $runs.Add($r.Value.Trim()) }
    return
  }
  if ($text.Contains('\rtf')) {
    foreach ($line in (ConvertFrom-Rtf $text) -split "`r?`n") {
      $l = ($line -replace '[ \t]+',' ').Trim()
      $l = $l.Trim('|').Trim()
      if ($l.Length -gt 0) { $runs.Add($l) }
    }
    return
  }
  $t = $text -replace '&amp;','&' -replace '&lt;','<' -replace '&gt;','>' `
             -replace '&quot;','"' -replace '&apos;',"'" -replace '&#39;',"'"
  $t = $TagRe.Replace($t, '  ')          # usun znaczniki, zostaw tresc
  foreach ($piece in $SplitRe.Split($t)) {
    foreach ($r in $RunRe.Matches($piece)) { $runs.Add($r.Value.Trim()) }
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

$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

# --- glowna petla ---------------------------------------------------
$total = 0

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
    $total++
  }
}

Write-Host ""
Write-Host ("Przetworzono plikow: {0}" -f $total) -ForegroundColor Cyan
if (-not $Raw) {
  Write-Host "Wskazowka: jesli w .txt czegos brakuje, uruchom z opcja --raw (wszystkie ciagi)." -ForegroundColor DarkGray
}
