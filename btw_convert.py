#!/usr/bin/env python3
"""
btw_convert.py — wyciąga tekst z plików etykiet BarTender (.btw) i zapisuje
go do pliku tekstowego (.txt).

Pliki .btw to dokumenty OLE2 (Compound File Binary Format) tworzone przez
program BarTender (Seagull Scientific). Wewnątrz przechowują m.in. teksty
etykiet (zwykle jako RTF / XML / tekst w kodowaniu UTF-16, często skompresowane
zlib). Ten skrypt otwiera taki plik, wydobywa z niego czytelne ciągi znaków
i zapisuje je w pliku tekstowym.

UWAGA: pełne, "pixel-perfect" odwzorowanie wyglądu etykiety (kody kreskowe,
grafiki, układ) potrafi wygenerować TYLKO sam BarTender. Ten skrypt służy do
wyciągnięcia ZAWARTOŚCI TEKSTOWEJ etykiet.

Użycie:
    python btw_convert.py PLIK.btw                  # -> PLIK.txt
    python btw_convert.py folder/                   # konwertuje wszystkie .btw
    python btw_convert.py folder/ -o wyniki/        # zapis do innego folderu
    python btw_convert.py PLIK.btw --raw            # wszystkie ciągi (weryfikacja)
"""

from __future__ import annotations

import argparse
import re
import sys
import zlib
from pathlib import Path

try:
    import olefile
except ImportError:  # pragma: no cover
    olefile = None


# --- Wyciąganie tekstu -------------------------------------------------------

# Sekwencje co najmniej 2 "drukowalnych" znaków. Obsługujemy polskie znaki.
_PRINTABLE = (
    r"[\x20-\x7E"
    r"\u00A0-\u017F"   # Latin-1 + Latin Extended-A (ą, ć, ę, ł, ó, ä, ö, ü, ß...)
    r"\u2010-\u2027"   # myślniki, cudzysłowy itp.
    r"\u20AC]"         # €
)
_RUN_RE = re.compile(_PRINTABLE + r"{2,}")

# Śmieci typowe dla wnętrza plików OLE / definicji obiektów, które chcemy odrzucić.
_NOISE_RE = re.compile(
    r"^(?:"
    r"[0-9A-Fa-f]{16,}"                       # długie ciągi hex
    r"|[\W_]+"                                # same znaki niealfanumeryczne
    r"|(?:[A-Za-z]\d*){1,2}"                  # bardzo krótkie kody typu A1, B2
    r")$"
)

# Słowa-klucze wewnętrznych nazw BarTendera, które zwykle nie są treścią etykiety.
_INTERNAL_KEYWORDS = (
    "BarTender", "Seagull", "btObject", "btField", "Microsoft",
    "xmlns", "http://", "https://", "schemas", "FontName",
    "DataSource", "FormatData", "RichTextData", "BackgroundData", "TextData",
    "DialogData", "ControlStringData", "LineControlData", "NumberFormatData",
    "PrintJobFieldDs", "BackgroundRFIDData", "ScreenDs", "ScriptEvent",
    "PredefinedStocksPage", "StatusPage", "DesignTemplatePage", "Status Page",
    "Print Quantity", "PrinterCopies", "SerializedCount", "BatchCount",
    "FormatID", "NICELbl",     "Functions and Subs", "OnProcessData",
    "OnPostSerialize", "Box Options", "Dialog Control", "Line Control",
    "Root.Folder", "Word Processor", "Default Paragraph Font",
    # Domyślne (auto-generowane) polskie nazwy obiektów BarTendera
    "Szablon", "Warstwa", "Pasek magnetyczny", "Tekst procesora tekstu",
    "Kontrolka przycisku", "Linia separatora", "Formularz", "Antena RFID",
    "Obraz tła", "Kolor tła", "Numery seryjne", "Użyj ustawień drukarki",
    "Wprowadź dan", "Przykładowy tekst", "Wspólne podprogramy",
    "odwołania innym", "Picture.bmp", "Bar Tender", "Format File",
)


# UTF-16LE: znak (młodszy bajt 0x09/0x0A/0x0D/0x20-0xFF) + starszy bajt 0x00-0x04.
# Dopuszczamy tab/CR/LF, żeby wieloliniowy tekst (np. RTF) nie był cięty.
_UTF16_RUN_RE = re.compile(rb"(?:[\x09\x0a\x0d\x20-\xff][\x00-\x04]){2,}")
# UTF-8: ASCII drukowalne (+ tab/CR/LF) oraz sekwencje 2- i 3-bajtowe.
_UTF8_RUN_RE = re.compile(
    rb"(?:[\x09\x0a\x0d\x20-\x7e]|[\xc2-\xdf][\x80-\xbf]|[\xe0-\xef][\x80-\xbf]{2}){3,}"
)
# Osadzony obraz PNG (podgląd etykiety) — usuwamy, bo to źródło większości śmieci.
_PNG_RE = re.compile(rb"\x89PNG\r\n\x1a\n.*?IEND.{4}", re.DOTALL)
# Usuwanie znaczników XML (z zachowaniem treści między nimi).
_TAG_RE = re.compile(r"<[^<>]{0,500}>")
_SPLIT_RE = re.compile(r"\s{2,}|[\r\n\t]+")
_ENTITIES = (
    ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
    ("&quot;", '"'), ("&apos;", "'"), ("&#39;", "'"),
)


# --- Konwersja RTF -> czysty tekst -------------------------------------------

_RTF_RE = re.compile(
    r"\\([a-zA-Z]{1,32})(-?\d{1,10})?[ ]?|\\'([0-9a-fA-F]{2})|\\([^a-zA-Z])|([{}])|[\r\n]+|(.)",
    re.DOTALL,
)
_RTF_DESTINATIONS = frozenset((
    "fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "objdata",
    "listtable", "listoverridetable", "list", "listlevel", "listoverride",
    "filetbl", "revtbl", "rsidtbl", "generator", "themedata", "datastore",
    "latentstyles", "defchp", "defpap", "pgptbl", "panose", "falt",
    "fldinst", "xmlnstbl", "wgrffmtfilter", "listpicture", "blipuid",
))
_RTF_SPECIAL = {
    "par": "\n", "sect": "\n", "page": "\n", "line": "\n", "tab": "\t",
    "cell": " | ", "row": "\n", "nestcell": " | ", "nestrow": "\n",
    "emdash": "\u2014", "endash": "\u2013", "bullet": "\u2022",
    "lquote": "\u2018", "rquote": "\u2019", "ldblquote": "\u201C", "rdblquote": "\u201D",
    "emspace": " ", "enspace": " ", "qmspace": " ",
}


def _rtf_to_text(text: str) -> str:
    """Zamienia RTF na czysty tekst (dekoduje \\uN i \\'xx, usuwa formatowanie)."""
    stack: list[tuple[int, bool]] = []
    ignorable = False
    ucskip = 1
    curskip = 0
    out: list[str] = []
    for m in _RTF_RE.finditer(text):
        word, arg, hexv, char, brace, tchar = m.groups()
        if brace:
            if brace == "{":
                stack.append((ucskip, ignorable))
            elif brace == "}" and stack:
                ucskip, ignorable = stack.pop()
        elif char is not None:
            if char == "~":
                if not ignorable:
                    out.append("\u00A0")
            elif char in "{}\\":
                if not ignorable:
                    out.append(char)
            elif char == "*":
                ignorable = True
        elif word:
            curskip = 0
            if word in _RTF_DESTINATIONS:
                ignorable = True
            elif ignorable:
                pass
            elif word in _RTF_SPECIAL:
                out.append(_RTF_SPECIAL[word])
            elif word == "uc":
                ucskip = int(arg) if arg else 1
            elif word == "u":
                c = int(arg)
                if c < 0:
                    c += 0x10000
                if not ignorable:
                    out.append(chr(c) if c <= 0x10FFFF else "?")
                curskip = ucskip
        elif hexv is not None:
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                out.append(bytes([int(hexv, 16)]).decode("cp1252", "replace"))
        elif tchar:
            if curskip > 0:
                curskip -= 1
            elif not ignorable:
                out.append(tchar)
    return "".join(out)


_WS_RE = re.compile(r"[ \t]+")


def _expand_region(text: str, raw: bool, out: list[str]) -> None:
    """Dla zdekodowanego fragmentu dorzuca do `out` znalezione ciągi znaków.

    Wykrywa RTF i zamienia go na czysty tekst; w trybie zwykłym dekoduje encje
    XML i usuwa znaczniki; w trybie raw zwraca wszystko bez zmian.
    """
    if raw:
        for match in _RUN_RE.findall(text):
            cleaned = match.strip()
            if cleaned:
                out.append(cleaned)
        return

    if "\\rtf" in text:
        for line in _rtf_to_text(text).splitlines():
            line = _WS_RE.sub(" ", line).strip(" |")
            if line:
                out.append(line)
        return

    t = text
    for ent, ch in _ENTITIES:
        t = t.replace(ent, ch)
    t = _TAG_RE.sub("  ", t)
    for piece in _SPLIT_RE.split(t):
        for match in _RUN_RE.findall(piece):
            cleaned = match.strip()
            if cleaned:
                out.append(cleaned)


def _decode_runs(data: bytes, raw: bool, out: list[str]) -> None:
    """Dorzuca do `out` tekst UTF-16LE i UTF-8 znaleziony w surowych bajtach.

    Szuka osobno tekstu UTF-16LE (najczęstszy w BarTenderze) oraz UTF-8/ASCII,
    dzięki czemu unika śmieci powstających przy "ślepym" dekodowaniu strumienia.
    """
    for chunk in _UTF16_RUN_RE.findall(data):
        _expand_region(chunk.decode("utf-16-le", errors="ignore"), raw, out)

    for chunk in _UTF8_RUN_RE.findall(data):
        _expand_region(chunk.decode("utf-8", errors="ignore"), raw, out)


def _scan_bytes(data: bytes, raw: bool) -> list[str]:
    """Wyciąga tekst z bajtów: usuwa osadzone PNG i rozpakowuje strumienie zlib.

    BarTender trzyma podgląd etykiety jako PNG (śmieci) oraz właściwą treść
    skompresowaną (deflate/zlib). Najpierw usuwamy PNG, potem znajdujemy i
    rozpakowujemy strumienie zlib (tam jest tekst etykiety), a ich skompresowane
    bajty „wygaszamy", żeby nie trafiły do wyniku jako śmieci.
    """
    out: list[str] = []
    work = bytearray(_PNG_RE.sub(b"  ", data))
    n = len(work)

    i = 0
    attempts = 0
    while i < n - 2:
        if work[i] != 0x78:
            i += 1
            continue
        if ((work[i] << 8) | work[i + 1]) % 31 != 0:  # reguła nagłówka zlib (mod 31)
            i += 1
            continue
        if attempts >= 60000:
            break
        attempts += 1
        try:
            dec = zlib.decompressobj()
            inflated = dec.decompress(bytes(work[i:]))
            inflated += dec.flush()
            consumed = n - i - len(dec.unused_data)
        except Exception:
            i += 1
            continue
        if len(inflated) > 16 and consumed >= 2:
            _decode_runs(inflated, raw, out)
            for j in range(i, min(i + consumed, n)):
                work[j] = 0
            i += consumed
        else:
            i += 1

    _decode_runs(bytes(work), raw, out)
    return out


def _looks_like_content(s: str) -> bool:
    """Heurystyka: czy dany ciąg wygląda na realną treść etykiety."""
    if len(s) < 2:
        return False
    if _NOISE_RE.match(s):
        return False
    if any(k.lower() in s.lower() for k in _INTERNAL_KEYWORDS):
        return False
    # musi zawierać przynajmniej jedną literę lub cyfrę (dowolny alfabet/język)
    if not any(ch.isalnum() for ch in s):
        return False
    # krótkie tokeny bez spacji z "kodowymi" znakami to zwykle śmieci binarne
    if len(s) < 9 and not re.search(r"\s", s) and re.search(r"[()&*<>|{}\[\]^~`\\=;%#$@+]", s):
        return False
    return True


def extract_text_from_btw(path: Path, raw: bool = False) -> list[str]:
    """Wyciąga uporządkowaną listę linii tekstu z pliku .btw.

    raw=True: zwraca wszystkie znalezione ciągi bez filtrowania (do weryfikacji).
    """
    raw_runs: list[str] = []

    if olefile is not None and olefile.isOleFile(str(path)):
        with olefile.OleFileIO(str(path)) as ole:
            for stream in ole.listdir():
                try:
                    data = ole.openstream(stream).read()
                except Exception:
                    continue
                raw_runs.extend(_scan_bytes(data, raw))
    else:
        # Plik nie jest OLE (lub brak olefile) — fallback: surowe wycinanie ciągów.
        raw_runs.extend(_scan_bytes(path.read_bytes(), raw))

    # Usuwanie duplikatów z zachowaniem kolejności (+ filtrowanie poza trybem raw).
    seen: set[str] = set()
    lines: list[str] = []
    for run in raw_runs:
        candidate = run.strip()
        if len(candidate) < 2:
            continue
        if not raw and not _looks_like_content(candidate):
            continue
        if candidate in seen:
            continue
        seen.add(candidate)
        lines.append(candidate)
    return lines


# --- Zapis wyników -----------------------------------------------------------

def write_txt(lines: list[str], dest: Path, source_name: str) -> None:
    header = f"# Etykieta: {source_name}\n# Wyodrębniony tekst ({len(lines)} pozycji)\n\n"
    dest.write_text(header + "\n".join(lines) + "\n", encoding="utf-8")


# --- CLI ---------------------------------------------------------------------

def collect_btw_files(target: Path) -> list[Path]:
    if target.is_dir():
        return sorted(p for p in target.rglob("*.btw"))
    if target.suffix.lower() == ".btw" or target.is_file():
        return [target]
    return []


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Konwersja etykiet BarTender (.btw) do plików tekstowych (.txt)."
    )
    parser.add_argument("input", help="Plik .btw lub folder z plikami .btw")
    parser.add_argument("-o", "--output-dir", help="Folder docelowy (domyślnie obok plików źródłowych)")
    parser.add_argument("--raw", action="store_true",
                        help="Bez filtrowania — wypisz WSZYSTKIE znalezione ciągi (do weryfikacji)")
    args = parser.parse_args(argv)

    if olefile is None:
        print("BŁĄD: brak biblioteki 'olefile'. Zainstaluj: pip install olefile",
              file=sys.stderr)
        return 2

    target = Path(args.input).expanduser()
    if not target.exists():
        print(f"BŁĄD: ścieżka nie istnieje: {target}", file=sys.stderr)
        return 2

    files = collect_btw_files(target)
    if not files:
        print(f"Nie znaleziono plików .btw w: {target}", file=sys.stderr)
        return 1

    out_dir = Path(args.output_dir).expanduser() if args.output_dir else None
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)

    count = 0
    for f in files:
        try:
            lines = extract_text_from_btw(f, raw=args.raw)
        except Exception as exc:
            print(f"  ! Pominięto {f.name}: {exc}", file=sys.stderr)
            continue

        base_dir = out_dir if out_dir else f.parent
        txt_path = base_dir / (f.stem + ".txt")
        write_txt(lines, txt_path, f.name)
        print(f"  ✓ {txt_path}  ({len(lines)} linii)")
        count += 1

    print(f"\nGotowe. Przetworzono {count} plików.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
