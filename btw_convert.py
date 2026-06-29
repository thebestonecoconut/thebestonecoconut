#!/usr/bin/env python3
"""
btw_convert.py — wyciąga tekst z plików etykiet BarTender (.btw) i zapisuje
go do pliku tekstowego (.txt) oraz do PDF.

Pliki .btw to dokumenty OLE2 (Compound File Binary Format) tworzone przez
program BarTender (Seagull Scientific). Wewnątrz przechowują m.in. teksty
etykiet (zwykle jako XML / tekst w kodowaniu UTF-16). Ten skrypt otwiera taki
plik, wydobywa z niego czytelne ciągi znaków i zapisuje je w formie tekstu/PDF.

UWAGA: pełne, "pixel-perfect" odwzorowanie wyglądu etykiety (kody kreskowe,
grafiki, układ) potrafi wygenerować TYLKO sam BarTender. Ten skrypt służy do
wyciągnięcia ZAWARTOŚCI TEKSTOWEJ etykiet.

Użycie:
    python btw_convert.py PLIK.btw                  # -> PLIK.txt i PLIK.pdf
    python btw_convert.py folder/                   # konwertuje wszystkie .btw
    python btw_convert.py folder/ -o wyniki/        # zapis do innego folderu
    python btw_convert.py folder/ --combined-pdf etykiety.pdf
    python btw_convert.py PLIK.btw --txt-only       # tylko .txt (bez PDF)
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

try:
    import olefile
except ImportError:  # pragma: no cover
    olefile = None


# --- Wyciąganie tekstu -------------------------------------------------------

# Sekwencje co najmniej 2 "drukowalnych" znaków. Obsługujemy polskie znaki.
_PRINTABLE = (
    r"[\x20-\x7E"
    r"\u00A0-\u024F"   # Latin-1 + Latin Extended-A/B (ą, ć, ę, ł, ó, ...)
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
    "BarTender", "Seagull", "btObject", "btField", "Microsoft", "OLE",
    "xmlns", "http://", "https://", "schemas", "GUID", "FontName",
)


# UTF-16LE: znak (młodszy bajt 0x20-0xFF) + starszy bajt 0x00-0x04
# (pokrywa ASCII, Latin-1 oraz polskie znaki z Latin Extended, U+0000–U+04FF).
_UTF16_RUN_RE = re.compile(rb"(?:[\x20-\xff][\x00-\x04]){2,}")
# Czysty ASCII (drukowalny) — łapie metadane/teksty zapisane jednobajtowo.
_ASCII_RUN_RE = re.compile(rb"[\x20-\x7e]{3,}")


def _decode_runs(data: bytes) -> list[str]:
    """Zwraca listę czytelnych ciągów znaków znalezionych w surowych bajtach.

    Szuka osobno tekstu UTF-16LE (najczęstszy w BarTenderze) oraz ASCII, dzięki
    czemu unika śmieci powstających przy "ślepym" dekodowaniu całego strumienia.
    """
    results: list[str] = []

    for raw in _UTF16_RUN_RE.findall(data):
        text = raw.decode("utf-16-le", errors="ignore")
        for match in _RUN_RE.findall(text):
            cleaned = match.strip()
            if cleaned:
                results.append(cleaned)

    for raw in _ASCII_RUN_RE.findall(data):
        text = raw.decode("ascii", errors="ignore")
        for match in _RUN_RE.findall(text):
            cleaned = match.strip()
            if cleaned:
                results.append(cleaned)

    return results


def _looks_like_content(s: str) -> bool:
    """Heurystyka: czy dany ciąg wygląda na realną treść etykiety."""
    if len(s) < 2:
        return False
    if _NOISE_RE.match(s):
        return False
    if any(k.lower() in s.lower() for k in _INTERNAL_KEYWORDS):
        return False
    # musi zawierać przynajmniej jedną literę lub cyfrę
    if not re.search(r"[0-9A-Za-z\u00A0-\u024F]", s):
        return False
    return True


def extract_text_from_btw(path: Path) -> list[str]:
    """Wyciąga uporządkowaną listę linii tekstu z pliku .btw."""
    raw_runs: list[str] = []

    if olefile is not None and olefile.isOleFile(str(path)):
        with olefile.OleFileIO(str(path)) as ole:
            for stream in ole.listdir():
                try:
                    data = ole.openstream(stream).read()
                except Exception:
                    continue
                raw_runs.extend(_decode_runs(data))
    else:
        # Plik nie jest OLE (lub brak olefile) — fallback: surowe wycinanie ciągów.
        raw_runs.extend(_decode_runs(path.read_bytes()))

    # Filtrowanie + usuwanie duplikatów z zachowaniem kolejności.
    seen: set[str] = set()
    lines: list[str] = []
    for run in raw_runs:
        candidate = run.strip()
        if not _looks_like_content(candidate):
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


def write_pdf(sections: list[tuple[str, list[str]]], dest: Path) -> None:
    """Tworzy PDF. `sections` to lista (nazwa_pliku, linie)."""
    from fpdf import FPDF

    pdf = FPDF(format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)

    # Czcionka z obsługą Unicode (jeśli dostępna w systemie), inaczej Helvetica.
    font_family = "Helvetica"
    unicode_font = _find_unicode_font()
    if unicode_font is not None:
        try:
            pdf.add_font("DejaVu", "", str(unicode_font))
            pdf.add_font("DejaVu", "B", str(unicode_font))
            font_family = "DejaVu"
        except Exception:
            font_family = "Helvetica"

    def out(text: str) -> str:
        if font_family == "Helvetica":
            # Helvetica = Latin-1; zastąp nieobsługiwane znaki.
            return text.encode("latin-1", "replace").decode("latin-1")
        return text

    width = pdf.epw  # efektywna szerokość strony (bez marginesów)
    for name, lines in sections:
        pdf.add_page()
        pdf.set_font(font_family, "B", 14)
        pdf.multi_cell(width, 8, out(f"Etykieta: {name}"), wrapmode="CHAR")
        pdf.ln(2)
        pdf.set_font(font_family, "", 11)
        if not lines:
            pdf.multi_cell(width, 6, out("(nie znaleziono tekstu)"), wrapmode="CHAR")
        for line in lines:
            pdf.multi_cell(width, 6, out("- " + line), wrapmode="CHAR")

    pdf.output(str(dest))


def _find_unicode_font() -> Path | None:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for c in candidates:
        if Path(c).exists():
            return Path(c)
    return None


# --- CLI ---------------------------------------------------------------------

def collect_btw_files(target: Path) -> list[Path]:
    if target.is_dir():
        return sorted(p for p in target.rglob("*.btw"))
    if target.suffix.lower() == ".btw" or target.is_file():
        return [target]
    return []


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Konwersja etykiet BarTender (.btw) do tekstu i PDF."
    )
    parser.add_argument("input", help="Plik .btw lub folder z plikami .btw")
    parser.add_argument("-o", "--output-dir", help="Folder docelowy (domyślnie obok plików źródłowych)")
    parser.add_argument("--txt-only", action="store_true", help="Zapisz tylko .txt (bez PDF)")
    parser.add_argument("--pdf-only", action="store_true", help="Zapisz tylko PDF (bez .txt)")
    parser.add_argument("--combined-pdf", metavar="PLIK.pdf",
                        help="Zapisz jeden wspólny PDF dla wszystkich etykiet")
    args = parser.parse_args(argv)

    if olefile is None:
        print("BŁĄD: brak biblioteki 'olefile'. Zainstaluj: pip install olefile fpdf2",
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

    sections: list[tuple[str, list[str]]] = []
    for f in files:
        try:
            lines = extract_text_from_btw(f)
        except Exception as exc:
            print(f"  ! Pominięto {f.name}: {exc}", file=sys.stderr)
            continue

        sections.append((f.name, lines))
        base_dir = out_dir if out_dir else f.parent

        if not args.pdf_only:
            txt_path = base_dir / (f.stem + ".txt")
            write_txt(lines, txt_path, f.name)
            print(f"  ✓ {txt_path}  ({len(lines)} linii)")

        if not args.txt_only and not args.combined_pdf:
            pdf_path = base_dir / (f.stem + ".pdf")
            write_pdf([(f.name, lines)], pdf_path)
            print(f"  ✓ {pdf_path}")

    if args.combined_pdf and not args.txt_only:
        combined = Path(args.combined_pdf).expanduser()
        if out_dir and not combined.is_absolute():
            combined = out_dir / combined
        write_pdf(sections, combined)
        print(f"  ✓ Wspólny PDF: {combined}")

    print(f"\nGotowe. Przetworzono {len(sections)} plików.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
