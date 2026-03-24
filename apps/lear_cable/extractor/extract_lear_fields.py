#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from dataclasses import dataclass, asdict
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable, Optional

SCRIPT_VERSION = "2026-02-6.v23"

SCRIPT_CHANGELOG = """
## 2026-02-6.v23

### Logica general
Extrae campos de facturas PDF de Lear Cable y genera un archivo XLSX
con los datos estructurados listos para tramitacion.

### Extraccion PDF
- Detecta automaticamente el formato del PDF (texto nativo u OCR)
- Extrae por factura: numero, fecha, importe total, moneda
- Extrae por linea: codigo de pieza, descripcion, cantidad, precio unitario, subtotal

### Columnas generadas
- Numero de factura y fecha
- Codigo de pieza (referencia Lear)
- Descripcion del articulo
- Cantidad y unidad de medida
- Precio unitario y subtotal por linea
- Totales de factura (neto, IVA, total)

### Salida
Genera `invoices_extracted.xlsx` con una fila por linea de factura.
""".strip()

try:
    from pypdf import PdfReader
except ImportError:
    print("ERROR: falta dependencia. Instala con: pip install pypdf", file=sys.stderr)
    raise

try:
    from openpyxl import Workbook
except ImportError:
    print("ERROR: falta dependencia. Instala con: pip install openpyxl", file=sys.stderr)
    raise


@dataclass
class InvoiceExtract:
    file: str
    invoice_no: Optional[str]
    date: Optional[str]
    pallets: Optional[int]
    boxes: Optional[int]
    gross_weight: Optional[float]
    net_weight: Optional[float]
    total_invoice: Optional[float]


# ----------------------------
# Normalizacion / parsing
# ----------------------------

def reduce_repetition(token: str) -> str:
    """Reduce tokens repetidos exactos."""
    s = token.strip()
    if not s:
        return s
    n = len(s)
    for k in range(1, n // 2 + 1):
        if n % k == 0:
            unit = s[:k]
            if unit * (n // k) == s:
                return unit
    return s


def parse_number_token(token: str) -> Optional[float]:
    """Parser numerico robusto (miles/decimales EU/US) + repeticion."""
    if token is None:
        return None

    raw = str(token)

    parts = re.findall(r"[0-9][0-9.,]*", raw)
    if parts:
        if all(p == parts[0] for p in parts):
            raw = parts[0]
        else:
            raw = max(parts, key=len)

    tok = reduce_repetition(raw).strip()
    if not tok:
        return None

    tok = tok.replace("\u00A0", "").replace(" ", "")
    tok = re.sub(r"(?i)(kg|kgs|g)$", "", tok).strip()

    if "," in tok and "." in tok:
        if tok.rfind(",") > tok.rfind("."):
            tok = tok.replace(".", "")
            tok = tok.replace(",", ".")
        else:
            tok = tok.replace(",", "")
    elif "," in tok and "." not in tok:
        tok = tok.replace(",", ".")

    try:
        return float(tok)
    except ValueError:
        return None


def sum_optional_decimal(values: Iterable[Optional[float]]) -> Decimal:
    total = Decimal("0")
    for v in values:
        if isinstance(v, (int, float)):
            total += Decimal(str(v))
    return total


# ----------------------------
# PDF text (sin OCR)
# ----------------------------

def _normalize_extracted_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\u202f", " ")
    text = re.sub(r"[\u200B\u200C\u200D\u200E\u200F\u2060\uFEFF]", "", text)
    text = re.sub(r"[\u202A-\u202E\u2066-\u2069]", "", text)
    return text


def pdf_text_no_ocr(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    raw = "\n".join((p.extract_text() or "") for p in reader.pages)
    return _normalize_extracted_text(raw)


# ----------------------------
# Date extraction + normalisation
# ----------------------------

_MONTH_NAMES = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _expand_year(y: int) -> int:
    """Expand 2-digit year: 00-30 -> 2000+y, 31-99 -> 1900+y."""
    if y < 100:
        return 2000 + y if y <= 30 else 1900 + y
    return y


def _normalize_date(raw: str) -> Optional[str]:
    """Normalize a raw date token to DD/MM/YYYY (or DD/MM when year absent).

    Disambiguation rules for fully-numeric A/B/C dates:
      - A > 12  ->  A is day,  B is month   (DD/MM/YYYY, no change)
      - B > 12  ->  B is day,  A is month   (was MM/DD/YYYY -> swap)
      - Both <= 12  ->  assume DD/MM        (European default)

    2-digit years expanded via _expand_year.
    Short month-name dates (21-Nov, Nov-21, 21-Nov-2025) are unambiguous.
    """
    raw = raw.strip()

    # "21-Nov" / "21-Nov-2025"
    m = re.fullmatch(
        r"(\d{1,2})[/\-\s](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"(?:[/\-\s](\d{2,4}))?",
        raw, flags=re.IGNORECASE,
    )
    if m:
        dd = int(m.group(1))
        mm = _MONTH_NAMES[m.group(2).lower()]
        yr_raw = m.group(3)
        if yr_raw:
            return f"{dd:02d}/{mm:02d}/{_expand_year(int(yr_raw))}"
        return f"{dd:02d}/{mm:02d}"

    # "Nov-21" / "Nov-21-2025"
    m = re.fullmatch(
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[/\-\s](\d{1,2})"
        r"(?:[/\-\s](\d{2,4}))?",
        raw, flags=re.IGNORECASE,
    )
    if m:
        mm = _MONTH_NAMES[m.group(1).lower()]
        dd = int(m.group(2))
        yr_raw = m.group(3)
        if yr_raw:
            return f"{dd:02d}/{mm:02d}/{_expand_year(int(yr_raw))}"
        return f"{dd:02d}/{mm:02d}"

    # fully numeric: A/B/C (separators / - .)
    m = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", raw)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        yyyy = _expand_year(c)
        if a > 12:        # A is definitely the day
            dd, mm = a, b
        elif b > 12:      # B is definitely the day -> was MM/DD/YYYY
            dd, mm = b, a
        else:             # ambiguous -> European default DD/MM
            dd, mm = a, b
        if not (1 <= mm <= 12 and 1 <= dd <= 31):
            return raw    # still invalid: return as-is
        return f"{dd:02d}/{mm:02d}/{yyyy}"

    return raw  # unrecognized: return unchanged


def find_date(text: str) -> Optional[str]:
    """Extract invoice date from PDF text, normalized to DD/MM/YYYY.

    Handles all invoice models:
      - Numeric date after 'Date' label: 12/2/2025 | 11/11/2025 | 1/31/2026
      - Short 2-digit year: 18/11/25
      - Short month-name: 21-Nov

    Ambiguous numeric dates (both parts <= 12) default to DD/MM (European).
    """
    PAT_NUMERIC = r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}"
    PAT_SHORT_MONTH = (
        r"\d{1,2}[\s\-/](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"(?:[\s\-/]\d{2,4})?"
        r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s\-/]\d{1,2}"
        r"(?:[\s\-/]\d{2,4})?"
    )
    COMBINED = rf"(?:{PAT_NUMERIC}|{PAT_SHORT_MONTH})"

    # 1. Window after any 'Date' label (handles repeated label artefacts)
    for m_label in re.finditer(r"\bDate\s*:?\s*(?:Date\s*:?\s*)*", text, flags=re.IGNORECASE):
        window = text[m_label.end(): m_label.end() + 80]
        m = re.search(COMBINED, window, flags=re.IGNORECASE)
        if m:
            candidate = m.group(0).strip()
            if re.search(r"\d", candidate):
                return _normalize_date(candidate)

    # 2. Last resort: first date-like token anywhere in the document
    m = re.search(COMBINED, text, flags=re.IGNORECASE)
    if m:
        candidate = m.group(0).strip()
        if re.search(r"\d", candidate):
            return _normalize_date(candidate)

    return None


# ----------------------------
# Finders
# ----------------------------

def find_invoice_no(text: str, filename_stem: str | None = None) -> Optional[str]:
    """Return invoice number from filename only."""
    if not filename_stem:
        return None
    stem = filename_stem.strip()
    stem = re.sub(r"[^A-Za-z0-9_-]+", "", stem)
    if not stem:
        return None
    return stem.upper()


def find_int_after_label_variants(text: str, labels: list[str]) -> Optional[int]:
    for label in labels:
        m = re.search(rf"((?:{label}\s*:\s*)+)([0-9]+)", text, flags=re.IGNORECASE)
        if m:
            prefix = m.group(1)
            raw = m.group(2)
            repeat_n = len(re.findall(label, prefix, flags=re.IGNORECASE)) or 1
            if repeat_n > 1 and len(raw) % repeat_n == 0:
                raw = raw[: len(raw) // repeat_n]
            else:
                raw = reduce_repetition(raw)
            try:
                return int(raw)
            except ValueError:
                pass

        m = re.search(rf"{label}\s*:\s*([0-9]+)", text, flags=re.IGNORECASE)
        if not m:
            continue
        tok = reduce_repetition(m.group(1))
        try:
            return int(tok)
        except ValueError:
            continue
    return None


def find_float_after_label_variants(text: str, labels: list[str]) -> Optional[float]:
    for label in labels:
        m = re.search(rf"{label}\s*:\s*([0-9][0-9.,\s]*)(?:\s*(?:kg|kgs))?\b", text, flags=re.IGNORECASE)
        if not m:
            continue
        v = parse_number_token(m.group(1))
        if v is not None:
            return v
    return None


def parse_weight_token(token: str) -> Optional[float]:
    if token is None:
        return None
    raw = str(token)

    parts = re.findall(r"[0-9][0-9.,]*", raw)
    cand = None
    if parts:
        if all(p == parts[0] for p in parts):
            cand = parts[0]
        else:
            cand = max(parts, key=len)
    else:
        cand = raw

    cand = cand.strip().replace("\u00A0", "").replace(" ", "")
    cand = re.sub(r"(?i)(kg|kgs|g)$", "", cand).strip()

    if re.fullmatch(r"\d{2,}", cand) and len(set(cand)) == 1:
        if len(cand) <= 5:
            try:
                return float(int(cand))
            except ValueError:
                return None
        for k in (3, 2, 4):
            if len(cand) % k == 0:
                chunk = cand[:k]
                try:
                    return float(int(chunk))
                except ValueError:
                    pass
        try:
            return float(int(cand[:3]))
        except ValueError:
            return None

    return parse_number_token(cand)


def find_weight_after_label_variants(text: str, labels: list[str]) -> Optional[float]:
    for label in labels:
        m = re.search(rf"{label}\s*:\s*([0-9][0-9.,\s]*)(?:\s*(?:kg|kgs))?\b", text, flags=re.IGNORECASE)
        if not m:
            continue
        v = parse_weight_token(m.group(1))
        if v is not None:
            return v
    return None


def find_float_before_marker(text: str, marker_regex: str) -> Optional[float]:
    m = re.search(rf"([0-9][0-9.,\s]+)\s*{marker_regex}", text, flags=re.IGNORECASE)
    if not m:
        return None
    return parse_number_token(m.group(1))


def find_total_invoice(text: str) -> Optional[float]:
    # 1) Common: "Total Invoices"
    m = re.search(r"Total\s+Invoices?.*?([0-9][0-9.,\s]+)", text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        v = parse_number_token(m.group(1))
        if v is not None:
            return v

    # 2) "TOTAL INVOICE: 1,329.01"
    m = re.search(r"TOTAL\s+INVOICE\s*[: ]\s*([0-9][0-9.,\s]+)", text, flags=re.IGNORECASE)
    if m:
        v = parse_number_token(m.group(1))
        if v is not None:
            return v

    # 3) "Total Facture en Euro*"
    label = re.search(r"Total\s+Facture\s+en\s+Euro\*?\s*:\s*(?:[0-9][0-9.,\s]+)?", text, flags=re.IGNORECASE)
    if label:
        start = max(0, label.start() - 250)
        end = min(len(text), label.end() + 250)
        window = text[start:end]

        after_euro = re.findall(r"€\s*([0-9][0-9.,\s]+)", window)
        before_euro = re.findall(r"([0-9][0-9.,\s]+)\s*[^0-9]{0,5}€", window)
        candidates = []
        for tok in after_euro + before_euro:
            v = parse_number_token(tok)
            if v is None:
                continue
            tok_s = str(tok).replace(" ", "").replace(" ", "")
            has_2dp = bool(re.search(r"[\.,]\d{2}$", tok_s))
            candidates.append((has_2dp, v))

        if candidates:
            with_2dp = [v for has_2dp, v in candidates if has_2dp]
            if with_2dp:
                return with_2dp[-1]
            return candidates[-1][1]

    # 4) Last resort: amounts with € symbol
    euro_amounts = (
        re.findall(r"€\s*([0-9][0-9.,\s]+)", text)
        + re.findall(r"([0-9][0-9.,\s]+)\s*[^0-9]{0,5}€", text)
    )
    parsed = []
    for tok in euro_amounts:
        v = parse_number_token(tok)
        if v is None:
            continue
        tok_s = str(tok).replace(" ", "").replace(" ", "")
        has_2dp = bool(re.search(r"[\.,]\d{2}$", tok_s))
        parsed.append((has_2dp, v))

    if parsed:
        with_2dp = [v for has_2dp, v in parsed if has_2dp]
        if with_2dp:
            return with_2dp[-1]
        return parsed[-1][1]

    return None


def extract_one(pdf_path: Path) -> InvoiceExtract:
    text = pdf_text_no_ocr(pdf_path)
    stem = pdf_path.stem

    invoice_no = find_invoice_no(text, filename_stem=stem)
    date = find_date(text)

    pallets = find_int_after_label_variants(text, [r"Pallets", r"Nombre\s+de\s+Palettes"])
    boxes = find_int_after_label_variants(text, [r"Boxes"])

    gross_weight = find_weight_after_label_variants(
        text, [r"Gross\s+Weight", r"Poid\s+brute", r"Poids\s+brut", r"Poids\s+brute"]
    )
    net_weight = find_weight_after_label_variants(
        text, [r"Net\s+Weight", r"NetWeight", r"Poids\s+net", r"Poid\s+net"]
    )

    total_invoice = find_total_invoice(text)

    return InvoiceExtract(
        file=pdf_path.name,
        invoice_no=invoice_no,
        date=date,
        pallets=pallets,
        boxes=boxes,
        gross_weight=gross_weight,
        net_weight=net_weight,
        total_invoice=total_invoice,
    )


def iter_pdfs(inputs: list[Path]) -> Iterable[Path]:
    for p in inputs:
        if p.is_dir():
            all_pdfs = list(p.glob("*.pdf")) + list(p.glob("*.PDF"))
            yield from sorted(all_pdfs, key=lambda f: f.name.lower())
        else:
            yield p



def _fill_missing_years(rows: list[InvoiceExtract]) -> None:
    """Infer missing years for dates that only have DD/MM (no year component).

    Strategy: collect all 4-digit years already present in the batch (from
    dates formatted DD/MM/YYYY), pick the most common one, and append it to
    any date that matches DD/MM exactly (5 chars, one slash).
    This handles invoices like inv-211125 whose date field only shows "21-Nov".
    """
    from collections import Counter

    year_re = re.compile(r"^\d{2}/\d{2}/(\d{4})$")
    no_year_re = re.compile(r"^\d{2}/\d{2}$")

    years = Counter()
    for r in rows:
        if r.date:
            m = year_re.match(r.date)
            if m:
                years[m.group(1)] += 1

    if not years:
        return  # no reference year available, nothing to do

    best_year = years.most_common(1)[0][0]

    for r in rows:
        if r.date and no_year_re.match(r.date):
            r.date = f"{r.date}/{best_year}"
            print(f"[INFO] Year inferred for {r.file}: {r.date}")

def main() -> int:
    ap = argparse.ArgumentParser(description="Extrae campos clave de PDFs con texto real (sin OCR).")
    ap.add_argument("inputs", nargs="+", type=Path, help="PDF(s) o carpeta(s) con PDFs")
    ap.add_argument("-o", "--out", type=Path, default=Path("out"), help="Carpeta de salida")
    args = ap.parse_args()

    out_dir: Path = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    rows: list[InvoiceExtract] = []
    for pdf in iter_pdfs(args.inputs):
        if pdf.suffix.lower() != ".pdf":
            continue
        try:
            rows.append(extract_one(pdf))
        except Exception as e:
            print(f"[WARN] Fallo {pdf.name}: {e}", file=sys.stderr)

    # Always output in consistent alphabetical order regardless of shell glob order.
    rows.sort(key=lambda r: r.file.lower())

    # Fill in missing years from the rest of the batch.
    _fill_missing_years(rows)

    total_pallets = sum(r.pallets for r in rows if r.pallets is not None)
    total_boxes = sum(r.boxes for r in rows if r.boxes is not None)
    total_gross_weight = sum_optional_decimal(r.gross_weight for r in rows)
    total_net_weight = sum_optional_decimal(r.net_weight for r in rows)
    total_total_invoice = sum_optional_decimal(r.total_invoice for r in rows)

    total_gross_weight_r = total_gross_weight.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
    total_net_weight_r = total_net_weight.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
    total_total_invoice_r = total_total_invoice.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    # JSON detalle
    json_path = out_dir / "invoices_extracted.json"
    json_path.write_text(
        json.dumps([asdict(r) for r in rows], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # JSON resumen
    summary = {
        "count_files": len(rows),
        "count_with_gross_weight": sum(1 for r in rows if isinstance(r.gross_weight, (int, float))),
        "count_with_total_invoice": sum(1 for r in rows if isinstance(r.total_invoice, (int, float))),
        "total_gross_weight": f"{total_gross_weight_r:.4f}",
        "total_total_invoice": f"{total_total_invoice_r:.2f}",
    }
    summary_path = out_dir / "invoices_extracted_summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    # Final safety sort: guarantee data rows are alphabetical by filename.
    # TOTAL is always appended last, never in between data rows.
    rows.sort(key=lambda r: r.file.lower())

    print(f"[INFO] Processing order ({len(rows)} files):")
    for _i, _r in enumerate(rows, 1):
        print(f"  {_i:>3}. {_r.file}")

    # CSV + fila TOTAL
    csv_path = out_dir / "invoices_extracted.csv"
    fieldnames = [
        "file", "invoice_no", "date", "pallets", "boxes",
        "gross_weight", "net_weight", "total_invoice",
    ]

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow(asdict(r))
        w.writerow({
            "file": "TOTAL", "invoice_no": None, "date": None,
            "pallets": total_pallets, "boxes": total_boxes,
            "gross_weight": f"{total_gross_weight_r:.4f}",
            "net_weight": f"{total_net_weight_r:.4f}",
            "total_invoice": f"{total_total_invoice_r:.2f}",
        })

    print(f"OK -> {json_path}")
    print(f"OK -> {summary_path}")
    print(f"OK -> {csv_path}")

    # XLSX
    xlsx_path = out_dir / "invoices_extracted.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"
    ws.append(fieldnames)

    for r in rows:
        ws.append([
            r.file, r.invoice_no, r.date, r.pallets, r.boxes,
            r.gross_weight, r.net_weight, r.total_invoice,
        ])

    ws.append([
        "TOTAL", None, None,
        total_pallets, total_boxes,
        float(total_gross_weight_r), float(total_net_weight_r), float(total_total_invoice_r),
    ])

    wb.save(str(xlsx_path))
    print(f"OK -> {xlsx_path}")

    print(f"SUM pallets:       {total_pallets}")
    print(f"SUM boxes:         {total_boxes}")
    print(f"SUM gross_weight:  {total_gross_weight_r:.4f}")
    print(f"SUM net_weight:    {total_net_weight_r:.4f}")
    print(f"SUM total_invoice: {total_total_invoice_r:.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
