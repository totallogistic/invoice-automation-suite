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

SCRIPT_VERSION = "2026-03-5.v21"

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
    """Parser numerico robusto (miles/decimales EU/US) + repeticion.

    Algunas extracciones (pypdf) devuelven el mismo numero varias veces
    separado por espacios o concatenado. En esos casos nos quedamos con
    el candidato numerico mas razonable antes de normalizar.
    """
    if token is None:
        return None

    raw = str(token)

    # Si vienen varios numeros (p.ej. "6734.73 6734.736734.73 6734.73"),
    # suele ser el mismo valor repetido. Elegimos el candidato mas largo.
    parts = re.findall(r"[0-9][0-9.,]*", raw)
    if parts:
        # si todos son iguales, usa uno
        if all(p == parts[0] for p in parts):
            raw = parts[0]
        else:
            raw = max(parts, key=len)

    tok = reduce_repetition(raw).strip()
    if not tok:
        return None

    # quita espacios y NBSP usados como separador de miles
    tok = tok.replace("\u00A0", "").replace(" ", "")

    # recorta sufijos tipo "kg" si vienen pegados (ej: "191kg")
    tok = re.sub(r"(?i)(kg|kgs|g)$", "", tok).strip()

    if "," in tok and "." in tok:
        # el ultimo separador suele ser el decimal
        if tok.rfind(",") > tok.rfind("."):
            # 1.329,01 -> 1329.01
            tok = tok.replace(".", "")
            tok = tok.replace(",", ".")
        else:
            # 1,329.01 -> 1329.01
            tok = tok.replace(",", "")
    elif "," in tok and "." not in tok:
        # 1329,01 -> 1329.01
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
    """Normaliza texto extraido (sin OCR) para hacer los regex mas fiables."""
    # normaliza espacios no-separables
    text = text.replace("\u00a0", " ").replace("\u202f", " ")
    # elimina caracteres de ancho cero y BOM
    text = re.sub(r"[\u200B\u200C\u200D\u200E\u200F\u2060\uFEFF]", "", text)
    # elimina controles bidi
    text = re.sub(r"[\u202A-\u202E\u2066-\u2069]", "", text)
    return text


def pdf_text_no_ocr(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    raw = "\n".join((p.extract_text() or "") for p in reader.pages)
    return _normalize_extracted_text(raw)


# ----------------------------
# Finders
# ----------------------------

def find_invoice_no(text: str, filename_stem: str | None = None) -> Optional[str]:
    """Return invoice number from filename only.

    We intentionally avoid extracting invoice number from PDF text because some
    invoice layouts yield false positives like 'NUMBER'. The filename (stem)
    is the source of truth and is expected to match the invoice id.
    """
    if not filename_stem:
        return None
    stem = filename_stem.strip()
    stem = re.sub(r"[^A-Za-z0-9_-]+", "", stem)
    if not stem:
        return None
    return stem.upper()


def find_date(text: str) -> Optional[str]:
    """Extract invoice date from PDF text.

    Handles the date formats observed across all invoice models:

    Model A (DM043899, 02343986, DS307207):
        Numeric date after 'Date' label:
        e.g.  12/2/2025 | 11/11/2025 | 1/31/2026

    Model B (S122412):
        Numeric date in 'Date' column of a table row:
        e.g.  18/11/25   (2-digit year)

    Model C (inv-211125):
        Short date after 'Date :' label:
        e.g.  21-Nov

    Strategy:
      1. Look for a numeric date (d/m/y or m/d/y, 2- or 4-digit year)
         within a window following any 'Date' label.
      2. Fall back to a short month-name date (e.g. 21-Nov or Nov-21)
         near a 'Date' label.
      3. Last resort: first plausible date-like token in the whole document.

    The raw string is returned as-is (no normalisation to ISO) so callers
    can decide how to interpret ambiguous formats.
    """
    # ------------------------------------------------------------------ #
    # Pattern definitions
    # ------------------------------------------------------------------ #
    # Full numeric date: 1-2 digit day-or-month, separator, 1-2 digit
    # day-or-month, separator, 2- or 4-digit year.
    # Separators: / - .
    PAT_NUMERIC = r"\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4}"

    # Short month-name date: "21-Nov", "21/Nov", "Nov-21", "21 Nov 2025", …
    PAT_SHORT_MONTH = (
        r"\d{1,2}[\s\-/](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"(?:[\s\-/]\d{2,4})?"
        r"|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\s\-/]\d{1,2}"
        r"(?:[\s\-/]\d{2,4})?"
    )

    COMBINED = rf"(?:{PAT_NUMERIC}|{PAT_SHORT_MONTH})"

    # ------------------------------------------------------------------ #
    # 1. Search in a window after every 'Date' label (robust to repetition)
    # ------------------------------------------------------------------ #
    # Labels can appear repeated (e.g. "Date\nDate\n12/2/2025") or with colon.
    for m_label in re.finditer(
        r"\bDate\s*:?\s*(?:Date\s*:?\s*)*",
        text,
        flags=re.IGNORECASE,
    ):
        # Look ahead up to 80 chars for a date token
        window = text[m_label.end(): m_label.end() + 80]
        m = re.search(COMBINED, window, flags=re.IGNORECASE)
        if m:
            candidate = m.group(0).strip()
            # Exclude obviously wrong hits like lone "Date" repeated
            if re.search(r"\d", candidate):
                return candidate

    # ------------------------------------------------------------------ #
    # 2. Last resort: first date-like token in the whole document
    # ------------------------------------------------------------------ #
    m = re.search(COMBINED, text, flags=re.IGNORECASE)
    if m:
        candidate = m.group(0).strip()
        if re.search(r"\d", candidate):
            return candidate

    return None


def find_int_after_label_variants(text: str, labels: list[str]) -> Optional[int]:
    """Find integer after any of the label regex variants like 'Pallets' / 'Nombre de Palettes'.

    Robust against duplicated text layers, e.g.:
      'Pallets:Pallets:Pallets:Pallets: 11111111'  -> 11

    Strategy (from v16):
      - If the label itself is repeated N times in the matched prefix, and the digit chunk
        length is divisible by N, assume the digits are repeated N times too and shrink.
      - Otherwise fall back to reduce_repetition().
    """
    for label in labels:
        # Robust case: label repeated N times then a (possibly repeated) digit chunk
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

        # Simple fallback: single label occurrence
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
    """Parse de pesos (gross/net) evitando colapsar '444' -> '4'.

    Algunos PDFs devuelven pesos como repeticiones del mismo dígito (p.ej. '444444444').
    Para pesos, si el número son dígitos repetidos, usamos el chunk plausible (normalmente 3 dígitos)
    y NO aplicamos reduce_repetition.
    """
    if token is None:
        return None
    raw = str(token)

    # Extrae candidatos numéricos
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

    # Caso especial: solo dígitos y todos iguales (p.ej. '444' o '444444444')
    if re.fullmatch(r"\d{2,}", cand) and len(set(cand)) == 1:
        # Si ya es un número razonable (2-5 dígitos), úsalo tal cual
        if len(cand) <= 5:
            try:
                return float(int(cand))
            except ValueError:
                return None
        # Si es largo, intenta chunk de 3 (preferido) o 2 si encaja
        for k in (3, 2, 4):
            if len(cand) % k == 0:
                chunk = cand[:k]
                try:
                    return float(int(chunk))
                except ValueError:
                    pass
        # Fallback: primeros 3 dígitos
        try:
            return float(int(cand[:3]))
        except ValueError:
            return None

    # Para el resto, usa el parser general existente
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
    # admite , . y espacios
    m = re.search(rf"([0-9][0-9.,\s]+)\s*{marker_regex}", text, flags=re.IGNORECASE)
    if not m:
        return None
    return parse_number_token(m.group(1))


def find_total_invoice(text: str) -> Optional[float]:
    # 1) Common: "Total Invoices" (mantener prioridad)
    m = re.search(r"Total\s+Invoices?.*?([0-9][0-9.,\s]+)", text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        v = parse_number_token(m.group(1))
        if v is not None:
            return v

    # 2) Rare: "TOTAL INVOICE: 1,329.01"
    m = re.search(r"TOTAL\s+INVOICE\s*[: ]\s*([0-9][0-9.,\s]+)", text, flags=re.IGNORECASE)
    if m:
        v = parse_number_token(m.group(1))
        if v is not None:
            return v

    # 3) Rare (FR): "Total Facture en Euro*".
    #    En algunos PDFs el valor NO queda en la misma linea (p.ej. el numero aparece justo antes y con el simbolo € al final: "808.88€").
    label = re.search(r"Total\s+Facture\s+en\s+Euro\*?\s*:\s*(?:[0-9][0-9.,\s]+)?", text, flags=re.IGNORECASE)
    if label:
        start = max(0, label.start() - 250)
        end = min(len(text), label.end() + 250)
        window = text[start:end]

        # Candidatos en ventana (dos formatos): "808.88€" y "€ 808.88"
        after_euro = re.findall(r"€\s*([0-9][0-9.,\s]+)", window)
        before_euro = re.findall(r"([0-9][0-9.,\s]+)\s*[^0-9]{0,5}€", window)
        candidates = []
        for tok in after_euro + before_euro:
            v = parse_number_token(tok)
            if v is None:
                continue
            # Heuristica: prioriza importes con 2 decimales en el token original (típico de totales)
            tok_s = str(tok).replace(" ", "").replace(" ", "")
            has_2dp = bool(re.search(r"[\.,]\d{2}$", tok_s))
            candidates.append((has_2dp, v))

        if candidates:
            # primero los que tienen 2 decimales, dentro de eso el ultimo encontrado suele ser el total
            with_2dp = [v for has_2dp, v in candidates if has_2dp]
            if with_2dp:
                return with_2dp[-1]
            return candidates[-1][1]

    # 4) Ultimo recurso: importes con simbolo € en todo el documento.
    #    Capturamos ambos formatos: "€ 1,329.01" y "1,329.01€".
    euro_amounts = re.findall(r"€\s*([0-9][0-9.,\s]+)", text) + re.findall(r"([0-9][0-9.,\s]+)\s*[^0-9]{0,5}€", text)
    parsed = []
    for tok in euro_amounts:
        v = parse_number_token(tok)
        if v is None:
            continue
        # evita capturar enteros enormes sin decimales si hay alternativas con decimales
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

    # common EN vs FR
    pallets = find_int_after_label_variants(text, [r"Pallets", r"Nombre\s+de\s+Palettes"])
    boxes = find_int_after_label_variants(text, [r"Boxes"])

    gross_weight = find_weight_after_label_variants(text, [r"Gross\s+Weight", r"Poid\s+brute", r"Poids\s+brut", r"Poids\s+brute"])
    net_weight = find_weight_after_label_variants(text, [r"Net\s+Weight", r"NetWeight", r"Poids\s+net", r"Poid\s+net"])

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
            yield from sorted(p.glob("*.pdf"))
            yield from sorted(p.glob("*.PDF"))
        else:
            yield p


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

    # Totales (solo redondeamos el TOTAL, no las facturas)
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

    # CSV + fila TOTAL
    csv_path = out_dir / "invoices_extracted.csv"
    fieldnames = [
        "file",
        "invoice_no",
        "date",
        "pallets",
        "boxes",
        "gross_weight",
        "net_weight",
        "total_invoice",
    ]

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow(asdict(r))

        w.writerow(
            {
                "file": "TOTAL",
                "invoice_no": None,
                "date": None,
                "pallets": total_pallets,
                "boxes": total_boxes,
                "gross_weight": f"{total_gross_weight_r:.4f}",
                "net_weight": f"{total_net_weight_r:.4f}",
                "total_invoice": f"{total_total_invoice_r:.2f}",
            }
        )

    print(f"OK -> {json_path}")
    print(f"OK -> {summary_path}")
    print(f"OK -> {csv_path}")

    # XLSX con la misma estructura que CSV
    xlsx_path = out_dir / "invoices_extracted.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"

    # Header
    ws.append(fieldnames)

    # Data rows
    for r in rows:
        row_data = [
            r.file,
            r.invoice_no,
            r.date,
            r.pallets,
            r.boxes,
            r.gross_weight,
            r.net_weight,
            r.total_invoice,
        ]
        ws.append(row_data)

    # TOTAL row
    total_row = [
        "TOTAL",
        None,
        None,
        total_pallets,
        total_boxes,
        float(total_gross_weight_r),
        float(total_net_weight_r),
        float(total_total_invoice_r),
    ]
    ws.append(total_row)

    wb.save(str(xlsx_path))
    print(f"OK -> {xlsx_path}")

    print(f"SUM pallets: {total_pallets}")
    print(f"SUM boxes: {total_boxes}")
    print(f"SUM gross_weight: {total_gross_weight_r:.4f}")
    print(f"SUM net_weight: {total_net_weight_r:.4f}")
    print(f"SUM total_invoice: {total_total_invoice_r:.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
