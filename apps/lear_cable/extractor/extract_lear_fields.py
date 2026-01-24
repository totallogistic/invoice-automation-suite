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

try:
    from pypdf import PdfReader
except ImportError:
    print("ERROR: falta dependencia. Instala con: pip install pypdf", file=sys.stderr)
    raise

# XLSX export (optional dependency in many environments, but available in yours)
try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None  # type: ignore

SCRIPT_VERSION = "2026-01-23.v15"
SCRIPT_DESCRIPTION = "Extract core invoice fields from text-based PDFs (no OCR). Exports JSON + CSV + XLSX with a fixed schema."

# Output schema (CSV/XLSX columns)
FIELDNAMES = [
    "file",
    "invoice_no",
    "pallets",
    "boxes",
    "gross_weight",
    "taxable_amount",
    "total_invoice",
]


@dataclass
class InvoiceExtract:
    file: str
    invoice_no: Optional[str]
    pallets: Optional[int]
    boxes: Optional[int]
    gross_weight: Optional[float]
    taxable_amount: Optional[float]
    total_invoice: Optional[float]


def reduce_repetition(token: str) -> str:
    """Reduce repetición exacta de un token.
    Ejemplos:
      - '3333' -> '3'
      - '65656565' -> '65'
      - '95007.8795007.87...' (si es repetición exacta) -> '95007.87'
    """
    s = (token or "").strip()
    if not s:
        return s

    n = len(s)
    for k in range(1, n // 2 + 1):
        if n % k == 0:
            unit = s[:k]
            if unit * (n // k) == s:
                return unit
    return s


def pdf_text_no_ocr(pdf_path: Path) -> str:
    """Extrae texto del PDF (sin OCR)."""
    reader = PdfReader(str(pdf_path))
    return "\n".join((p.extract_text() or "") for p in reader.pages)


def find_invoice_no(text: str, filename_stem: str | None = None) -> Optional[str]:
    """Busca Invoice No tipo DM######. Prioriza el stem del filename si coincide."""
    matches = re.findall(r"DM\d{6}", text)
    if not matches:
        return None
    if filename_stem:
        for m in matches:
            if m == filename_stem:
                return m
    return matches[0]


def find_float_before_marker(text: str, marker_regex: str) -> Optional[float]:
    """Encuentra un float inmediatamente antes de un marcador regex (p.ej. 'Vat %')."""
    m = re.search(rf"([0-9][0-9.]+)\s*{marker_regex}", text, flags=re.IGNORECASE)
    if not m:
        return None
    tok = reduce_repetition(m.group(1))
    try:
        return float(tok)
    except ValueError:
        return None


def find_int_after_label(text: str, label: str) -> Optional[int]:
    """Encuentra un int en formato 'Label: 123'."""
    m = re.search(rf"{re.escape(label)}\s*:\s*([0-9]+)", text, flags=re.IGNORECASE)
    if not m:
        return None
    tok = reduce_repetition(m.group(1))
    try:
        return int(tok)
    except ValueError:
        return None


def find_float_after_label(text: str, label: str) -> Optional[float]:
    """Encuentra un float en formato 'Label: 123.45'."""
    m = re.search(rf"{re.escape(label)}\s*:\s*([0-9.]+)", text, flags=re.IGNORECASE)
    if not m:
        return None
    tok = reduce_repetition(m.group(1))
    try:
        return float(tok)
    except ValueError:
        return None


def parse_amount(token: str) -> Optional[float]:
    """Parsea importes con separadores comunes (1,329.01 / 1329,01 / 1329.01)."""
    tok = reduce_repetition(token).strip().replace(" ", "")
    if not tok:
        return None

    # Caso 1: 1,329.01  -> quitar comas de miles
    if "," in tok and "." in tok:
        tok = tok.replace(",", "")
    # Caso 2: 1329,01 -> coma decimal
    elif "," in tok and "." not in tok:
        tok = tok.replace(",", ".")

    try:
        return float(tok)
    except ValueError:
        return None


def find_total_invoice(text: str) -> Optional[float]:
    """Encuentra Total Invoice (varios formatos)."""
    m = re.search(r"TOTAL\s+INVOICE\s*[: ]\s*([0-9][0-9.,]+)", text, flags=re.IGNORECASE)
    if m:
        return parse_amount(m.group(1))

    m = re.search(r"Total\s+Invoices?.*?([0-9][0-9.,]+)", text, flags=re.IGNORECASE | re.DOTALL)
    if m:
        return parse_amount(m.group(1))

    return None


def extract_one(pdf_path: Path) -> InvoiceExtract:
    text = pdf_text_no_ocr(pdf_path)
    stem = pdf_path.stem

    invoice_no = find_invoice_no(text, filename_stem=stem)

    pallets = find_int_after_label(text, "Pallets")
    boxes = find_int_after_label(text, "Boxes")

    gross_weight = find_float_after_label(text, "Gross Weight")

    # Taxable Amount: en estos PDFs suele aparecer justo antes de 'Vat %'
    taxable_amount = find_float_before_marker(text, r"Vat\s*%")

    total_invoice = find_total_invoice(text)


    return InvoiceExtract(
        file=pdf_path.name,
        invoice_no=invoice_no,
        pallets=pallets,
        boxes=boxes,
        gross_weight=gross_weight,
        taxable_amount=taxable_amount,
        total_invoice=total_invoice,
    )


def iter_pdfs(inputs: list[Path]) -> Iterable[Path]:
    for p in inputs:
        if p.is_dir():
            yield from sorted(p.glob("*.pdf"))
        else:
            yield p


def sum_optional_decimal(values: Iterable[Optional[float]]) -> Decimal:
    total = Decimal("0")
    for v in values:
        if isinstance(v, (int, float)):
            total += Decimal(str(v))  # evita artefactos binarios típicos de float
    return total



def sum_optional_int(values: Iterable[Optional[int]]) -> int:
    total = 0
    for v in values:
        if isinstance(v, int):
            total += v
    return total
def write_csv(csv_path: Path, rows: list[InvoiceExtract], total_pallets: int, total_boxes: int, total_gross_weight_r: Decimal, total_taxable_amount_r: Decimal, total_total_invoice_r: Decimal) -> None:
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        for r in rows:
            w.writerow({k: asdict(r).get(k) for k in FIELDNAMES})

        # TOTAL row (solo suma de gross_weight y total_invoice, como venías usando)
        w.writerow(
            {
                "file": "TOTAL",
                "invoice_no": None,
                "pallets": total_pallets,
                "boxes": total_boxes,
                "gross_weight": f"{total_gross_weight_r:.4f}",
                "taxable_amount": f"{total_taxable_amount_r:.2f}",
                "total_invoice": f"{total_total_invoice_r:.2f}",
            }
        )


def write_xlsx(xlsx_path: Path, rows: list[InvoiceExtract], total_pallets: int, total_boxes: int, total_gross_weight_r: Decimal, total_taxable_amount_r: Decimal, total_total_invoice_r: Decimal) -> None:
    if Workbook is None:
        raise RuntimeError("Falta dependencia para XLSX. Instala con: pip install openpyxl")

    wb = Workbook()
    ws = wb.active
    ws.title = "invoices"

    # Header
    ws.append(FIELDNAMES)

    # Rows
    for r in rows:
        d = asdict(r)
        ws.append([d.get(k) for k in FIELDNAMES])

    # TOTAL row
    ws.append(
        [
            "TOTAL",
            None,
            total_pallets,
            total_boxes,
            float(total_gross_weight_r),  # numérico en Excel
            float(total_taxable_amount_r),
            float(total_total_invoice_r),
        ]
    )

    wb.save(xlsx_path)


def main() -> int:
    ap = argparse.ArgumentParser(description=SCRIPT_DESCRIPTION)
    ap.add_argument("inputs", nargs="+", type=Path, help="PDF(s) o carpeta(s) con PDFs")
    ap.add_argument("-o", "--out", type=Path, default=Path("out"), help="Carpeta de salida (puede ser absoluta)")
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
            print(f"[WARN] Falló {pdf.name}: {e}", file=sys.stderr)

    total_pallets = sum_optional_int(r.pallets for r in rows)
    total_boxes = sum_optional_int(r.boxes for r in rows)
    total_gross_weight = sum_optional_decimal(r.gross_weight for r in rows)
    total_taxable_amount = sum_optional_decimal(r.taxable_amount for r in rows)
    total_total_invoice = sum_optional_decimal(r.total_invoice for r in rows)

    total_gross_weight_r = total_gross_weight.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
    total_taxable_amount_r = total_taxable_amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    total_total_invoice_r = total_total_invoice.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    # JSON detalle (fuente)
    json_path = out_dir / "invoices_extracted.json"
    json_path.write_text(
        json.dumps([asdict(r) for r in rows], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # CSV + XLSX (solo campos FIELDNAMES)
    csv_path = out_dir / "invoices_extracted.csv"
    xlsx_path = out_dir / "invoices_extracted.xlsx"
    write_csv(csv_path, rows, total_pallets, total_boxes, total_gross_weight_r, total_taxable_amount_r, total_total_invoice_r)
    write_xlsx(xlsx_path, rows, total_pallets, total_boxes, total_gross_weight_r, total_taxable_amount_r, total_total_invoice_r)

    # Console output (mantiene tu estilo de OK + SUM)
    print(f"[INFO] extract_lear_fields.py version: {SCRIPT_VERSION}")
    print(f"OK → {json_path}")
    print(f"OK → {csv_path}")
    print(f"OK → {xlsx_path}")
    print(f"SUM pallets: {total_pallets}")
    print(f"SUM boxes: {total_boxes}")
    print(f"SUM gross_weight: {total_gross_weight_r:.4f}")
    print(f"SUM taxable_amount: {total_taxable_amount_r:.2f}")
    print(f"SUM total_invoice: {total_total_invoice_r:.2f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
