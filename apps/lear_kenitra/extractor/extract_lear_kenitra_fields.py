#!/usr/bin/env python3
"""Extractor de campos para facturas Lear Kenitra (DSxxxxxxx).

Genera un XLSX con la MISMA estructura de columnas que los extractores
de Lear Cable y Lear TAC, para que las tres salidas puedan consolidarse
sin transformaciones adicionales.
"""
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

SCRIPT_VERSION = "2026-05-24.v2"

SCRIPT_CHANGELOG = """
## 2026-05-24.v2

### Fix: extraccion vacia en entornos con pypdf distinto
Reportado: pallets/boxes/weights/total salian vacios aunque invoice_no
y date funcionaban. Causa probable: el modo `layout` de pypdf renderiza
los PDFs Kenitra de forma muy distinta entre versiones; en algunas, los
labels y valores quedan en lineas separadas con tanta distancia que las
regex con `\\s*` (que solo cruzan whitespace contiguo) no los emparejan.

Cambios:
- `pdf_text_combined()`: extrae el texto en AMBOS modos (default y layout)
  y los concatena. Asi siempre hay al menos una representacion donde el
  valor sigue de cerca a su label.
- Helpers `_find_int_after_label` y `_find_weight_after_label`:
  ventana permisiva `.{0,300}?` con `re.DOTALL` entre el label y el
  valor, para tolerar que pypdf inserte otras celdas o saltos de linea
  entre medias. La busqueda es perezosa, asi que coge el primer entero
  /decimal valido.
- Ancla `(?<!\\d)` en los helpers para no arrancar a mitad de un numero
  anterior (defensa contra falsos positivos tipo '65/19/2026').

## 2026-05-24.v1

### Logica general
Extrae campos de facturas Lear Kenitra (formato emitido desde Kenitra,
Morocco) y genera un XLSX con la misma estructura que los extractores
Lear Cable y Lear TAC.

### Particularidades del formato Kenitra
- El PDF es bilingue (ingles/frances) y pypdf renderiza cada label y cada
  valor REPETIDOS hasta 4 veces (es como dibuja el PDF las celdas
  alineadas en ambos idiomas, mas un fantasma del template).
  Ejemplo: "Net Weight:Net Weight:3230.98323230.98323230.98323230.9832"
- Las regex toleran labels repetidos via `(?:LABEL\\s*:\\s*)*` y los
  valores numericos repetidos se reducen con `reduce_repetition`.
- Fecha en formato US (MM/DD/YYYY) -> se normaliza a DD/MM/YYYY.
- Numero de factura tipo "DS307499" se extrae del campo "Invoice Number"
  (no del nombre del fichero) y se fuerza a mayusculas.

### Extraccion PDF
- Usa pypdf con extraction_mode="layout". En modo texto plano los valores
  aparecen separados de sus labels y la asociacion se rompe.

### Salida
Genera `invoices_extracted.xlsx` con la estructura comun:
file, invoice_no, date, pallets, boxes, gross_weight, net_weight, total_invoice.
"""

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
    """Reduce un token que sea N copias exactas de una unidad.
    'DS307499DS307499' -> 'DS307499'
    '3230.98323230.9832' -> '3230.9832'
    '27272727' -> '27'
    """
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
    """Parser numerico robusto: tolera tokens repetidos y separadores
    EU/US."""
    if token is None:
        return None

    raw = str(token)
    parts = re.findall(r"[0-9][0-9.,]*", raw)
    if parts:
        raw = parts[0] if all(p == parts[0] for p in parts) else max(parts, key=len)

    tok = reduce_repetition(raw).strip()
    if not tok:
        return None

    tok = tok.replace("\u00A0", "").replace(" ", "")
    tok = re.sub(r"(?i)(kg|kgs|g)$", "", tok).strip()

    if "," in tok and "." in tok:
        if tok.rfind(",") > tok.rfind("."):
            tok = tok.replace(".", "").replace(",", ".")
        else:
            tok = tok.replace(",", "")
    elif "," in tok and "." not in tok:
        # En Kenitra los decimales usan punto, asi que la coma deberia ser
        # separador de miles. Aun asi, si hay <=2 digitos despues, tratamos
        # como decimal para no romper formatos mixtos.
        parts = tok.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            tok = tok.replace(",", ".")
        else:
            tok = tok.replace(",", "")

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
# PDF text (layout mode)
# ----------------------------

def _normalize_extracted_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\u202f", " ")
    text = re.sub(r"[\u200B\u200C\u200D\u200E\u200F\u2060\uFEFF]", "", text)
    text = re.sub(r"[\u202A-\u202E\u2066-\u2069]", "", text)
    return text


def pdf_text_combined(pdf_path: Path) -> str:
    """Extrae el texto en DOS modos y los concatena.

    Razon: pypdf cambia el layout entre versiones para PDFs como Kenitra
    que tienen labels y valores en celdas alineadas verticalmente. En
    pypdf 5.x el modo `layout` agrupa label y valor bien, pero en otras
    versiones (o segun el PDF) los separa en lineas distintas y los
    valores aparecen lejos de sus labels. Concatenar ambas extracciones
    asegura que SIEMPRE haya al menos UNA representacion donde label y
    valor estan suficientemente cerca para que las regex los emparejen.
    """
    reader = PdfReader(str(pdf_path))
    parts = []
    for p in reader.pages:
        try:
            parts.append(p.extract_text() or "")
        except Exception:
            pass
        try:
            parts.append(p.extract_text(extraction_mode="layout") or "")
        except TypeError:
            # pypdf antiguo sin extraction_mode -> ya tenemos default
            pass
        except Exception:
            pass
    return _normalize_extracted_text("\n".join(parts))


# ----------------------------
# Fechas (formato US: MM/DD/YYYY)
# ----------------------------

def _expand_year(y: int) -> int:
    if y < 100:
        return 2000 + y if y <= 30 else 1900 + y
    return y


def _normalize_date_us(raw: str) -> Optional[str]:
    """Convierte 'M/D/YYYY' (US) a 'DD/MM/YYYY'."""
    raw = raw.strip()
    m = re.fullmatch(r"(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", raw)
    if not m:
        return raw
    a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
    yyyy = _expand_year(c)
    if a > 12:
        dd, mm = a, b
    elif b > 12:
        dd, mm = b, a
    else:
        dd, mm = b, a  # default US -> swap
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return raw
    return f"{dd:02d}/{mm:02d}/{yyyy}"


def find_date(text: str) -> Optional[str]:
    """Captura la primera fecha con anyo de 4 digitos.

    En Kenitra la cabecera aparece como '5/19/20265/19/20265/19/20265/19/2026'
    (4 copias pegadas sin separador). Capturamos la secuencia completa y la
    reducimos con `reduce_repetition`. Las lineas de detalle usan formato
    '26/05/19' (anyo 2 digitos) y se ignoran gracias al ancla \\d{4}.

    Los anclajes (?<!\\d) y (?!\\d) impiden que el motor empiece a mitad de
    un anyo (p.ej. atrapar '65/19/2026' arrancando en el '6' de un '2026'
    previo).
    """
    m = re.search(r"(?<!\d)((?:\d{1,2}/\d{1,2}/\d{4})+)(?!\d)", text)
    if m:
        raw = reduce_repetition(m.group(1))
        return _normalize_date_us(raw)
    return None


# ----------------------------
# Finders especificos
# ----------------------------

def find_invoice_no(text: str, filename_stem: Optional[str] = None) -> Optional[str]:
    """Captura el numero de factura.

    En layout mode el texto se ve como:
      'Invoice Number:Invoice Number:Invoice Number:Invoice Number:   DS307499DS307499Customer Code:'
    Estrategia: saltarse las repeticiones del label y capturar
    LETRAS+DIGITOS (que para 'DS307499' va a parar exactamente).
    """
    m = re.search(
        r"Invoice\s+Number\s*:\s*(?:Invoice\s+Number\s*:\s*)*([A-Z]+\d+)",
        text, flags=re.IGNORECASE,
    )
    if m:
        return m.group(1).upper()
    if filename_stem:
        stem = re.sub(r"[^A-Za-z0-9_-]+", "", filename_stem.strip())
        return stem.upper() or None
    return None


def _find_int_after_label(text: str, label_regex: str) -> Optional[int]:
    """Busca un entero PURO (sin punto decimal) tras `label_regex`.

    Permite hasta ~300 chars de basura entre el label y el valor: pypdf
    puede insertar otras celdas (Boxes, Quantity, Devise...) o renderizar
    label y valor en lineas distintas con bastante separacion. La busqueda
    es perezosa para coger el PRIMER entero valido.

    La condicion `(?![\\d.])` evita capturar la parte entera de un decimal
    como '5255' de '5255.98329' (un valor de Gross Weight que a veces se
    cuela entre el label Pallets: y su valor real).
    """
    m = re.search(
        rf"{label_regex}.{{0,300}}?(?<!\d)([0-9]+)(?![\d.])",
        text, flags=re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return None
    return int(reduce_repetition(m.group(1)))


def find_pallets(text: str) -> Optional[int]:
    return _find_int_after_label(text, r"Pallets\s*:")


def find_boxes(text: str) -> Optional[int]:
    return _find_int_after_label(text, r"Boxes\s*:")


def _find_weight_after_label(text: str, label_regex: str) -> Optional[float]:
    """Busca un numero decimal/entero tras `label_regex`.

    Igual que para enteros, permite hasta ~300 chars de basura entre label
    y valor. El numero puede aparecer concatenado N veces; lo reduce
    `parse_number_token` via `reduce_repetition`. El ancla `(?<!\\d)` evita
    empezar a mitad de otro numero anterior.
    """
    m = re.search(
        rf"{label_regex}.{{0,300}}?(?<!\d)([\d.]+)",
        text, flags=re.IGNORECASE | re.DOTALL,
    )
    if not m:
        return None
    return parse_number_token(m.group(1))


def find_gross_weight(text: str) -> Optional[float]:
    return _find_weight_after_label(text, r"Gross\s+Weight\s*:")


def find_net_weight(text: str) -> Optional[float]:
    return _find_weight_after_label(text, r"Net\s+Weight\s*:")


def find_total_invoice(text: str) -> Optional[float]:
    """Captura el total. En Kenitra:
      'Total InvoicesTotal Invoices                90307.2890307.28...'
      seguido en otra linea por 'TOTAL facrure' (sin valor adyacente).
    """
    # Patron 1: 'Total Invoices ... 90307.28'
    m = re.search(
        r"Total\s+Invoices?\b[^\d]*?([\d.,]+)",
        text, flags=re.IGNORECASE | re.DOTALL,
    )
    if m:
        v = parse_number_token(m.group(1))
        if v is not None:
            return v
    # Patron 2: 'XXX TOTAL facrure' (label tras el valor)
    m = re.search(
        r"([\d.,]+)\s*TOTAL\s+facrure",
        text, flags=re.IGNORECASE | re.DOTALL,
    )
    if m:
        v = parse_number_token(m.group(1))
        if v is not None:
            return v
    return None


# ----------------------------
# Extraccion por factura
# ----------------------------

def extract_one(pdf_path: Path) -> InvoiceExtract:
    text = pdf_text_combined(pdf_path)
    stem = pdf_path.stem

    return InvoiceExtract(
        file=pdf_path.name,
        invoice_no=find_invoice_no(text, filename_stem=stem),
        date=find_date(text),
        pallets=find_pallets(text),
        boxes=find_boxes(text),
        gross_weight=find_gross_weight(text),
        net_weight=find_net_weight(text),
        total_invoice=find_total_invoice(text),
    )


def iter_pdfs(inputs: list[Path]) -> Iterable[Path]:
    for p in inputs:
        if p.is_dir():
            all_pdfs = list(p.glob("*.pdf")) + list(p.glob("*.PDF"))
            yield from sorted(all_pdfs, key=lambda f: f.name.lower())
        else:
            yield p


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Extrae campos clave de facturas Lear Kenitra (PDF con texto real, sin OCR)."
    )
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

    rows.sort(key=lambda r: r.file.lower())

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

    print(f"[INFO] Processing order ({len(rows)} files):")
    for _i, _r in enumerate(rows, 1):
        print(f"  {_i:>3}. {_r.file}")

    # CSV
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