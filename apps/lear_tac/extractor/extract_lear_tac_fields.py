#!/usr/bin/env python3
"""Extractor de campos para facturas Lear TAC (JLR MLA Wire Harnesses).

Genera un XLSX con la MISMA estructura de columnas que el extractor de
Lear Cable (extract_lear_fields.py), de modo que ambas salidas puedan
consolidarse sin transformaciones adicionales.
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

### Fix: pesos en facturas con layout alternativo
En algunas facturas (p.ej. Invoice_30007472.pdf) pypdf en modo layout
separa el label "Total Net Weight" del valor "4608 KG" colocando otros
campos (Total Price, Total Pallets) entre ambos. La regex
"label seguido de N KG" devolvia None.

Nueva estrategia: localizar TODAS las apariciones de "N KG" y emparejar
cada una con el label mas proximo por posicion. Si ambos labels eligen
el mismo numero, el mas cercano se queda y el otro toma el siguiente.

## 2026-05-24.v1

### Logica general
Extrae campos de facturas PDF de Lear TAC (formato "JLR MLA Wire Harnesses")
y genera un XLSX con la misma estructura que el extractor de Lear Cable.

### Diferencias respecto al extractor Lear Cable
- Cabecera distinta: "Shippment No" / "Shipment Date" (formato US: MM/DD/YYYY).
- Totales agrupados en la ultima pagina:
    Total Harnesses, Total Gross Weight, Total Net Weight,
    Total Price (€), Total Pallets.
- Mapeo de columnas de salida (identico al original):
    pallets  <- Total Pallets
    boxes    <- Total Harnesses   (cada harness es un item individual;
                                   cambialo a 0 si prefieres no mezclarlos)

### Extraccion PDF
- Usa pypdf con extraction_mode="layout" para preservar la alineacion de las
  celdas de totales. En modo texto plano los valores aparecen ANTES de sus
  labels (artefacto de orden de dibujo del PDF) y son irrecuperables.

### Salida
Genera `invoices_extracted.xlsx` con la misma estructura de columnas:
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

def parse_number_token(token: str) -> Optional[float]:
    """Parser numerico robusto (miles/decimales EU/US)."""
    if token is None:
        return None
    tok = str(token).strip()
    tok = tok.replace("\u00A0", "").replace(" ", "")
    tok = re.sub(r"(?i)(kg|kgs|g|€)$", "", tok).strip()

    if "," in tok and "." in tok:
        # Mixed: el ultimo es el decimal, el otro es separador de miles.
        if tok.rfind(",") > tok.rfind("."):
            tok = tok.replace(".", "").replace(",", ".")
        else:
            tok = tok.replace(",", "")
    elif "," in tok and "." not in tok:
        # Solo coma: si hay <= 2 digitos despues, es decimal (formato EU);
        # si no, es separador de miles (improbable en este formato).
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


def pdf_text_layout(pdf_path: Path) -> str:
    """Extraccion con layout preservado.

    Imprescindible aqui porque los totales se dibujan como una unica fila
    repartida en celdas a lo largo de la pagina: en modo texto plano los
    valores aparecen DELANTE de sus labels y no hay forma fiable de
    asociarlos.
    """
    reader = PdfReader(str(pdf_path))
    raw = "\n".join(
        (p.extract_text(extraction_mode="layout") or "")
        for p in reader.pages
    )
    return _normalize_extracted_text(raw)


# ----------------------------
# Date extraction (formato US: MM/DD/YYYY)
# ----------------------------

def _expand_year(y: int) -> int:
    if y < 100:
        return 2000 + y if y <= 30 else 1900 + y
    return y


def _normalize_date_us(raw: str) -> Optional[str]:
    """Normaliza una fecha US (MM/DD/YYYY) a DD/MM/YYYY.

    Lear TAC usa formato US: "5/20/2026" = 20 mayo 2026.
    Reglas de desambiguacion:
      - A > 12  -> A es dia (ya parecia DD/MM, sin swap).
      - B > 12  -> B es dia (era MM/DD -> swap a DD/MM).
      - Ambos <= 12  -> default US, asumimos MM/DD y swap.
    """
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
        dd, mm = b, a  # US default
    if not (1 <= mm <= 12 and 1 <= dd <= 31):
        return raw
    return f"{dd:02d}/{mm:02d}/{yyyy}"


def find_date(text: str) -> Optional[str]:
    """Extrae 'Shipment Date' y lo normaliza a DD/MM/YYYY."""
    m = re.search(
        r"Shipment\s+Date\s*:?\s*(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4})",
        text, flags=re.IGNORECASE,
    )
    if m:
        return _normalize_date_us(m.group(1))
    # Fallback: primera fecha numerica del documento.
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{2,4})", text)
    if m:
        return _normalize_date_us(m.group(1))
    return None


# ----------------------------
# Finders
# ----------------------------

def find_invoice_no(text: str, filename_stem: Optional[str] = None) -> Optional[str]:
    """Devuelve el numero de envio. Prefiere 'Shippment No' (texto PDF), con
    fallback al nombre del fichero."""
    # Forma canonica del PDF (typo "Shippment" tal y como aparece en Lear TAC)
    m = re.search(r"Shipp?ment\s+No\s*:?\s*(\d+)", text, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    if filename_stem:
        stem = re.sub(r"[^A-Za-z0-9_-]+", "", filename_stem.strip())
        return stem.upper() or None
    return None


def find_total_pallets(text: str) -> Optional[int]:
    m = re.search(r"Total\s+Pallets\s+(\d+)", text, flags=re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass
    return None


def find_total_harnesses(text: str) -> Optional[int]:
    m = re.search(r"Total\s+Harnesses\s+(\d+)", text, flags=re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            pass
    return None


def find_weights(text: str) -> tuple[Optional[float], Optional[float]]:
    """Devuelve (gross_weight, net_weight).

    No se puede asumir que label y valor estan en la misma linea: pypdf en
    modo layout decide la disposicion fila a fila segun la densidad de la
    pagina, y la celda de pesos a veces sale como:

        Total Gross Weight   5088 KG
        Total Net Weight                Total Price  93,807.59 €   Total Pallets  12
                                                                       4608 KG

    Estrategia: buscar TODAS las apariciones de "N KG" y emparejar cada una
    con su label mas cercano por posicion en el texto. Si ambos labels
    eligen el mismo numero (porque uno esta sin valor en su propia linea),
    el label mas proximo se queda con el; el otro toma el siguiente mas
    cercano.
    """
    kg_matches = [
        (m.start(), parse_number_token(m.group(1)))
        for m in re.finditer(r"(\d[\d.,]*)\s*KG\b", text)
    ]
    kg_matches = [(p, v) for p, v in kg_matches if v is not None]
    if not kg_matches:
        return None, None

    gross_lbl = re.search(r"Total\s+Gross\s+Weight", text, flags=re.IGNORECASE)
    net_lbl = re.search(r"Total\s+Net\s+Weight", text, flags=re.IGNORECASE)
    if not gross_lbl and not net_lbl:
        return None, None

    # Un solo KG en el doc -> lo da al label mas cercano.
    if len(kg_matches) == 1:
        pos, val = kg_matches[0]
        gd = abs(pos - gross_lbl.start()) if gross_lbl else float("inf")
        nd = abs(pos - net_lbl.start()) if net_lbl else float("inf")
        return (val, None) if gd <= nd else (None, val)

    def closest_idx(label_pos: int) -> int:
        return min(range(len(kg_matches)), key=lambda i: abs(kg_matches[i][0] - label_pos))

    gross_idx = closest_idx(gross_lbl.start()) if gross_lbl else None
    net_idx = closest_idx(net_lbl.start()) if net_lbl else None

    # Colision: el label mas proximo se queda, el otro va al siguiente mas cercano.
    if gross_idx is not None and net_idx is not None and gross_idx == net_idx:
        gd = abs(kg_matches[gross_idx][0] - gross_lbl.start())
        nd = abs(kg_matches[net_idx][0] - net_lbl.start())
        if gd <= nd:
            candidates = [
                (i, abs(kg_matches[i][0] - net_lbl.start()))
                for i in range(len(kg_matches)) if i != gross_idx
            ]
            net_idx = min(candidates, key=lambda x: x[1])[0]
        else:
            candidates = [
                (i, abs(kg_matches[i][0] - gross_lbl.start()))
                for i in range(len(kg_matches)) if i != net_idx
            ]
            gross_idx = min(candidates, key=lambda x: x[1])[0]

    g = kg_matches[gross_idx][1] if gross_idx is not None else None
    n = kg_matches[net_idx][1] if net_idx is not None else None
    return g, n


def find_total_price(text: str) -> Optional[float]:
    """Busca 'Total Price 66,988.77 €'."""
    m = re.search(
        r"Total\s+Price\s+([0-9][0-9.,]*)\s*€",
        text, flags=re.IGNORECASE,
    )
    if m:
        return parse_number_token(m.group(1))
    # Fallback: ultima cifra con € del documento (el total suele ser la mayor).
    candidates = re.findall(r"([0-9][0-9.,]*)\s*€", text)
    if candidates:
        parsed = [v for v in (parse_number_token(c) for c in candidates) if v is not None]
        if parsed:
            return max(parsed)
    return None


# ----------------------------
# Extraccion por factura
# ----------------------------

def extract_one(pdf_path: Path) -> InvoiceExtract:
    text = pdf_text_layout(pdf_path)
    stem = pdf_path.stem

    invoice_no = find_invoice_no(text, filename_stem=stem)
    date = find_date(text)
    pallets = find_total_pallets(text)
    # MAPPING boxes <- Total Harnesses. Si prefieres dejarlo siempre en 0
    # (porque no hay concepto real de "cajas" en TAC), pon: boxes = 0
    boxes = find_total_harnesses(text)
    gross_weight, net_weight = find_weights(text)
    total_invoice = find_total_price(text)

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


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Extrae campos clave de facturas Lear TAC (PDF con texto real, sin OCR)."
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

    # Orden alfabetico estable independiente del shell glob.
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