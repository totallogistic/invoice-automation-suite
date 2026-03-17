"""Camion Export Processor.

Procesa el packing list de exportación de camiones integrando los datos
extraídos de los documentos T1 (declaraciones de tránsito) y el documento
de acompañamiento (DOC).

CLI de terminal (modo standalone):
    python camion_export_processor.py \\
        --xlsx  20260310_Kenitra_6586-08_PL.xlsx \\
        --t1    T1_1.pdf T1_2.pdf \\
        --doc   20260310_Kenitra_6586-08_DOC.pdf \\
        --output 20260310_Kenitra_6586-08_PL-PROCESSED.xlsx

CLI de integración web (modo stack, --output no es obligatorio):
    python camion_export_processor.py \\
        --xlsx  packing.xlsx \\
        --t1    T1_1.pdf T1_2.pdf \\
        --doc   doc.pdf \\
        -o      /data/camion/out/batch_id
"""
from __future__ import annotations

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional

SCRIPT_VERSION = "2026-03-17.v1"

SCRIPT_CHANGELOG = """
v1  2026-03-17  Versión inicial integrada en el stack web.
"""

# ── Imports opcionales ──────────────────────────────────────────────────────
try:
    import pdfplumber
    _PDFPLUMBER = True
except ImportError:
    _PDFPLUMBER = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    _OPENPYXL = True
except ImportError:
    _OPENPYXL = False


# ── Extracción de texto PDF ─────────────────────────────────────────────────

def extract_pdf_text(pdf_path: Path) -> str:
    """Extrae todo el texto de un PDF."""
    if not _PDFPLUMBER:
        print(f"  [WARN] pdfplumber no disponible; saltando extracción de {pdf_path.name}")
        return ""
    text_parts: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            text_parts.append(t)
    return "\n".join(text_parts)


# ── Parsers específicos ──────────────────────────────────────────────────────

def parse_t1(text: str, filename: str) -> dict:
    """Extrae campos clave de un documento T1."""
    data: dict = {"filename": filename}

    # MRN (Movement Reference Number) — formato estándar EU: 2 dígitos año + 2 letras país + …
    mrn_match = re.search(r"\b(\d{2}[A-Z]{2}[A-Z0-9]{14,16})\b", text)
    if mrn_match:
        data["mrn"] = mrn_match.group(1)

    # Fecha del T1 (DD/MM/YYYY o YYYY-MM-DD)
    date_match = re.search(
        r"\b(\d{2}[/-]\d{2}[/-]\d{4}|\d{4}[/-]\d{2}[/-]\d{2})\b", text
    )
    if date_match:
        data["fecha"] = date_match.group(1)

    # Peso bruto total (p.ej. "1 234,56 kg" o "1234.56 KG")
    peso_match = re.search(
        r"(?:peso\s+bruto|gross\s+weight)[^\d]*(\d[\d\s.,]+)\s*kg",
        text,
        re.IGNORECASE,
    )
    if peso_match:
        raw = peso_match.group(1).replace(" ", "").replace(",", ".")
        try:
            data["peso_bruto_kg"] = float(raw)
        except ValueError:
            pass

    # Número de bultos / packages
    bultos_match = re.search(
        r"(?:n[uú]mero\s+de\s+bultos|packages?|nb\.\s*colis)[^\d]*(\d+)",
        text,
        re.IGNORECASE,
    )
    if bultos_match:
        data["bultos"] = int(bultos_match.group(1))

    return data


def parse_doc(text: str, filename: str) -> dict:
    """Extrae campos clave del documento de acompañamiento (DOC)."""
    data: dict = {"filename": filename}

    # Número de referencia / albarán
    ref_match = re.search(
        r"(?:ref(?:erencia)?|n[uú]m(?:ero)?\.?\s*(?:de\s+)?(?:expedici[oó]n|albar[aá]n|doc(?:umento)?))[^\w]*([A-Z0-9][A-Z0-9\-_/]{3,})",
        text,
        re.IGNORECASE,
    )
    if ref_match:
        data["referencia"] = ref_match.group(1)

    # Fecha del documento
    date_match = re.search(
        r"\b(\d{2}[/-]\d{2}[/-]\d{4}|\d{4}[/-]\d{2}[/-]\d{2})\b", text
    )
    if date_match:
        data["fecha"] = date_match.group(1)

    # Peso neto y bruto
    for label in ("neto", "bruto"):
        m = re.search(
            rf"(?:peso\s+{label}|{label}\s+weight)[^\d]*(\d[\d\s.,]+)\s*kg",
            text,
            re.IGNORECASE,
        )
        if m:
            raw = m.group(1).replace(" ", "").replace(",", ".")
            try:
                data[f"peso_{label}_kg"] = float(raw)
            except ValueError:
                pass

    return data


# ── Escritura del XLSX procesado ────────────────────────────────────────────

_HDR_FILL = PatternFill("solid", fgColor="1A4D7E") if _OPENPYXL else None
_HDR_FONT = Font(bold=True, color="FFFFFF") if _OPENPYXL else None
_ALT_FILL = PatternFill("solid", fgColor="EBF0F7") if _OPENPYXL else None


def _style_header_row(ws, row_idx: int, ncols: int):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = _HDR_FILL
        cell.font = _HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def build_processed_xlsx(
    xlsx_path: Path,
    t1_data_list: List[dict],
    doc_data: dict,
    output_path: Path,
) -> Path:
    """Copia el XLSX de entrada y añade una hoja 'Procesado_Camion' con los
    datos extraídos de los PDFs T1 y DOC."""

    if not _OPENPYXL:
        raise RuntimeError("openpyxl no está instalado. Instala con: pip install openpyxl")

    # Copiar el XLSX original al destino
    shutil.copy2(str(xlsx_path), str(output_path))

    wb = openpyxl.load_workbook(str(output_path))

    # Eliminar hoja de resultado anterior si existe
    sheet_name = "Procesado_Camion"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)

    # ── Encabezado general ───────────────────────────────────────────────────
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    ws.append(["CAMION EXPORT PROCESSOR", f"Generado: {now_str}", f"v{SCRIPT_VERSION}"])
    _style_header_row(ws, ws.max_row, 3)
    ws.append([])

    # ── Sección T1 ───────────────────────────────────────────────────────────
    ws.append(["DOCUMENTOS T1"])
    _style_header_row(ws, ws.max_row, 1)

    t1_headers = ["Fichero", "MRN", "Fecha", "Peso bruto (kg)", "Bultos"]
    ws.append(t1_headers)
    _style_header_row(ws, ws.max_row, len(t1_headers))

    for i, t1 in enumerate(t1_data_list):
        row = [
            t1.get("filename", ""),
            t1.get("mrn", "—"),
            t1.get("fecha", "—"),
            t1.get("peso_bruto_kg", "—"),
            t1.get("bultos", "—"),
        ]
        ws.append(row)
        if i % 2 == 1 and _ALT_FILL:
            for col in range(1, len(row) + 1):
                ws.cell(row=ws.max_row, column=col).fill = _ALT_FILL

    ws.append([])

    # ── Sección DOC ──────────────────────────────────────────────────────────
    ws.append(["DOCUMENTO DOC"])
    _style_header_row(ws, ws.max_row, 1)

    doc_headers = ["Fichero", "Referencia", "Fecha", "Peso neto (kg)", "Peso bruto (kg)"]
    ws.append(doc_headers)
    _style_header_row(ws, ws.max_row, len(doc_headers))

    ws.append([
        doc_data.get("filename", ""),
        doc_data.get("referencia", "—"),
        doc_data.get("fecha", "—"),
        doc_data.get("peso_neto_kg", "—"),
        doc_data.get("peso_bruto_kg", "—"),
    ])

    ws.append([])

    # ── Ajuste de anchos de columna ──────────────────────────────────────────
    col_widths = [30, 22, 14, 18, 10]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    wb.save(str(output_path))
    return output_path


# ── Main ─────────────────────────────────────────────────────────────────────

def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Camion Export Processor — integra T1 y DOC en el packing list"
    )
    parser.add_argument(
        "--xlsx",
        required=True,
        metavar="PACKING_LIST.xlsx",
        help="Packing list de exportación (XLSX)",
    )
    parser.add_argument(
        "--t1",
        nargs="+",
        required=True,
        metavar="T1.pdf",
        help="Uno o más documentos T1 (PDF)",
    )
    parser.add_argument(
        "--doc",
        required=True,
        metavar="DOC.pdf",
        help="Documento de acompañamiento (PDF)",
    )

    out_group = parser.add_mutually_exclusive_group()
    out_group.add_argument(
        "--output",
        metavar="OUTPUT.xlsx",
        help="Ruta del fichero XLSX de salida (modo standalone)",
    )
    out_group.add_argument(
        "-o",
        metavar="OUTPUT_DIR",
        dest="output_dir",
        help="Directorio de salida (modo integrado en el stack web)",
    )

    parser.add_argument("--version", action="version", version=SCRIPT_VERSION)
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None):
    args = parse_args(argv)

    xlsx_path = Path(args.xlsx)
    t1_paths = [Path(p) for p in args.t1]
    doc_path = Path(args.doc)

    # ── Validar entradas ─────────────────────────────────────────────────────
    errors: List[str] = []
    if not xlsx_path.exists():
        errors.append(f"No se encuentra el packing list: {xlsx_path}")
    if not xlsx_path.suffix.lower() == ".xlsx":
        errors.append(f"El packing list debe ser .xlsx, no '{xlsx_path.suffix}'")
    for t1 in t1_paths:
        if not t1.exists():
            errors.append(f"No se encuentra T1: {t1}")
    if not doc_path.exists():
        errors.append(f"No se encuentra el DOC: {doc_path}")
    if errors:
        for e in errors:
            print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)

    # ── Determinar ruta de salida ────────────────────────────────────────────
    if args.output:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
    elif args.output_dir:
        out_dir = Path(args.output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        stem = xlsx_path.stem
        output_path = out_dir / f"{stem}-PROCESSED.xlsx"
    else:
        # Sin argumento de salida: guardar junto al XLSX de entrada
        output_path = xlsx_path.parent / f"{xlsx_path.stem}-PROCESSED.xlsx"

    print(f"Camion Export Processor  {SCRIPT_VERSION}")
    print(f"{'─' * 50}")
    print(f"Packing list : {xlsx_path.name}")
    print(f"T1 docs      : {', '.join(p.name for p in t1_paths)}")
    print(f"DOC          : {doc_path.name}")
    print(f"Salida       : {output_path}")
    print()

    # ── Extraer datos de PDFs ────────────────────────────────────────────────
    t1_data_list: List[dict] = []
    for t1_path in t1_paths:
        print(f"Procesando T1: {t1_path.name} …")
        text = extract_pdf_text(t1_path)
        t1_data = parse_t1(text, t1_path.name)
        t1_data_list.append(t1_data)
        print(f"  MRN      : {t1_data.get('mrn', '(no detectado)')}")
        print(f"  Fecha    : {t1_data.get('fecha', '(no detectado)')}")
        print(f"  Peso     : {t1_data.get('peso_bruto_kg', '(no detectado)')} kg")
        print(f"  Bultos   : {t1_data.get('bultos', '(no detectado)')}")

    print()
    print(f"Procesando DOC: {doc_path.name} …")
    doc_text = extract_pdf_text(doc_path)
    doc_data = parse_doc(doc_text, doc_path.name)
    print(f"  Referencia : {doc_data.get('referencia', '(no detectado)')}")
    print(f"  Fecha      : {doc_data.get('fecha', '(no detectado)')}")
    print(f"  Peso neto  : {doc_data.get('peso_neto_kg', '(no detectado)')} kg")
    print(f"  Peso bruto : {doc_data.get('peso_bruto_kg', '(no detectado)')} kg")
    print()

    # ── Generar XLSX procesado ───────────────────────────────────────────────
    print("Generando XLSX procesado …")
    result = build_processed_xlsx(xlsx_path, t1_data_list, doc_data, output_path)
    print(f"✓ Fichero generado: {result}")
    print()
    print("RESUMEN")
    print(f"  T1 procesados : {len(t1_data_list)}")
    mrns = [d.get("mrn") for d in t1_data_list if d.get("mrn")]
    if mrns:
        print(f"  MRNs          : {', '.join(mrns)}")
    print(f"  DOC referencia: {doc_data.get('referencia', '—')}")
    print(f"  Salida        : {result.name}")


if __name__ == "__main__":
    main()
