#!/usr/bin/env python3
"""
Croton Packing List Processor
==============================
Inputs : ODS packing-list (Hoja5 format — one row per roll/bulto)
         XLSX product-mapping (columns: CODIGO, DESCRIPCION_CONTAINS, MERCANCIA, PARTIDA)
Output : The same ODS with a new tab "Resumen_Partidas" containing the
         aggregated customs-tariff summary.

Transformation steps:
  Hoja5  →  group rolls by consecutive CANT_TOTAL markers
         →  map (CODIGO, DESCRIPCION) to (MERCANCIA, PARTIDA)  [mapping XLSX]
         →  aggregate by (MERCANCIA, PARTIDA)  →  new tab in original ODS
"""

import os
import sys
import csv
import logging
import argparse
import shutil
from pathlib import Path
from collections import OrderedDict

from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TextProperties, TableCellProperties
from odf.namespaces import OFFICENS

# ─────────────────────────────────────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
log = logging.getLogger(__name__)

EXTRACTOR_DIR = Path(__file__).parent
# Fallback CSV mapping (used when no XLSX mapping is provided)
DEFAULT_MAPPING_CSV = EXTRACTOR_DIR / "product_mapping.csv"

# ─────────────────────────────────────────────────────────────────────────────
# Helpers – read ODS cells/rows
# ─────────────────────────────────────────────────────────────────────────────

def _cell_text(cell) -> str:
    """Return the plain-text content of an ODS cell."""
    result = []
    for p in cell.getElementsByType(P):
        for node in p.childNodes:
            if hasattr(node, "data"):
                result.append(node.data)
    return "".join(result).strip()


def _sheet_rows(sheet) -> list:
    """Return all rows of a sheet as lists of strings (trailing blanks trimmed)."""
    rows = []
    for row in sheet.getElementsByType(TableRow):
        cells = []
        for cell in row.getElementsByType(TableCell):
            repeat = cell.getAttribute("numbercolumnsrepeated")
            val = _cell_text(cell)
            count = int(repeat) if (repeat and int(repeat) < 50) else 1
            cells.extend([val] * count)
        while cells and cells[-1] == "":
            cells.pop()
        rows.append(cells)
    return rows


def _parse_number(s: str) -> float:
    """Parse a Spanish-formatted number (comma decimal) or return 0."""
    if not s:
        return 0.0
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return 0.0


def _fmt(value: float) -> str:
    """Format a float back to Spanish style (comma decimal), strip trailing zeros."""
    if value == 0:
        return ""
    s = f"{value:.4f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")

# ─────────────────────────────────────────────────────────────────────────────
# Step 1 – read Hoja5 and group rolls
# ─────────────────────────────────────────────────────────────────────────────
# Hoja5 columns (0-based, after the header rows):
#   0=LINEA  1=NBULTO  2=DESCRIPCION  3=CODIGO  4=NPALET
#   5=PESO_NETO  6=PESO_BRUTO  7=CANT  8=CANT_TOTAL  9=M2  10=VALOR (optional)

HOJA5_SHEET_NAME = "Hoja5"
_HEADER_ROWS = 8  # data starts at row 9 (0-based index 8)


def read_hoja5(ods_path: str) -> list:
    """
    Read the input ODS and return a list of 'group' dicts, each representing
    one aggregated shipment unit:

      {
        'codigo':      str,
        'descripcion': str,   # from first row in group
        'bx':          int,   # number of individual rolls in group
        'neto':        float, # sum of PESO_NETO
        'bruto':       float, # sum of PESO_BRUTO
        'cant_total':  float, # from last row
        'm2':          float, # from last row
        'valor':       float, # declared customs value (from last row)
      }

    A 'group' is a run of consecutive data rows that ends when
    CANT_TOTAL (col 8) is non-empty.
    """
    doc = load(ods_path)
    sheets = doc.spreadsheet.getElementsByType(Table)

    hoja5 = None
    for s in sheets:
        if s.getAttribute("name") == HOJA5_SHEET_NAME:
            hoja5 = s
            break
    if hoja5 is None:
        hoja5 = sheets[0]
        log.warning("Sheet '%s' not found; using first sheet.", HOJA5_SHEET_NAME)

    all_rows = _sheet_rows(hoja5)

    groups = []
    current_bx = 0
    current_neto = 0.0
    current_bruto = 0.0
    first_codigo = ""
    first_desc = ""
    in_group = False

    for row in all_rows[_HEADER_ROWS:]:
        if not row:
            continue

        linea = row[0] if len(row) > 0 else ""
        if not linea or not linea.strip().lstrip("-").isdigit():
            continue

        while len(row) < 10:
            row.append("")

        descripcion = row[2]
        codigo = row[3]
        peso_neto = _parse_number(row[5])
        peso_bruto = _parse_number(row[6])
        cant_total = row[8].strip()
        m2 = row[9].strip()
        valor = row[10].strip() if len(row) > 10 else ""

        if not in_group:
            first_codigo = codigo
            first_desc = descripcion
            in_group = True

        current_bx += 1
        current_neto += peso_neto
        current_bruto += peso_bruto

        if cant_total:
            groups.append({
                "codigo":      first_codigo,
                "descripcion": first_desc,
                "bx":          current_bx,
                "neto":        round(current_neto, 2),
                "bruto":       round(current_bruto, 2),
                "cant_total":  _parse_number(cant_total),
                "m2":          _parse_number(m2),
                "valor":       _parse_number(valor),
            })
            current_bx = 0
            current_neto = 0.0
            current_bruto = 0.0
            first_codigo = ""
            first_desc = ""
            in_group = False

    if in_group and current_bx > 0:
        log.warning(
            "Last group (%s) has no CANT_TOTAL marker — including with partial data.",
            first_codigo,
        )
        groups.append({
            "codigo":      first_codigo,
            "descripcion": first_desc,
            "bx":          current_bx,
            "neto":        round(current_neto, 2),
            "bruto":       round(current_bruto, 2),
            "cant_total":  0.0,
            "m2":          0.0,
            "valor":       0.0,
        })

    log.info("Read %d groups from %s.", len(groups), ods_path)
    return groups

# ─────────────────────────────────────────────────────────────────────────────
# Step 2 – load product mapping (XLSX or CSV)
# ─────────────────────────────────────────────────────────────────────────────

def load_mapping_xlsx(xlsx_path: str) -> list:
    """Load product mapping from an XLSX file.

    Expected columns: CODIGO, DESCRIPCION_CONTAINS, MERCANCIA, PARTIDA
    Rules are evaluated in file order; first match wins.
    """
    try:
        from openpyxl import load_workbook
    except ImportError as e:
        raise ImportError(
            "openpyxl is required to read XLSX mapping files. "
            "Install it with: pip install openpyxl"
        ) from e

    wb = load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active

    rules = []
    header = None
    for row in ws.iter_rows(values_only=True):
        if header is None:
            header = [str(c).strip() if c is not None else "" for c in row]
            continue
        if not any(row):
            continue
        row_dict = dict(zip(header, row))
        codigo = str(row_dict.get("CODIGO", "") or "").strip()
        if not codigo or codigo.startswith("#"):
            continue
        rules.append({
            "codigo":    codigo.upper(),
            "contains":  str(row_dict.get("DESCRIPCION_CONTAINS", "") or "").strip().upper(),
            "mercancia": str(row_dict.get("MERCANCIA", "") or "").strip(),
            "partida":   str(row_dict.get("PARTIDA", "") or "").strip(),
        })

    wb.close()
    log.info("Loaded %d mapping rules from %s.", len(rules), xlsx_path)
    return rules


def load_mapping_csv(csv_path: str) -> list:
    """Load product mapping from a CSV file (fallback).

    Expected columns: CODIGO, DESCRIPCION_CONTAINS, MERCANCIA, PARTIDA
    """
    rules = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            codigo = row.get("CODIGO", "").strip()
            if not codigo or codigo.startswith("#"):
                continue
            rules.append({
                "codigo":    codigo.upper(),
                "contains":  row.get("DESCRIPCION_CONTAINS", "").strip().upper(),
                "mercancia": row.get("MERCANCIA", "").strip(),
                "partida":   row.get("PARTIDA", "").strip(),
            })
    log.info("Loaded %d mapping rules from %s.", len(rules), csv_path)
    return rules


def load_mapping(path: str) -> list:
    """Load product mapping from XLSX or CSV depending on file extension."""
    p = Path(path)
    if p.suffix.lower() in (".xlsx", ".xls"):
        return load_mapping_xlsx(str(p))
    return load_mapping_csv(str(p))


def lookup(codigo: str, descripcion: str, rules: list) -> tuple:
    """Return (MERCANCIA, PARTIDA) for a group. Raises ValueError if not found."""
    cod_up = codigo.strip().upper()
    desc_up = descripcion.strip().upper()
    for rule in rules:
        if rule["codigo"] != cod_up:
            continue
        if rule["contains"] and rule["contains"] not in desc_up:
            continue
        return rule["mercancia"], rule["partida"]
    raise ValueError(
        f"No mapping found for CODIGO='{codigo}' DESCRIPCION='{descripcion}'"
    )

# ─────────────────────────────────────────────────────────────────────────────
# Step 3 – aggregate by (MERCANCIA, PARTIDA)
# ─────────────────────────────────────────────────────────────────────────────

def aggregate(groups: list, rules: list) -> list:
    """Map each group to (MERCANCIA, PARTIDA) and sum BX, NETO, BRUTO, M2."""
    seen: OrderedDict = OrderedDict()
    errors = []

    for g in groups:
        try:
            mercancia, partida = lookup(g["codigo"], g["descripcion"], rules)
        except ValueError as e:
            errors.append(str(e))
            log.error(str(e))
            continue

        key = (mercancia, partida)
        if key not in seen:
            seen[key] = {
                "mercancia": mercancia,
                "partida":   partida,
                "bx":        0,
                "valor":     0.0,
                "bruto":     0.0,
                "neto":      0.0,
                "m2":        0.0,
            }
        seen[key]["bx"]    += g["bx"]
        seen[key]["bruto"] += g["bruto"]
        seen[key]["neto"]  += g["neto"]
        seen[key]["m2"]    += g["m2"]
        seen[key]["valor"] += g["valor"]

    if errors:
        unmapped = "\n  ".join(errors)
        raise RuntimeError(
            f"Could not map {len(errors)} group(s) — update product mapping:\n  {unmapped}"
        )

    result = list(seen.values())
    for r in result:
        r["bruto"] = round(r["bruto"], 2)
        r["neto"]  = round(r["neto"],  2)
        r["m2"]    = round(r["m2"],    2)
        r["valor"] = round(r["valor"], 2)

    log.info("Aggregated into %d MERCANCIA/PARTIDA rows.", len(result))
    return result

# ─────────────────────────────────────────────────────────────────────────────
# Step 4 – add new tab to original ODS
# ─────────────────────────────────────────────────────────────────────────────

TAB_NAME = "Resumen_Partidas"
HEADER_BG = "#C0C0C0"


def _make_cell(doc, value: str, bold: bool = False, bg: str = None) -> TableCell:
    """Create a styled ODS table cell with a string value."""
    style_name = None

    if bold or bg:
        style = Style(name=f"ce_{abs(hash(value + str(bold) + str(bg)))}", family="table-cell")
        if bold:
            style.addElement(TextProperties(fontweight="bold"))
        if bg:
            style.addElement(TableCellProperties(backgroundcolor=bg))
        doc.automaticstyles.addElement(style)
        style_name = style.getAttribute("name")

    cell = TableCell()
    if style_name:
        cell.setAttribute("stylename", style_name)

    p = P(text=value)
    cell.addElement(p)
    return cell


def _row_from_values(doc, values: list, bold: bool = False, bg: str = None) -> TableRow:
    row = TableRow()
    for v in values:
        row.addElement(_make_cell(doc, str(v) if v is not None else "", bold=bold, bg=bg))
    return row


def add_summary_tab(ods_path: str, summary: list, output_path: str) -> None:
    """Load the original ODS, add/replace the summary tab, and save to output_path."""
    doc = load(ods_path)

    # Remove existing tab with same name to avoid duplicates
    for sheet in list(doc.spreadsheet.getElementsByType(Table)):
        if sheet.getAttribute("name") == TAB_NAME:
            doc.spreadsheet.removeChild(sheet)
            log.info("Removed existing '%s' tab.", TAB_NAME)

    sheet = Table(name=TAB_NAME)
    doc.spreadsheet.addElement(sheet)

    headers = [
        "MERCANCIA", "PARTIDA",
        "Suma - BX", "Suma - VALOR",
        "Suma - BRUTO", "Suma - NETO", "Suma - M2",
    ]
    sheet.addElement(_row_from_values(doc, headers, bold=True, bg=HEADER_BG))

    for r in summary:
        def fmt(v):
            if isinstance(v, float) and v != 0:
                return _fmt(v)
            return str(v) if v else ""

        row_vals = [
            r["mercancia"],
            r["partida"],
            str(r["bx"]),
            fmt(r["valor"]) if r["valor"] else "",
            fmt(r["bruto"]) if r["bruto"] else "",
            fmt(r["neto"])  if r["neto"]  else "",
            fmt(r["m2"])    if r["m2"]    else "",
        ]
        sheet.addElement(_row_from_values(doc, row_vals))

    doc.save(output_path)
    log.info("Output written to %s", output_path)

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Croton packing-list processor: ODS + XLSX mapping → ODS with new tab"
    )
    parser.add_argument(
        "inputs",
        nargs="+",
        metavar="FILE",
        help=(
            "Input files: one ODS packing-list and one XLSX product-mapping. "
            "Order does not matter; files are identified by extension."
        ),
    )
    parser.add_argument(
        "-o", "--output",
        dest="output",
        required=True,
        help="Output directory (or full output file path)",
    )
    parser.add_argument(
        "--mapping",
        dest="mapping_override",
        default=None,
        help="Override path to product mapping file (XLSX or CSV)",
    )
    args = parser.parse_args()

    # Identify ODS and XLSX from the positional inputs
    ods_path = None
    xlsx_path = None
    for f in args.inputs:
        ext = Path(f).suffix.lower()
        if ext == ".ods":
            if ods_path is not None:
                log.warning("Multiple ODS files provided; using first: %s", ods_path)
            else:
                ods_path = f
        elif ext in (".xlsx", ".xls"):
            if xlsx_path is not None:
                log.warning("Multiple XLSX files provided; using first: %s", xlsx_path)
            else:
                xlsx_path = f

    if ods_path is None:
        parser.error("No ODS file provided.")

    if not os.path.exists(ods_path):
        log.error("ODS file not found: %s", ods_path)
        sys.exit(1)

    # Determine mapping source
    if args.mapping_override:
        mapping_path = args.mapping_override
    elif xlsx_path:
        mapping_path = xlsx_path
    elif DEFAULT_MAPPING_CSV.exists():
        mapping_path = str(DEFAULT_MAPPING_CSV)
        log.info("No XLSX mapping provided; using default CSV: %s", mapping_path)
    else:
        parser.error(
            "No XLSX mapping file provided and no default product_mapping.csv found. "
            "Supply an XLSX file alongside the ODS."
        )

    if not os.path.exists(mapping_path):
        log.error("Mapping file not found: %s", mapping_path)
        sys.exit(1)

    # Determine output path
    output = args.output
    if os.path.isdir(output):
        ods_stem = Path(ods_path).stem
        output_file = os.path.join(output, f"CROTON_{ods_stem}.ods")
    else:
        output_file = output

    # Run pipeline
    groups  = read_hoja5(ods_path)
    rules   = load_mapping(mapping_path)
    summary = aggregate(groups, rules)
    add_summary_tab(ods_path, summary, output_file)

    print(
        f"✅ Done — {len(summary)} partidas aduaneras written to new tab "
        f"'{TAB_NAME}' in {output_file}"
    )


if __name__ == "__main__":
    main()
