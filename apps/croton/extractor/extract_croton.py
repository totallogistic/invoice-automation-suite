#!/usr/bin/env python3
"""
Croton Packing List Processor
==============================
Input : ODS packing-list (Hoja5 format — one row per roll/bulto)
Output: ODS summary aggregated by customs tariff (Hoja4 format)

Transformation steps replicated from the manual workflow:
  Hoja5  →  group rolls by consecutive CANT_TOTAL markers
         →  map (CODIGO, DESCRIPCION) to (MERCANCIA, PARTIDA)  [product_mapping.csv]
         →  aggregate by (MERCANCIA, PARTIDA)  →  Hoja4 output ODS
"""

import os
import sys
import csv
import logging
import argparse
from pathlib import Path
from collections import defaultdict, OrderedDict

from odf.opendocument import load, OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import (
    Style, TextProperties, TableCellProperties,
    ParagraphProperties
)
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
MAPPING_FILE  = EXTRACTOR_DIR / "product_mapping.csv"

# ─────────────────────────────────────────────────────────────────────────────
# Helpers – read ODS
# ─────────────────────────────────────────────────────────────────────────────

def _cell_text(cell) -> str:
    """Return the plain-text content of an ODS cell."""
    result = []
    for p in cell.getElementsByType(P):
        for node in p.childNodes:
            if hasattr(node, "data"):
                result.append(node.data)
    return "".join(result).strip()


def _sheet_rows(sheet) -> list[list[str]]:
    """Return all rows of a sheet as lists of strings (trailing blanks trimmed)."""
    rows = []
    for row in sheet.getElementsByType(TableRow):
        cells = []
        for cell in row.getElementsByType(TableCell):
            repeat = cell.getAttribute("numbercolumnsrepeated")
            val    = _cell_text(cell)
            count  = int(repeat) if (repeat and int(repeat) < 50) else 1
            cells.extend([val] * count)
        # trim trailing empty
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
    s = f"{value:.4f}".rstrip("0").rstrip(",").rstrip(".")
    # back to comma
    return s.replace(".", ",")

# ─────────────────────────────────────────────────────────────────────────────
# Step 1 – read Hoja5 and group rolls
# ─────────────────────────────────────────────────────────────────────────────
# Hoja5 columns (0-based, after the header rows):
#   0=LINEA  1=NBULTO  2=DESCRIPCION  3=CODIGO  4=NPALET
#   5=PESO_NETO  6=PESO_BRUTO  7=CANT  8=CANT_TOTAL  9=M2

HOJA5_SHEET_NAME = "Hoja5"

# Rows before the data area (header + sub-header rows to skip)
# Rows 1-8 are headers / page info; data starts at row 9 (0-based index 8)
_HEADER_ROWS = 8


def read_hoja5(ods_path: str) -> list[dict]:
    """
    Read the input ODS and return a list of 'group' dicts, each representing
    one aggregated shipment unit (one entry in Hoja1):

      {
        'codigo':     str,
        'descripcion': str,   # from first row in group
        'bx':         int,    # number of individual rolls in group
        'neto':       float,  # sum of PESO_NETO
        'bruto':      float,  # sum of PESO_BRUTO
        'cant_total': float,  # from last row
        'm2':         float,  # from last row
      }

    A 'group' is a run of consecutive data rows that ends when
    CANT_TOTAL (col 8) is non-empty.
    """
    doc    = load(ods_path)
    sheets = doc.spreadsheet.getElementsByType(Table)

    # Find Hoja5 (first sheet or by name)
    hoja5  = None
    for s in sheets:
        if s.getAttribute("name") == HOJA5_SHEET_NAME:
            hoja5 = s
            break
    if hoja5 is None:
        hoja5 = sheets[0]
        log.warning("Sheet '%s' not found; using first sheet.", HOJA5_SHEET_NAME)

    all_rows = _sheet_rows(hoja5)

    groups      = []
    current_bx  = 0
    current_neto  = 0.0
    current_bruto = 0.0
    first_codigo  = ""
    first_desc    = ""
    in_group      = False

    for row in all_rows[_HEADER_ROWS:]:
        if not row:
            continue

        linea = row[0] if len(row) > 0 else ""
        # Skip non-data rows (totals, palet summaries, etc.)
        if not linea or not linea.strip().lstrip("-").isdigit():
            continue

        # Pad to 10 cols
        while len(row) < 10:
            row.append("")

        nbulto      = row[1]
        descripcion = row[2]
        codigo      = row[3]
        peso_neto   = _parse_number(row[5])
        peso_bruto  = _parse_number(row[6])
        cant_total  = row[8].strip()   # non-empty → end of group
        m2          = row[9].strip()

        if not in_group:
            # Start new group
            first_codigo = codigo
            first_desc   = descripcion
            in_group     = True

        current_bx    += 1
        current_neto  += peso_neto
        current_bruto += peso_bruto

        if cant_total:
            # End of group
            groups.append({
                "codigo":      first_codigo,
                "descripcion": first_desc,
                "bx":          current_bx,
                "neto":        round(current_neto, 2),
                "bruto":       round(current_bruto, 2),
                "cant_total":  _parse_number(cant_total),
                "m2":          _parse_number(m2),
            })
            # Reset
            current_bx    = 0
            current_neto  = 0.0
            current_bruto = 0.0
            first_codigo  = ""
            first_desc    = ""
            in_group      = False

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
        })

    log.info("Read %d groups from %s.", len(groups), ods_path)
    return groups

# ─────────────────────────────────────────────────────────────────────────────
# Step 2 – load product mapping CSV
# ─────────────────────────────────────────────────────────────────────────────
# CSV columns: CODIGO, DESCRIPCION_CONTAINS, MERCANCIA, PARTIDA
#   DESCRIPCION_CONTAINS — if empty, matches any description for that CODIGO
#   Rules are evaluated in file order; first match wins.

def load_mapping(csv_path: str) -> list[dict]:
    rules = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            codigo = row.get("CODIGO", "").strip()
            # Skip comment lines (rows where CODIGO starts with #)
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


def lookup(codigo: str, descripcion: str, rules: list[dict]) -> tuple[str, str]:
    """Return (MERCANCIA, PARTIDA) for a group.  Raises if not found."""
    cod_up  = codigo.strip().upper()
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

def aggregate(groups: list[dict], rules: list[dict]) -> list[dict]:
    """
    Map each group to (MERCANCIA, PARTIDA) and sum BX, NETO, BRUTO, M2.
    Order of output rows = order of first appearance.
    """
    seen:   OrderedDict[tuple, dict] = OrderedDict()
    errors: list[str] = []

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
                "valor":     0.0,   # declared value — not in Hoja5; left 0
                "bruto":     0.0,
                "neto":      0.0,
                "m2":        0.0,
            }
        seen[key]["bx"]    += g["bx"]
        seen[key]["bruto"] += g["bruto"]
        seen[key]["neto"]  += g["neto"]
        seen[key]["m2"]    += g["m2"]

    if errors:
        unmapped = "\n  ".join(errors)
        raise RuntimeError(
            f"Could not map {len(errors)} group(s) — update product_mapping.csv:\n  {unmapped}"
        )

    result = list(seen.values())
    for r in result:
        r["bruto"] = round(r["bruto"], 2)
        r["neto"]  = round(r["neto"],  2)
        r["m2"]    = round(r["m2"],    2)

    log.info("Aggregated into %d MERCANCIA/PARTIDA rows.", len(result))
    return result

# ─────────────────────────────────────────────────────────────────────────────
# Step 4 – write output ODS  (Hoja4 format)
# ─────────────────────────────────────────────────────────────────────────────

def _make_cell(doc, value: str, bold: bool = False, bg: str = None) -> TableCell:
    """Create a styled ODS table cell with a string value."""
    style_name = None

    if bold or bg:
        style = Style(name=f"ce_{id(value)}_{bold}_{bg}", family="table-cell")
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
    if bold:
        # Inline bold via text style
        ts = Style(name=f"ts_{id(value)}", family="text")
        ts.addElement(TextProperties(fontweight="bold"))
        doc.automaticstyles.addElement(ts)
        p.setAttribute("stylename", ts.getAttribute("name"))
    cell.addElement(p)
    return cell


def _row_from_values(doc, values: list, bold=False, bg=None) -> TableRow:
    row = TableRow()
    for v in values:
        row.addElement(_make_cell(doc, str(v) if v is not None else "", bold=bold, bg=bg))
    return row


def write_output(summary: list[dict], output_path: str) -> None:
    """Write the Hoja4-equivalent ODS."""
    doc   = OpenDocumentSpreadsheet()
    sheet = Table(name="Resumen_Partidas")
    doc.spreadsheet.addElement(sheet)

    HEADER_BG = "#C0C0C0"

    # ── Header row ──────────────────────────────────────────────────────────
    headers = [
        "MERCANCIA", "PARTIDA",
        "Suma - BX", "Suma - VALOR",
        "Suma - BRUTO", "Suma - NETO", "Suma - M22",
    ]
    sheet.addElement(_row_from_values(doc, headers, bold=True, bg=HEADER_BG))

    # ── Data rows ───────────────────────────────────────────────────────────
    for r in summary:
        def fmt(v):
            return _fmt(v) if isinstance(v, float) and v != 0 else (str(v) if v else "")

        row_vals = [
            r["mercancia"],
            r["partida"],
            str(r["bx"]),
            fmt(r["valor"])  if r["valor"] else "",
            fmt(r["bruto"])  if r["bruto"] else "",
            fmt(r["neto"])   if r["neto"]  else "",
            fmt(r["m2"])     if r["m2"]    else "",
        ]
        sheet.addElement(_row_from_values(doc, row_vals))

    doc.save(output_path)
    log.info("Output written to %s", output_path)

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Croton packing-list ODS → customs summary ODS"
    )
    parser.add_argument("input_ods",  help="Input ODS (Hoja5 format)")
    parser.add_argument("output_ods", help="Output ODS (Hoja4 summary)")
    parser.add_argument(
        "--mapping",
        default=str(MAPPING_FILE),
        help="Path to product_mapping.csv (default: next to this script)",
    )
    args = parser.parse_args()

    if not os.path.exists(args.input_ods):
        log.error("Input file not found: %s", args.input_ods)
        sys.exit(1)

    if not os.path.exists(args.mapping):
        log.error("Mapping file not found: %s", args.mapping)
        sys.exit(1)

    groups  = read_hoja5(args.input_ods)
    rules   = load_mapping(args.mapping)
    summary = aggregate(groups, rules)
    write_output(summary, args.output_ods)

    print(f"✅ Done — {len(summary)} partidas aduaneras → {args.output_ods}")


if __name__ == "__main__":
    main()
