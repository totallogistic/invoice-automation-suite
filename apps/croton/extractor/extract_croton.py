#!/usr/bin/env python3
"""
Croton Packing List Processor
==============================
Input : ODS packing-list (Hoja5 / Hoja1 format — one row per roll/bulto)
        [optional] FACTURA XLSX — invoice with commercial values per CODIGO
Output: ODS summary aggregated by customs tariff (Hoja4 format)

Transformation steps:
  Hoja5  →  group rolls by consecutive CANT_TOTAL markers
         →  [optional] enrich VALOR from FACTURA XLSX
         →  map (CODIGO, DESCRIPCION) to (MERCANCIA, PARTIDA)  [product_mapping.csv]
         →  aggregate by (MERCANCIA, PARTIDA)  →  Hoja4 output ODS

VALOR logic (when --factura is provided):
  For each unique CODIGO in the packing list, the TOTAL invoice amount for
  that CODIGO is placed on its LAST group. All prior groups of the same
  CODIGO keep VALOR=0. This matches the manual workflow observed in Hoja1.
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
# Step 1 – read Hoja5 (or Hoja1) and group rolls
# ─────────────────────────────────────────────────────────────────────────────
# Column layout (0-based, after header rows):
#   0=LINEA  1=NBULTO  2=DESCRIPCION  3=CODIGO  4=NPALET
#   5=PESO_NETO  6=PESO_BRUTO  7=CANT  8=CANT_TOTAL  9=M2  10=VALOR (optional)
#
# VALOR (col 10) is filled on the last row of each group in some formats.
# When a FACTURA is provided, missing VALOR values are computed automatically.

# Sheet names to try, in order of preference
_SHEET_CANDIDATES = ["Hoja5", "Hoja1"]

# CROTON Hoja5 format has 8 header rows (company info + column headers).
# Simpler formats (e.g. 2.ods Hoja1) have just 1 header row.
# Auto-detected by inspecting whether the first non-empty col-0 value is numeric.
_HEADER_ROWS_DEFAULT = 8


def _detect_header_rows(all_rows: list[list[str]]) -> int:
    """
    Return the number of header rows to skip.
    Scans from the top until a row whose first cell looks like a line number (digit).
    """
    for i, row in enumerate(all_rows):
        if row and row[0].strip().lstrip("-").isdigit():
            return i
    return _HEADER_ROWS_DEFAULT


def read_hoja5(ods_path: str) -> list[dict]:
    """
    Read the input ODS and return a list of 'group' dicts, each representing
    one aggregated shipment unit (one entry in Hoja4):

      {
        'codigo':      str,
        'descripcion': str,   # from first row in group
        'bx':          int,   # number of individual rolls in group
        'neto':        float, # sum of PESO_NETO
        'bruto':       float, # sum of PESO_BRUTO
        'cant_total':  float, # from last row (CANT_TOTAL marker)
        'm2':          float, # from last row
        'valor':       float, # from last row (col 10); 0 if absent
      }

    A 'group' is a run of consecutive data rows that ends when
    CANT_TOTAL (col 8) is non-empty.
    """
    doc    = load(ods_path)
    sheets = doc.spreadsheet.getElementsByType(Table)

    # Find the data sheet: try preferred names, fall back to first sheet
    data_sheet = None
    for name in _SHEET_CANDIDATES:
        for s in sheets:
            if s.getAttribute("name") == name:
                data_sheet = s
                break
        if data_sheet is not None:
            break
    if data_sheet is None:
        data_sheet = sheets[0]
        log.warning(
            "None of %s sheets found; using first sheet '%s'.",
            _SHEET_CANDIDATES,
            data_sheet.getAttribute("name"),
        )
    else:
        log.info("Using sheet '%s' from %s.", data_sheet.getAttribute("name"), ods_path)

    all_rows    = _sheet_rows(data_sheet)
    header_rows = _detect_header_rows(all_rows)
    log.info("Detected %d header row(s) in %s.", header_rows, ods_path)

    # Detect which column contains VALOR by scanning the header row(s).
    # 2.ods Hoja1 has "valor" at col 14; CROTON Hoja5 has no explicit VALOR column
    # (default to col 10 for backward compatibility with files that use that position).
    _valor_col = 10  # default
    for row in all_rows[:header_rows]:
        for i, cell in enumerate(row):
            if cell.strip().lower() == "valor":
                _valor_col = i
                log.info("VALOR column auto-detected at index %d.", _valor_col)
                break

    groups        = []
    current_bx    = 0
    current_neto  = 0.0
    current_bruto = 0.0
    first_codigo  = ""
    first_desc    = ""
    in_group      = False

    for row in all_rows[header_rows:]:
        if not row:
            continue

        linea = row[0] if len(row) > 0 else ""
        # Skip non-data rows (totals, palet summaries, etc.)
        if not linea or not linea.strip().lstrip("-").isdigit():
            continue

        # Pad to at least 10 cols
        while len(row) < 10:
            row.append("")

        descripcion = row[2]
        codigo      = row[3]
        peso_neto   = _parse_number(row[5])
        peso_bruto  = _parse_number(row[6])
        cant_total  = row[8].strip()   # non-empty → end of group
        m2          = row[9].strip()
        valor       = row[_valor_col].strip() if len(row) > _valor_col else ""

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
                "valor":       _parse_number(valor),
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
            "valor":       0.0,
        })

    log.info("Read %d groups from %s.", len(groups), ods_path)
    return groups

# ─────────────────────────────────────────────────────────────────────────────
# Step 1b – load FACTURA XLSX and enrich groups with VALOR
# ─────────────────────────────────────────────────────────────────────────────
# FACTURA column layout (sheet "TOTAL FRA." or first sheet):
#   col 0 = REFERENCIA (product code, matches CODIGO in packing list)
#   col 5 = CANTIDAD MTS / KGS
#   col 8 = IMPORTE (€)
#
# VALOR assignment rule (matches observed manual workflow in Hoja1):
#   For each unique CODIGO in the groups list, the TOTAL FACTURA IMPORTE
#   for that CODIGO is assigned to its LAST group.  All earlier groups of
#   the same CODIGO retain their existing VALOR (usually 0 from Hoja5).
#   CODIGOs absent in the FACTURA are left with VALOR=0.

# Known non-data row prefixes in the FACTURA
_FACTURA_SKIP_PREFIXES = (
    "referencia", "importe", "forma", "total", "transferencia",
    "domicilio", "incoterm",
)


def load_factura(xlsx_path: str) -> dict[str, float]:
    """
    Load invoice XLSX and return {CODIGO_UPPER: total_importe_euros}.

    Multiple rows with the same REFERENCIA are summed (e.g. RF385 BLANCA +
    RF385 COLORES → single total for RF385).
    """
    try:
        import openpyxl
    except ImportError:
        log.error("openpyxl is required for FACTURA support: pip install openpyxl")
        sys.exit(1)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # Prefer "TOTAL FRA." sheet; fall back to first sheet
    sheet_name = "TOTAL FRA." if "TOTAL FRA." in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    log.info("Reading FACTURA from sheet '%s' in %s.", sheet_name, xlsx_path)

    totals: dict[str, float] = defaultdict(float)

    for row in ws.iter_rows(values_only=True):
        raw_codigo = row[0]
        if raw_codigo is None:
            continue
        codigo = str(raw_codigo).strip()
        if not codigo:
            continue
        # Skip header / summary rows
        if codigo.lower().split()[0] in _FACTURA_SKIP_PREFIXES:
            continue
        # Skip rows without a numeric IMPORTE in col 8
        importe_raw = row[8]
        if importe_raw is None:
            continue
        try:
            importe = float(importe_raw)
        except (ValueError, TypeError):
            continue
        # Only include meaningful product lines (importe > 1 €)
        if importe < 1.0:
            continue

        totals[codigo.upper()] += round(importe, 4)

    result = dict(totals)
    log.info(
        "Loaded FACTURA: %d distinct CODIGOs, total %.2f €.",
        len(result),
        sum(result.values()),
    )
    for cod, val in sorted(result.items()):
        log.debug("  FACTURA  %-15s  %.2f €", cod, val)
    return result


def enrich_valor_from_factura(
    groups: list[dict],
    factura: dict[str, float],
) -> list[dict]:
    """
    Assign invoice VALOR to groups.

    Rule: for each unique CODIGO, place the TOTAL FACTURA amount on the
    LAST group of that CODIGO.  Previous groups of the same CODIGO keep
    their current valor (normally 0).  CODIGOs not in the FACTURA are
    left untouched (valor stays 0 unless already set from Hoja5 col 10).
    """
    # Find the index of the last occurrence of each CODIGO
    last_idx: dict[str, int] = {}
    for i, g in enumerate(groups):
        last_idx[g["codigo"].upper()] = i

    # Detect CODIGOs that already carry a non-zero VALOR from the ODS itself.
    # For these, the ODS value represents the total for that CODIGO (as produced
    # by the spreadsheet formulas in files like 2.ods / Hoja1), so we must not
    # also add the FACTURA amount — that would double-count the value.
    already_valued: set[str] = {
        g["codigo"].upper()
        for g in groups
        if g["valor"] != 0.0
    }
    if already_valued:
        log.info(
            "CODIGOs already carrying ODS VALOR (skipped for FACTURA enrichment): %s",
            ", ".join(sorted(already_valued)),
        )

    enriched = 0
    for i, g in enumerate(groups):
        cod = g["codigo"].upper()
        if cod in already_valued:
            continue  # ODS value takes precedence for this CODIGO
        if i == last_idx[cod] and cod in factura:
            g["valor"] = factura[cod]
            enriched += 1
            log.debug(
                "Assigned VALOR %.2f to last group of %s (group %d).",
                g["valor"], cod, i,
            )

    log.info(
        "Enriched %d group(s) with VALOR from FACTURA (%d CODIGOs matched).",
        enriched,
        len([c for c in last_idx if c in factura]),
    )
    return groups

# ─────────────────────────────────────────────────────────────────────────────
# Step 2 – load product mapping CSV
# ─────────────────────────────────────────────────────────────────────────────

def load_mapping(csv_path: str) -> list[dict]:
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
    Map each group to (MERCANCIA, PARTIDA) and sum BX, NETO, BRUTO, M2, VALOR.
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
            f"Could not map {len(errors)} group(s) — update product_mapping.csv:\n  {unmapped}"
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


def inject_sheet_into_ods(summary: list[dict], source_ods: str, output_path: str) -> None:
    """
    Copy *source_ods* to *output_path* and append a 'Resumen_Partidas' sheet
    containing the summary rows.  Existing sheets are preserved unchanged.

    This is the preferred output mode for the unified processor workflow:
    the user receives the original packing-list ODS with the customs summary
    added as an extra sheet, rather than a separate standalone ODS.
    """
    import shutil

    # Work on a copy so the original is never modified
    shutil.copy2(source_ods, output_path)

    doc = load(output_path)

    # Remove any pre-existing 'Resumen_Partidas' sheet (idempotent re-runs)
    for existing in doc.spreadsheet.getElementsByType(Table):
        if existing.getAttribute("name") == "Resumen_Partidas":
            doc.spreadsheet.removeChild(existing)
            log.info("Replaced existing 'Resumen_Partidas' sheet in %s.", output_path)
            break

    # Build the new summary sheet and attach it
    sheet = Table(name="Resumen_Partidas")
    doc.spreadsheet.addElement(sheet)
    _fill_summary_sheet(doc, sheet, summary)

    doc.save(output_path)
    log.info("Injected 'Resumen_Partidas' sheet into %s", output_path)


def write_output(summary: list[dict], output_path: str) -> None:
    """Write the Hoja4-equivalent ODS (standalone, no source ODS required)."""
    doc   = OpenDocumentSpreadsheet()
    sheet = Table(name="Resumen_Partidas")
    doc.spreadsheet.addElement(sheet)
    _fill_summary_sheet(doc, sheet, summary)
    doc.save(output_path)
    log.info("Output written to %s", output_path)


def _fill_summary_sheet(doc, sheet, summary: list[dict]) -> None:
    """Populate *sheet* with headers + data rows.  Shared by both output modes."""

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

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Croton packing-list ODS → customs summary ODS"
    )
    parser.add_argument(
        "input_ods",
        help="Input ODS (Hoja5 or Hoja1 format)",
    )
    parser.add_argument(
        "-o", "--output",
        dest="output_ods",
        help="Output ODS path (used by unified processor)",
    )
    parser.add_argument(
        "output_ods_pos",
        nargs="?",
        help=argparse.SUPPRESS,  # legacy second positional
    )
    parser.add_argument(
        "--mapping",
        default=str(MAPPING_FILE),
        help="Path to product_mapping.csv (default: next to this script)",
    )
    parser.add_argument(
        "--factura",
        default=None,
        help=(
            "Path to the FACTURA XLSX file.  When provided, the commercial "
            "VALOR for each CODIGO is read from the invoice and assigned to "
            "the last group of that CODIGO before aggregation."
        ),
    )
    parser.add_argument(
        "--inject",
        action="store_true",
        default=False,
        help=(
            "Instead of creating a standalone output ODS, copy the input ODS "
            "to the output path and add 'Resumen_Partidas' as a new sheet. "
            "This is the default mode used by the unified processor."
        ),
    )
    args = parser.parse_args()

    # -o flag takes priority; fall back to second positional
    output_path = args.output_ods or args.output_ods_pos
    if not output_path:
        parser.error(
            "Output path required: use -o <path> or pass as second positional argument"
        )

    if not os.path.exists(args.input_ods):
        log.error("Input file not found: %s", args.input_ods)
        sys.exit(1)

    if not os.path.exists(args.mapping):
        log.error("Mapping file not found: %s", args.mapping)
        sys.exit(1)

    if args.factura and not os.path.exists(args.factura):
        log.error("FACTURA file not found: %s", args.factura)
        sys.exit(1)

    # Validate output path is a file path, not a directory.
    # The unified processor always passes a full file path via -o.
    # Passing a directory would silently produce a wrong filename — fail fast instead.
    if os.path.isdir(output_path):
        log.error(
            "Output path '%s' is a directory. Pass a full file path with -o, "
            "e.g. -o /data/croton/out/batch_id/CROTON_my_file.ods",
            output_path,
        )
        sys.exit(1)

    # ── Pipeline ────────────────────────────────────────────────────────────
    groups = read_hoja5(args.input_ods)

    if args.factura:
        factura = load_factura(args.factura)
        groups  = enrich_valor_from_factura(groups, factura)
    else:
        log.info(
            "No --factura provided; VALOR will be taken from the ODS (col 10) only."
        )

    rules   = load_mapping(args.mapping)
    summary = aggregate(groups, rules)

    if args.inject:
        inject_sheet_into_ods(summary, args.input_ods, output_path)
    else:
        write_output(summary, output_path)

    # ── Summary report ──────────────────────────────────────────────────────
    total_bx    = sum(r["bx"]    for r in summary)
    total_valor = sum(r["valor"] for r in summary)
    total_bruto = sum(r["bruto"] for r in summary)
    total_neto  = sum(r["neto"]  for r in summary)
    total_m2    = sum(r["m2"]    for r in summary)

    print(f"✅ Done — {len(summary)} partidas aduaneras → {output_path}")
    print(f"   Totales:  BX={total_bx}  VALOR={total_valor:.2f}€  "
          f"BRUTO={total_bruto:.2f}kg  NETO={total_neto:.2f}kg  M2={total_m2:.2f}")


if __name__ == "__main__":
    main()