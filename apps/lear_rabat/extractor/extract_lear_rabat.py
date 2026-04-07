#!/usr/bin/env python3
"""
Lear Rabat Invoice Extractor - Unified Architecture Version
Extracts data from Lear Rabat invoices and generates ODS files with DUA data.

Usage (unified processor interface):
  extract_lear_rabat.py <pdf1> <pdf2> ... -o <output_dir> [-t <template.ods>]
"""

import argparse
import re
import sys

SCRIPT_VERSION = "2026-03-12.v40"

SCRIPT_CHANGELOG = """
## 2026-03-12.v40

### Logica general
Extrae datos de facturas PDF de Lear Rabat (Marruecos) y genera archivos ODS
listos para tramitacion DUA, usando una plantilla ODS como base.

### Extraccion PDF
- Detecta automaticamente formato de numeros US (1,234.56) o EU (1.234,56)
- Extrae por factura: numero, fecha, pallets, peso bruto/neto totales
- Extrae por linea: codigo parte, descripcion, codigo HS, cantidad, precio
  unitario, subtotal, OPR Material Cost, Partial Weight, pallets por proyecto

### Columnas generadas en Sheet1
- G: Codigo parte (numerico, sin ceros iniciales)
- J: Valor DUA (subtotal FG)
- K: OPR Material Cost (7009)
- L: Pallets por proyecto
- M: V.E = J + K  (formula preservada del template)
- N: PESO BR = (O * R2) / R4  (formula preservada, cache actualizado)
- O: PESO NET por linea (2 decimales, del PDF)
- P: Unidades (cantidad)
- R: FACTURA - valores totales segun el PDF (R2=PESO BRUT, R4=PESO NET)
- S: CALCULO - nuestros totales calculados (pre-ajuste)
- T: DIFF - diferencia CALCULO - FACTURA
- U: Nota de ajuste automatico (registro, valor antes y despues)

### Ajuste automatico de pesos
Si la suma de PESO NET de las lineas no coincide exactamente con el total
del PDF, la diferencia se suma al ultimo registro para que la columna O
sume exactamente el total de la factura. PESO BR se recalcula
automaticamente via la formula de la columna N al abrir el archivo.
La columna R/S/T/U documenta el ajuste para trazabilidad.

### Hoja Resumen
Agrupa por (DESCRIPCION, PARTIDA) con totales de:
PALETS, VALOR DUA, 7009, PESO BR, PESO NET, UN.
- PESO NET por grupo: suma de decimales de Sheet1
- PESO BR total: total_peso_net * ratio (evita acumulacion de redondeos)
- Fila Total Resultado refleja los mismos totales que Sheet1
"""
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional

SCRIPT_VERSION = "2026-03-07.v31"

try:
    import pdfplumber
except ImportError:
    print("ERROR: missing dependency. Install with: pip install pdfplumber", file=sys.stderr)
    sys.exit(1)

try:
    from odf import opendocument, table, text as odftext
    from odf.opendocument import load as load_ods
    from odf.namespaces import TABLENS, OFFICENS
except ImportError:
    print("ERROR: missing dependency. Install with: pip install odfpy", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# PDF extraction helpers
# ---------------------------------------------------------------------------

def read_pdf_text(pdf_path: Path) -> str:
    parts = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                parts.append(text)
    return "\n".join(parts)


def detect_number_format(text: str) -> str:
    us_pattern = re.findall(r'\d{1,3},\d{3}\.\d{2}', text)
    eu_pattern = re.findall(r'\d{1,3}\s\d{3},\d{2}', text)
    return "US" if len(us_pattern) > len(eu_pattern) else "EU"


def parse_number(s: str, fmt: str = "EU") -> Optional[float]:
    if not s or s.strip() == "":
        return None
    s = s.strip()
    if fmt == "US":
        s = s.replace(",", "")
    else:
        s = s.replace(" ", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def extract_invoice_number(text: str) -> str:
    lines = text.splitlines()
    for line in lines[:20]:
        m = re.match(r'^(\d{8})\s', line.strip())
        if m:
            return m.group(1)
        if re.match(r'^\d{8}$', line.strip()):
            return line.strip()
    return ""


def extract_invoice_metadata(text: str, fmt: str) -> Dict:
    lines = text.splitlines()
    invoice_no = extract_invoice_number(text)
    total_invoice = None
    peso_brut = None
    peso_net = None
    pallets = None
    total_opr = None

    for i, line in enumerate(lines):
        m = re.search(r'Total:\s+([\d\s,\.]+)', line)
        if m:
            total_invoice = parse_number(m.group(1), fmt)

        if "Net:" in line and "Gross:" in line:
            # K unit is optional (some invoices omit it)
            m = re.search(r'Net:\s*([\d\s,\.]+?)\s*K?\s+.*?Gross:\s*([\d\s,\.]+?)\s*K?\s*$', line)
            if m:
                peso_net = parse_number(m.group(1).strip(), "EU")
                peso_brut = parse_number(m.group(2).strip(), "EU")

        if i < 15 and ("FCA" in line or "EXW" in line or "DAP" in line):
            m = re.search(r'\b(\d{1,3})\s*$', line.strip())
            if m:
                pallets = int(m.group(1))

    for line in reversed(lines):
        m = re.search(r'OPR Material Cost:\s+([\d\s,\.]+)', line)
        if m:
            total_opr = parse_number(m.group(1), fmt)
            break

    return {
        "invoice_no": invoice_no,
        "total_invoice": total_invoice,
        "peso_brut": peso_brut,
        "peso_net": peso_net,
        "pallets": pallets,
        "total_opr": total_opr,
    }


def extract_line_items(text: str, fmt: str) -> List[Dict]:
    lines = text.splitlines()
    raw_items = []

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        fg_match = re.match(
            r'(\d+)\s+FG\s+(\S+)\s+ESOPO\d+-\s*(\d+)I?\s+(?:(\d{7,})\s+)?([\d,\.\s]+)',
            line
        )
        if fg_match:
            line_no = fg_match.group(1)
            part_number = fg_match.group(2)
            code = fg_match.group(3)
            hs_code = fg_match.group(4) or ""
            qty_str = fg_match.group(5).strip()
            quantity = parse_number(qty_str, fmt)

            subtotal_price = None
            partial_weight = None

            for j in range(i + 1, min(i + 50, len(lines))):
                check_line = lines[j].strip()
                if "Subtotal FG:" in check_line:
                    m = re.search(r'Subtotal FG:\s+\S+\s+([\d\s,\.]+)', check_line)
                    if m:
                        subtotal_price = parse_number(m.group(1), fmt)
                    break
                if "Partial Weight:" in check_line:
                    m = re.search(r'Partial Weight:\s+([\d\s,\.]+)\s*K?', check_line)
                    if m:
                        partial_weight = parse_number(m.group(1), fmt)
                if re.match(r'\d+\s+FG\s+', check_line):
                    break

            raw_items.append({
                "line_no": line_no,
                "line_idx": i,
                "part_number": part_number,
                "code": code,
                "hs_code": hs_code,
                "quantity": quantity,
                "subtotal_price": subtotal_price,
                "partial_weight": partial_weight,
            })
        i += 1

    projects = []
    for i, line_text in enumerate(lines):
        line = line_text.strip()
        project_match = re.match(r'(O_\w+)\s+Project:\s+(\d+)\s+Plts', line)
        if project_match:
            project_name = project_match.group(1)
            project_pallets = int(project_match.group(2))
            opr_material = None
            for j in range(i + 1, min(i + 10, len(lines))):
                m = re.search(r'OPR Material Cost:\s+([\d\s,\.]+)', lines[j])
                if m:
                    opr_material = parse_number(m.group(1), fmt)
                    break
            projects.append({
                "name": project_name,
                "pallets": project_pallets,
                "opr_material": opr_material,
                "line_idx": i,
            })

    for item in raw_items:
        item_line = item["line_idx"]
        best_project = None
        for proj in projects:
            if proj["line_idx"] > item_line:
                best_project = proj
                break
        if best_project:
            item["project"] = best_project["name"]
            item["project_pallets"] = best_project["pallets"]
            item["opr_material"] = best_project["opr_material"]
        else:
            item["project"] = None
            item["project_pallets"] = None
            item["opr_material"] = None

    project_items = {}
    for idx, item in enumerate(raw_items):
        proj = item.get("project")
        if proj:
            if proj not in project_items:
                project_items[proj] = []
            project_items[proj].append(idx)

    last_in_project = set()
    for proj, indices in project_items.items():
        last_in_project.add(indices[-1])

    for idx, item in enumerate(raw_items):
        if idx not in last_in_project:
            item["opr_material"] = None
            item["project_pallets"] = None

    for item in raw_items:
        item.pop("line_idx", None)

    return raw_items


# ---------------------------------------------------------------------------
# ODS helpers
# ---------------------------------------------------------------------------

def get_cell_value(cell):
    for p in cell.getElementsByType(odftext.P):
        return str(p).strip()
    return ""


def set_cell_value(cell, value):
    if cell is None:
        return
    try:
        cell.removeAttribute("formula")
    except Exception:
        pass
    try:
        cell.removeAttribute("valuetype")
    except Exception:
        pass
    try:
        cell.removeAttribute("value")
    except Exception:
        pass
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except Exception:
                pass
    except Exception:
        pass
    p = odftext.P()
    p.addText(str(value))
    cell.appendChild(p)


def set_numeric_value(cell, numeric_value, display_text=None):
    """Set a numeric value, removing any formula (use for data cells we own)."""
    if cell is None:
        return
    float_val = float(numeric_value)
    try:
        cell.removeAttribute("formula")
    except Exception:
        pass
    cell.setAttribute("valuetype", "float")
    cell.setAttribute("value", str(float_val))
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except Exception:
                pass
    except Exception:
        pass
    if display_text is None:
        display_text = format_eu_number(float_val)
    p = odftext.P()
    p.addText(str(display_text))
    cell.appendChild(p)


def update_formula_cache(cell, numeric_value, display_text=None):
    """Update the cached office:value of a formula cell WITHOUT touching the formula.

    This is critical for PESO BR and V.E cells: the formula must stay intact
    so LibreOffice can recalculate, but the cached value must be correct so
    pivot tables built before recalculation work correctly.
    """
    if cell is None:
        return
    float_val = float(numeric_value)
    # Update cached value attributes only - do NOT touch table:formula
    cell.setAttribute("valuetype", "float")
    cell.setAttribute("value", str(float_val))
    # Update display text
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except Exception:
                pass
    except Exception:
        pass
    if display_text is None:
        display_text = format_eu_number(float_val)
    p = odftext.P()
    p.addText(str(display_text))
    cell.appendChild(p)


def clear_cell(cell):
    if cell is None:
        return
    try:
        cell.removeAttribute("formula")
    except Exception:
        pass
    try:
        cell.removeAttribute("valuetype")
    except Exception:
        pass
    try:
        cell.removeAttribute("value")
    except Exception:
        pass
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except Exception:
                pass
    except Exception:
        pass


def format_eu_number(value, decimals=2):
    """Format number EU style (comma decimal). Always shows at least 2 decimal places.
    Extra trailing zeros beyond 2 decimals are stripped, but 2 are always kept.
    e.g. 8647.5 -> 8647,50 | 134.75 -> 134,75 | 10761.622 -> 10761,62
    """
    if value is None:
        return ""
    # Round to 2 decimals for display consistency
    rounded = round(float(value), decimals)
    formatted = f'{rounded:.{decimals}f}'.replace(".", ",")
    return formatted


def build_cell_map(row, writable_cols=None):
    """Build a mapping from column index to cell, splitting repeated cells
    only for columns in writable_cols.

    IMPORTANT: Formula-only columns (V.E col 12, PESO BR col 13) should NOT
    be in writable_cols for data rows - they are standalone cells and are
    accessible without splitting.
    """
    if writable_cols is None:
        writable_cols = set()

    cells = list(row.getElementsByType(table.TableCell))
    cell_map = {}
    col = 0

    for cell in cells:
        rep = cell.getAttribute("numbercolumnsrepeated")
        rep = int(rep) if rep else 1

        needs_split = any((col + r) in writable_cols for r in range(rep))

        if needs_split and rep > 1:
            try:
                cell.removeAttribute("numbercolumnsrepeated")
            except Exception:
                pass

            cell_map[col] = cell
            prev = cell
            for r in range(1, rep):
                new_cell = table.TableCell()
                val = get_cell_value(cell)
                if val:
                    p = odftext.P()
                    p.addText(val)
                    new_cell.appendChild(p)
                parent = cell.parentNode
                next_sib = prev.nextSibling
                if next_sib:
                    parent.insertBefore(new_cell, next_sib)
                else:
                    parent.appendChild(new_cell)
                cell_map[col + r] = new_cell
                prev = new_cell
        else:
            for r in range(rep):
                cell_map[col + r] = cell

        col += rep

    return cell_map


def get_col_b_value(cell_map):
    cell = cell_map.get(1)
    if cell:
        return get_cell_value(cell)
    return ""


# ---------------------------------------------------------------------------
# Template code mapping (for summary sheet)
# ---------------------------------------------------------------------------

def build_code_mapping(doc) -> Dict[int, tuple]:
    """Read col B -> (col C DESCRIPCION, col D PARTIDA) from Sheet1.

    The VLOOKUP formula in Sheet1 is: =VLOOKUP(G_row, $B$1:$D$247, 2/3, 0)
    So col B is the lookup key, col C = DESCRIPCION, col D = PARTIDA.
    We read all 247 rows to build the complete mapping.
    """
    sheet = doc.spreadsheet.getElementsByType(table.Table)[0]
    rows = sheet.getElementsByType(table.TableRow)
    mapping = {}

    for row in rows[:250]:
        cells = list(row.getElementsByType(table.TableCell))
        col_vals = []
        for cell in cells:
            rep = cell.getAttribute("numbercolumnsrepeated")
            rep = int(rep) if rep else 1
            val = get_cell_value(cell)
            num_val = cell.getAttribute("value")
            for _ in range(rep):
                col_vals.append((val, num_val))

        if len(col_vals) < 4:
            continue

        # col B = index 1, col C = index 2, col D = index 3
        b_val, b_num = col_vals[1]
        c_val, _ = col_vals[2]
        d_val, _ = col_vals[3]

        # Try to get numeric code from col B
        code_str = b_num or b_val
        if code_str:
            try:
                code_int = int(float(code_str))
                if code_int > 0 and c_val:
                    mapping[code_int] = (c_val, d_val)
            except (ValueError, TypeError):
                pass

    return mapping


# ---------------------------------------------------------------------------
# Summary / pivot sheet generation
# ---------------------------------------------------------------------------

def _ensure_number_styles(doc):
    """Add number format styles to the document for the Resumen sheet."""
    from odf import number as odfnumber
    from odf.style import Style

    fmt2 = odfnumber.Number(decimalplaces="2", minintegerdigits="1")
    nstyle2 = odfnumber.NumberStyle(name="resumen_n2")
    nstyle2.addElement(fmt2)
    doc.styles.addElement(nstyle2)

    fmt0 = odfnumber.Number(decimalplaces="0", minintegerdigits="1")
    nstyle0 = odfnumber.NumberStyle(name="resumen_n0")
    nstyle0.addElement(fmt0)
    doc.styles.addElement(nstyle0)

    cs2 = Style(name="resumen_cell_n2", family="table-cell",
                datastylename="resumen_n2")
    doc.styles.addElement(cs2)

    cs0 = Style(name="resumen_cell_n0", family="table-cell",
                datastylename="resumen_n0")
    doc.styles.addElement(cs0)

    return {"n2": "resumen_cell_n2", "n0": "resumen_cell_n0"}


def add_summary_sheet(doc, items: List[Dict], invoice_data: Dict,
                      code_mapping: Dict[int, tuple], brut_net_ratio: float):
    """Add a 'Resumen' sheet with a pivot-table-like summary grouped by
    (DESCRIPCION, PARTIDA), mirroring the LibreOffice Tabla Dinamica.

    Uses the code_mapping (col B -> col C, col D) from the template to
    resolve each item's code to its description and customs tariff code.
    """
    styles = _ensure_number_styles(doc)

    # --- Compute per-item PESO BR ---
    enriched = []
    for item in items:
        code_raw = item.get("code", "")
        try:
            code_int = int(code_raw.lstrip("0") or "0")
        except ValueError:
            code_int = 0

        descripcion, partida = code_mapping.get(code_int, ("#N/D", "#N/D"))

        pw = round(item.get("partial_weight") or 0, 2)
        peso_br = round(pw * brut_net_ratio, 2)

        enriched.append({
            "descripcion": descripcion,
            "partida": partida,
            "palets": item.get("project_pallets") or 0,
            "valor_dua": item.get("subtotal_price") or 0,
            "opr": item.get("opr_material") or 0,
            "peso_br": peso_br,
            "peso_net": pw,
            "un": item.get("quantity") or 0,
        })

    # --- Group by (DESCRIPCION, PARTIDA) ---
    groups = defaultdict(lambda: {
        "palets": 0, "valor_dua": 0.0, "opr": 0.0,
        "peso_br": 0.0, "peso_net": 0.0, "un": 0
    })
    desc_seen = set()

    for e in enriched:
        key = (e["descripcion"], e["partida"])
        desc_seen.add(e["descripcion"])
        g = groups[key]
        g["palets"] += e["palets"]
        g["valor_dua"] += e["valor_dua"]
        g["opr"] += e["opr"]
        g["peso_br"] += e["peso_br"]
        g["peso_net"] += e["peso_net"]
        g["un"] += int(e["un"])

    # Sort alphabetically by DESCRIPCION then PARTIDA
    sorted_keys = sorted(groups.keys(), key=lambda k: (k[0].upper(), k[1]))

    # --- Build ODS table rows ---
    summary_table = table.Table(name="Resumen")

    def make_row(*cells_data):
        """cells_data: list of (value, is_numeric, bold_header)
        Integer cols: palets(2), peso_br(5), peso_net(6), UN(7).
        """
        INT_COLS = {2, 7}  # palets and UN as integers; peso_br/net keep 2 decimals
        tr = table.TableRow()
        for col_idx, (val, is_num, is_header) in enumerate(cells_data):
            tc = table.TableCell()
            if is_num and val is not None:
                float_val = float(val)
                tc.setAttribute("valuetype", "float")
                tc.setAttribute("value", str(float_val))
                p = odftext.P()
                if col_idx in INT_COLS:
                    p.addText(str(round(float_val)))
                else:
                    p.addText(format_eu_number(float_val))
                tc.appendChild(p)
            else:
                if val:
                    tc.setAttribute("valuetype", "string")
                    p = odftext.P()
                    p.addText(str(val))
                    tc.appendChild(p)
            tr.addElement(tc)
        return tr

    def make_header_row(*labels):
        tr = table.TableRow()
        for label in labels:
            tc = table.TableCell()
            tc.setAttribute("valuetype", "string")
            p = odftext.P()
            p.addText(str(label))
            tc.appendChild(p)
            tr.addElement(tc)
        return tr

    # Header row
    headers = ["DESCRIPCION", "PARTIDA", "Suma - PALETS", "Suma - VALOR DUA",
               "Suma - 7009", "Suma - PESO BR", "Suma - PESO NET", "Suma - UN"]
    summary_table.addElement(make_header_row(*headers))

    # Data rows
    prev_desc = None
    totals = {k: 0 for k in ["palets", "valor_dua", "opr", "peso_br", "peso_net", "un"]}

    for key in sorted_keys:
        desc, partida = key
        g = groups[key]
        display_desc = desc if desc != prev_desc else ""
        prev_desc = desc

        summary_table.addElement(make_row(
            (display_desc, False, False),
            (partida, False, False),
            (g["palets"], True, False),
            (g["valor_dua"], True, False),
            (g["opr"], True, False),
            (g["peso_br"], True, False),
            (g["peso_net"], True, False),
            (g["un"], True, False),
        ))

        totals["palets"]    += g["palets"]
        totals["valor_dua"] += g["valor_dua"]
        totals["opr"]       += g["opr"]
        totals["peso_br"]   += g["peso_br"]
        totals["peso_net"]  += g["peso_net"]
        totals["un"]        += g["un"]

    # Totals row
    summary_table.addElement(make_row(
        ("Total Resultado", False, True),
        ("", False, True),
        (totals["palets"],            True, True),
        (totals["valor_dua"],         True, True),
        (totals["opr"],               True, True),
        # Use total_net * ratio to avoid accumulated rounding errors across rows
        # This gives exactly peso_brut since NET was auto-adjusted to match invoice
        (round(totals["peso_net"] * brut_net_ratio, 2), True, True),
        (round(totals["peso_net"], 2), True, True),
        (totals["un"],                True, True),
    ))

    # Remove existing Resumen sheet if present
    for existing in doc.spreadsheet.getElementsByType(table.Table):
        if existing.getAttribute("name") == "Resumen":
            doc.spreadsheet.removeChild(existing)
            break

    doc.spreadsheet.addElement(summary_table)
    print(f"  [OK] Resumen sheet: {len(sorted_keys)} product groups")


# ---------------------------------------------------------------------------
# Main ODS update
# ---------------------------------------------------------------------------

def update_ods_template(template_path: Path, items: List[Dict],
                        invoice_data: Dict, output_path: Path):
    """Update ODS template with extracted invoice data.

    Column layout (0-based):
      A=0  B=1  C=2  D=3  E=4  F=5  G=6  H=7  I=8
      J=9  K=10 L=11 M=12 N=13 O=14 P=15 Q=16 R=17

    Written by extractor:  G(6) J(9) K(10) L(11) O(14) P(15) R(17)
    Formula cells (do NOT remove formula):  M(12)=V.E=J+K  N(13)=PESO_BR=(O*R2)/R4
    We DO update cached office:value for M and N so pivot tables work immediately.
    """

    doc = load_ods(str(template_path))

    # Build code->(DESCRIPCION, PARTIDA) mapping BEFORE writing (template is clean)
    code_mapping = build_code_mapping(doc)

    sheet = doc.spreadsheet.getElementsByType(table.Table)[0]
    rows = sheet.getElementsByType(table.TableRow)

    COL_CODE_G = 6
    COL_VALOR_DUA = 9
    COL_OPR_MAT = 10
    COL_PALETS = 11
    COL_VE = 12       # V.E = J+K  <- FORMULA, update cache only
    COL_PESO_BR = 13  # PESO BR = (O*R2)/R4  <- FORMULA per row, value for totals
    COL_PESO_NET = 14
    COL_UN = 15
    COL_PESO_VAL = 17   # R: FACTURA values (PDF)
    COL_CALCULO  = 18   # S: our calculated values
    COL_DIFF     = 19   # T: difference CALCULO - FACTURA
    COL_NOTE     = 20   # U: adjustment note

    # Columns where we split and write direct values (no formula kept)
    DATA_WRITE_COLS = {COL_CODE_G, COL_VALOR_DUA, COL_OPR_MAT, COL_PALETS,
                       COL_PESO_NET, COL_UN, COL_PESO_VAL, COL_CALCULO, COL_DIFF, COL_NOTE}

    # For totals row: also write direct value to PESO BR (replaces formula there)
    TOTAL_WRITE_COLS = DATA_WRITE_COLS | {COL_PESO_BR}

    # Columns to clear on data rows that are not used
    CLEAR_COLS = {COL_CODE_G, COL_VALOR_DUA, COL_OPR_MAT, COL_PALETS,
                  COL_PESO_NET, COL_UN}

    # Pre-calculate brut/net ratio for PESO BR formula cache update
    peso_brut = invoice_data.get("peso_brut") or 0
    peso_net_total = invoice_data.get("peso_net") or 0
    brut_net_ratio = peso_brut / peso_net_total if peso_net_total > 0 else 1.0

    # Auto-adjust last item's partial_weight so NET sum == PDF total exactly.
    # This means PESO BR (formula =(O*R2)/R4) will also recalculate to match
    # the invoice total automatically when the file is opened in LibreOffice.
    if items and peso_net_total:
        raw_sum = sum(item.get("partial_weight") or 0 for item in items)
        net_diff = round(peso_net_total - raw_sum, 4)
        if net_diff != 0:
            last = items[-1]
            _pw_before = round((last.get("partial_weight") or 0), 2)
            last["partial_weight"] = round(_pw_before + net_diff, 2)
            _pw_after  = last["partial_weight"]
            # Store pre-adjustment values and item info for display in ODS col R/S/T
            invoice_data["_adj_net_diff"]      = net_diff
            invoice_data["_adj_net_raw"]       = round(raw_sum, 2)
            invoice_data["_adj_br_raw"]        = round(raw_sum * brut_net_ratio, 2)
            invoice_data["_adj_item_code"]     = last.get("code", "")
            invoice_data["_adj_item_idx"]      = len(items) - 1
            invoice_data["_adj_item_pw_before"]= _pw_before
            invoice_data["_adj_item_pw_after"] = _pw_after
            print(f"  [AUTO] PESO NET ajustado: +{net_diff} al ultimo item "
                  f"({last.get('code','')}) -> suma={peso_net_total}")

    # Pre-calculate totals
    total_valor = sum(item.get("subtotal_price") or 0 for item in items)
    total_opr = sum(item.get("opr_material") or 0 for item in items)
    total_pallets = sum(item.get("project_pallets") or 0 for item in items)

    item_idx = 0

    for row_idx, row in enumerate(rows):
        if row_idx == 0:
            continue

        # Use different writable cols for totals vs data rows to avoid
        # splitting formula cells (M=V.E, N=PESO BR) unnecessarily
        is_totals_row = False
        cell_map_check = build_cell_map(row, set())  # peek without splitting
        col_b = get_col_b_value(cell_map_check)
        if col_b == "520":
            is_totals_row = True

        if is_totals_row:
            cell_map = build_cell_map(row, TOTAL_WRITE_COLS)
        else:
            # DATA_WRITE_COLS does NOT include COL_VE (12) or COL_PESO_BR (13)
            # Those formula cells are standalone in the template and accessible
            # directly without needing a split.
            cell_map = build_cell_map(row, DATA_WRITE_COLS)

        if is_totals_row:
            # Preserve the SUM() formulas the template has in this row.
            # We only update the cached office:value so the file opens correctly
            # without recalculation, while letting LibreOffice/Excel recompute
            # the real sum from the data rows we just wrote.
            total_valor = sum(item.get("subtotal_price") or 0 for item in items)
            total_opr   = sum(item.get("opr_material") or 0 for item in items)
            total_pallets = sum(item.get("project_pallets") or 0 for item in items)
            # Decimal sums - match what the per-row values actually add up to
            sum_peso_net = round(sum(item.get("partial_weight") or 0 for item in items), 2)
            sum_peso_br  = round(sum_peso_net * brut_net_ratio, 2)
            update_formula_cache(cell_map.get(COL_VALOR_DUA), total_valor,
                                 format_eu_number(total_valor))
            update_formula_cache(cell_map.get(COL_OPR_MAT), total_opr,
                                 format_eu_number(total_opr))
            update_formula_cache(cell_map.get(COL_PALETS), total_pallets, str(total_pallets))
            update_formula_cache(cell_map.get(COL_PESO_BR), sum_peso_br,
                                 str(round(sum_peso_br)))
            update_formula_cache(cell_map.get(COL_PESO_NET), sum_peso_net,
                                 format_eu_number(sum_peso_net))
            total_un = sum(int(item.get("quantity") or 0) for item in items)
            update_formula_cache(cell_map.get(COL_UN), total_un, str(total_un))
            continue

        # --- DATA ROWS ---
        # Clear source cells for this row
        for cc in CLEAR_COLS:
            clear_cell(cell_map.get(cc))

        # Col R=FACTURA  S=CALCULO  T=DIFF  U=nota ajuste
        # Rows 2/3: PESO BRUT (pre / post ajuste)
        # Rows 4/5: PESO NET  (pre / post ajuste)
        _adj_code = invoice_data.get("_adj_item_code", "")
        _adj_diff = invoice_data.get("_adj_net_diff", 0)
        if row_idx == 1:
            # PESO BRUT - pre ajuste (raw PDF-derived calc)
            _raw_br  = invoice_data.get("_adj_br_raw") or round(
                sum(item.get("partial_weight") or 0 for item in items) * brut_net_ratio, 2)
            _diff_br = round(_raw_br - peso_brut, 2)
            set_numeric_value(cell_map.get(COL_PESO_VAL), peso_brut, format_eu_number(peso_brut))
            set_numeric_value(cell_map.get(COL_CALCULO),  _raw_br,   format_eu_number(_raw_br))
            set_numeric_value(cell_map.get(COL_DIFF),     _diff_br,  format_eu_number(_diff_br))
        if row_idx == 2:
            # PESO BRUT - post ajuste (after NET correction, formula recalculates)
            _post_br  = round(peso_net_total * brut_net_ratio, 2)
            set_numeric_value(cell_map.get(COL_PESO_VAL), peso_brut,  format_eu_number(peso_brut))
            set_numeric_value(cell_map.get(COL_CALCULO),  _post_br,   format_eu_number(_post_br))
            set_numeric_value(cell_map.get(COL_DIFF),     0.0,        "0,00")
            if _adj_code:
                # BRUT recalculates via formula from NET; show NET adjustment made
                _pw_b = invoice_data.get("_adj_item_pw_before", 0)
                _pw_a = invoice_data.get("_adj_item_pw_after", 0)
                set_cell_value(cell_map.get(COL_NOTE),
                               f"BRUT recalculado via formula | "
                               f"NET ajustado en {_adj_code}: {_pw_b} -> {_pw_a} ({_adj_diff:+.4f})")
        if row_idx == 3:
            # PESO NET - pre ajuste
            _raw_net  = invoice_data.get("_adj_net_raw") or round(
                sum(item.get("partial_weight") or 0 for item in items), 2)
            _diff_net = round(_raw_net - peso_net_total, 2)
            set_numeric_value(cell_map.get(COL_PESO_VAL), peso_net_total, format_eu_number(peso_net_total))
            set_numeric_value(cell_map.get(COL_CALCULO),  _raw_net,        format_eu_number(_raw_net))
            set_numeric_value(cell_map.get(COL_DIFF),     _diff_net,       format_eu_number(_diff_net))
        if row_idx == 4:
            # PESO NET - post ajuste
            _post_net = round(sum(item.get("partial_weight") or 0 for item in items), 2)
            set_numeric_value(cell_map.get(COL_PESO_VAL), peso_net_total, format_eu_number(peso_net_total))
            set_numeric_value(cell_map.get(COL_CALCULO),  _post_net,       format_eu_number(_post_net))
            set_numeric_value(cell_map.get(COL_DIFF),     0.0,             "0,00")
            if _adj_code:
                _pw_b = invoice_data.get("_adj_item_pw_before", 0)
                _pw_a = invoice_data.get("_adj_item_pw_after", 0)
                set_cell_value(cell_map.get(COL_NOTE),
                               f"ajuste en {_adj_code}: {_pw_b} -> {_pw_a} ({_adj_diff:+.4f})")

        if item_idx < len(items):
            item = items[item_idx]

            # CODE (col G)
            code = item.get("code", "")
            code_display = code.lstrip("0") or code
            try:
                set_numeric_value(cell_map.get(COL_CODE_G),
                                  int(code_display), code_display)
            except (ValueError, TypeError):
                set_cell_value(cell_map.get(COL_CODE_G), code_display)

            # VALOR DUA (col J)
            subtotal = item.get("subtotal_price")
            if subtotal is not None:
                set_numeric_value(cell_map.get(COL_VALOR_DUA), subtotal)

            # OPR Material Cost (col K)
            opr = item.get("opr_material")
            if opr is not None:
                set_numeric_value(cell_map.get(COL_OPR_MAT), opr)

            # PALETS (col L)
            pallets = item.get("project_pallets")
            if pallets is not None:
                set_numeric_value(cell_map.get(COL_PALETS), pallets, str(pallets))

            # PESO NET (col O) - stored as integer to match template style
            partial_weight = item.get("partial_weight")
            if partial_weight is not None:
                pw_val = round(partial_weight, 2)
                set_numeric_value(cell_map.get(COL_PESO_NET), pw_val,
                                  format_eu_number(pw_val))

            # UN / quantity (col P)
            quantity = item.get("quantity")
            if quantity is not None:
                set_numeric_value(cell_map.get(COL_UN), int(quantity), str(int(quantity)))

            # ----------------------------------------------------------------
            # Update cached values of FORMULA cells so pivot tables work
            # immediately without requiring LibreOffice to recalculate first.
            #
            # V.E (col M, index 12) = J + K
            # PESO BR (col N, index 13) = (O * R2) / R4
            #
            # We do NOT remove the formula - update_formula_cache preserves it.
            # LibreOffice will still recalculate on open, but the cached value
            # will already be correct so pivot tables built without recalc work.
            # ----------------------------------------------------------------
            ve_val = (subtotal or 0) + (opr or 0)
            update_formula_cache(cell_map.get(COL_VE), ve_val)

            pb_val = (partial_weight or 0) * brut_net_ratio
            update_formula_cache(cell_map.get(COL_PESO_BR), pb_val,
                                 str(round(pb_val)))

            item_idx += 1
        else:
            # Empty row: formula cells keep their cached 0 (correct for empty O)
            # No action needed - clear_cell above cleared the data cells
            pass

    # --- Generate Resumen summary sheet ---
    add_summary_sheet(doc, items, invoice_data, code_mapping, brut_net_ratio)

    doc.save(str(output_path))


# ---------------------------------------------------------------------------
# Top-level invoice processor
# ---------------------------------------------------------------------------

def process_invoice(pdf_path: Path, template_path: Path, output_dir: Path) -> Path:
    print(f"Processing: {pdf_path.name}")
    text = read_pdf_text(pdf_path)
    fmt = detect_number_format(text)
    print(f"  Number format: {fmt}")
    invoice_data = extract_invoice_metadata(text, fmt)
    items = extract_line_items(text, fmt)
    invoice_no = invoice_data.get("invoice_no", "UNKNOWN")

    # --- Summary & difference check vs PDF (before auto-adjustment) ---
    calc_valor = sum(item.get("subtotal_price") or 0 for item in items)
    calc_opr   = sum(item.get("opr_material") or 0 for item in items)
    calc_pnet  = sum(item.get("partial_weight") or 0 for item in items)
    pdf_pbrut  = invoice_data.get("peso_brut") or 0
    pdf_pnet   = invoice_data.get("peso_net") or 0
    ratio      = pdf_pbrut / pdf_pnet if pdf_pnet > 0 else 1.0
    calc_pbrut = round(calc_pnet * ratio, 2)

    pdf_valor = invoice_data.get("total_invoice") or 0
    pdf_opr   = invoice_data.get("total_opr") or 0

    print(f"  Invoice: {invoice_no}, Items: {len(items)}, Pallets: {invoice_data.get('pallets')}")

    diffs = []
    if pdf_valor and abs(pdf_valor - calc_valor) > 0.01:
        diffs.append(f"  Valor DUA : PDF {pdf_valor:>12.2f}  CALC {calc_valor:>12.2f}  DIFF {calc_valor-pdf_valor:+.2f}")
    if pdf_opr and abs(pdf_opr - calc_opr) > 0.01:
        diffs.append(f"  OPR 7009  : PDF {pdf_opr:>12.2f}  CALC {calc_opr:>12.2f}  DIFF {calc_opr-pdf_opr:+.2f}")
    if pdf_pnet and abs(pdf_pnet - calc_pnet) > 0.01:
        diffs.append(f"  Peso Net  : PDF {pdf_pnet:>12.2f}  CALC {calc_pnet:>12.2f}  DIFF {calc_pnet-pdf_pnet:+.2f}  [AUTO-ajustado]")
    if pdf_pbrut and abs(pdf_pbrut - calc_pbrut) > 0.01:
        diffs.append(f"  Peso Brut : PDF {pdf_pbrut:>12.2f}  CALC {calc_pbrut:>12.2f}  DIFF {calc_pbrut-pdf_pbrut:+.2f}  [recalc via formula]")

    if diffs:
        warn_msg = "AVISO - Diferencias pre-ajuste (se corrigen automaticamente):\n" + "\n".join(diffs)
        print(f"  {warn_msg}")
        invoice_data["_warnings"] = warn_msg
    else:
        print(f"  OK - Totales coinciden con PDF")

    output_file = output_dir / f"COMPLETADO_{invoice_no}.ods"
    update_ods_template(template_path, items, invoice_data, output_file)
    print(f"  [OK] Created: {output_file.name}")
    return output_file


def main(argv: List[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Lear Rabat Invoice Extractor - Unified Architecture"
    )
    parser.add_argument("pdfs", nargs="+", help="PDF invoice files to process")
    parser.add_argument("-o", "--output", required=True, help="Output directory")
    parser.add_argument(
        "-t", "--template",
        help="ODS template file (default: COMPLETADO_TEMPLATE.ods in same dir as script)"
    )
    parser.add_argument("--version", action="version",
                        version=f"%(prog)s {SCRIPT_VERSION}")

    args = parser.parse_args(argv)

    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.template:
        template_path = Path(args.template).resolve()
    else:
        template_path = Path(__file__).parent / "COMPLETADO_TEMPLATE.ods"

    if not template_path.exists():
        print(f"ERROR: Template not found: {template_path}", file=sys.stderr)
        return 2

    processed_files = []
    for pdf_file in args.pdfs:
        pdf_path = Path(pdf_file).resolve()
        if not pdf_path.exists():
            print(f"ERROR: File not found: {pdf_path}", file=sys.stderr)
            return 2
        try:
            output_file = process_invoice(pdf_path, template_path, output_dir)
            processed_files.append(output_file)
        except Exception as e:
            print(f"ERROR processing {pdf_path.name}: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            return 3

    print(f"\n[OK] Successfully processed {len(processed_files)} invoice(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))