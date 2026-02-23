#!/usr/bin/env python3
"""
Lear Rabat Invoice Extractor - Unified Architecture Version
Extracts data from Lear Rabat invoices and generates ODS files with DUA data.

This extractor processes Lear Automotive Morocco invoices and populates
a template ODS file with customs clearance data.

Usage (unified processor interface):
  extract_lear_rabat.py <pdf1> <pdf2> ... -o <output_dir> [-t <template.ods>]
"""

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional

SCRIPT_VERSION = "2026-02-21.v9"

try:
    import pdfplumber
except ImportError:
    print("ERROR: missing dependency. Install with: pip install pdfplumber", file=sys.stderr)
    sys.exit(1)

try:
    from odf import opendocument, table, text as odftext
    from odf.opendocument import load as load_ods
except ImportError:
    print("ERROR: missing dependency. Install with: pip install odfpy", file=sys.stderr)
    sys.exit(1)


def read_pdf_text(pdf_path: Path) -> str:
    """Extract text from all pages."""
    parts = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                parts.append(text)
    return "\n".join(parts)


def detect_number_format(text: str) -> str:
    """Detect whether the invoice uses EU or US number format.
    
    EU format: spaces for thousands, comma for decimal  (e.g. "1 536,00" or "13 210,09")
    US format: commas for thousands, dot for decimal    (e.g. "1,536.00" or "13,210.09")
    """
    us_pattern = re.findall(r'\d{1,3},\d{3}\.\d{2}', text)
    eu_pattern = re.findall(r'\d{1,3}\s\d{3},\d{2}', text)
    
    if len(us_pattern) > len(eu_pattern):
        return "US"
    return "EU"


def parse_number(s: str, fmt: str = "EU") -> Optional[float]:
    """Parse a number string, handling both EU and US formats."""
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
    """Extract invoice number from header."""
    lines = text.splitlines()
    for line in lines[:20]:
        m = re.match(r'^(\d{8})\s', line.strip())
        if m:
            return m.group(1)
        if re.match(r'^\d{8}$', line.strip()):
            return line.strip()
    return ""


def extract_invoice_metadata(text: str, fmt: str) -> Dict:
    """Extract invoice header data (weights, pallets, totals)."""
    lines = text.splitlines()
    
    invoice_no = extract_invoice_number(text)
    total_invoice = None
    peso_brut = None
    peso_net = None
    pallets = None
    total_opr = None
    
    for i, line in enumerate(lines):
        # Extract total: "Total: 110 134,51" or "Total: 94,200.01"
        m = re.search(r'Total:\s+([\d\s,\.]+)', line)
        if m:
            total_invoice = parse_number(m.group(1), fmt)
        
        # Extract weights
        if "Net:" in line and "Gross:" in line:
            m = re.search(r'Net:\s*([\d\s]+)\s*K\s+.*?Gross:\s*([\d\s]+)\s*K', line)
            if m:
                peso_net = parse_number(m.group(1), "EU")
                peso_brut = parse_number(m.group(2), "EU")
        
        # Extract pallet count from delivery conditions line
        if i < 15 and ("FCA" in line or "EXW" in line or "DAP" in line):
            m = re.search(r'\b(\d{1,3})\s*$', line.strip())
            if m:
                pallets = int(m.group(1))
    
    # Extract total OPR from end of document
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
    """Extract line items with project grouping.
    
    Key insight: Items appear BEFORE their project header.
    The project header appears after the last item in that project,
    followed by OPR Material Cost for that project.
    """
    lines = text.splitlines()
    
    # First pass: collect all FG items
    raw_items = []
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Detect FG lines - flexible regex for both number formats
        # HS code is optional (some items don't have it)
        fg_match = re.match(
            r'(\d+)\s+FG\s+(\S+)\s+ESOPO\d+-\s*(\d+)I?\s+(?:(\d{7,})\s+)?([\d,\.\s]+)',
            line
        )
        if fg_match:
            line_no = fg_match.group(1)
            part_number = fg_match.group(2)
            code = fg_match.group(3)
            hs_code = fg_match.group(4) or ""  # may be None if no HS code
            qty_str = fg_match.group(5).strip()
            quantity = parse_number(qty_str, fmt)
            
            # Search forward for Subtotal FG and Partial Weight
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
                
                # Stop at next FG line
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
    
    # Second pass: find project headers and OPR costs
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
    
    # Assign projects to items: each item belongs to the NEXT project header
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
    
    # Mark last item per project (only last gets OPR/pallets)
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
    
    # Clean up
    for item in raw_items:
        item.pop("line_idx", None)
    
    return raw_items


def get_cell_value(cell):
    """Get text value from ODS cell."""
    for p in cell.getElementsByType(odftext.P):
        return str(p).strip()
    return ""


def set_cell_value(cell, value):
    """Set text value in ODS cell, removing any formula."""
    if cell is None:
        return
    # Remove formula attribute if present (prevents formula from overriding our value)
    try:
        cell.removeAttribute("formula")
    except Exception:
        pass
    # Remove cached value attributes
    try:
        cell.removeAttribute("valuetype")
    except Exception:
        pass
    try:
        cell.removeAttribute("value")
    except Exception:
        pass
    # Clear children
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except (ValueError, Exception):
                pass
    except Exception:
        pass
    p = odftext.P()
    p.addText(str(value))
    cell.appendChild(p)


def set_numeric_value(cell, numeric_value, display_text=None):
    """Set a proper ODS numeric value with office:value-type and office:value.
    
    This ensures formulas like SUM() can read the cell correctly.
    """
    if cell is None:
        return
    
    float_val = float(numeric_value)
    
    # Remove formula attribute if present
    try:
        cell.removeAttribute("formula")
    except Exception:
        pass
    
    # Set ODS numeric attributes
    cell.setAttribute("valuetype", "float")
    cell.setAttribute("value", str(float_val))
    
    # Clear children and set display text
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except (ValueError, Exception):
                pass
    except Exception:
        pass
    
    if display_text is None:
        display_text = format_eu_number(float_val)
    
    p = odftext.P()
    p.addText(str(display_text))
    cell.appendChild(p)


def clear_cell(cell):
    """Clear all content from an ODS cell, leaving it empty.
    
    Removes formula, value attributes, and all child nodes.
    """
    if cell is None:
        return
    # Remove formula
    try:
        cell.removeAttribute("formula")
    except Exception:
        pass
    # Remove value attributes
    try:
        cell.removeAttribute("valuetype")
    except Exception:
        pass
    try:
        cell.removeAttribute("value")
    except Exception:
        pass
    # Clear children
    try:
        for child in list(cell.childNodes):
            try:
                cell.removeChild(child)
            except (ValueError, Exception):
                pass
    except Exception:
        pass


def format_eu_number(value, decimals=2):
    """Format a number in EU style: comma as decimal separator, strip trailing zeros."""
    if value is None:
        return ""
    formatted = f'{value:.{decimals}f}'.replace(".", ",")
    # Strip trailing zeros after comma, and strip comma if no decimals left
    if "," in formatted:
        formatted = formatted.rstrip("0").rstrip(",")
    return formatted


def format_peso_br(value):
    """Format PESO BR with dot as thousands separator (integer)."""
    if value is None:
        return ""
    int_val = round(value)
    if int_val >= 1000:
        # Format with dot as thousands separator
        s = str(int_val)
        parts = []
        while s:
            parts.append(s[-3:])
            s = s[:-3]
        return ".".join(reversed(parts))
    return str(int_val)


def build_cell_map(row, writable_cols=None):
    """Build a mapping from column index to cell, handling repeated cells.
    
    For columns listed in writable_cols, if the cell is repeated,
    split it so each column has its own independent cell.
    Returns dict of {col_index: cell}.
    """
    if writable_cols is None:
        writable_cols = set()
    
    cells = list(row.getElementsByType(table.TableCell))
    cell_map = {}
    col = 0
    
    for cell in cells:
        rep = cell.getAttribute("numbercolumnsrepeated")
        rep = int(rep) if rep else 1
        
        # Check if any writable column falls within this repeated range
        needs_split = any((col + r) in writable_cols for r in range(rep))
        
        if needs_split and rep > 1:
            # Remove the repeat attribute
            try:
                cell.removeAttribute("numbercolumnsrepeated")
            except Exception:
                pass
            
            cell_map[col] = cell
            
            # Create individual cells for the remaining repetitions
            prev = cell
            for r in range(1, rep):
                import copy
                new_cell = table.TableCell()
                # Copy the text content
                val = get_cell_value(cell)
                if val:
                    p = odftext.P()
                    p.addText(val)
                    new_cell.appendChild(p)
                # Insert after previous cell
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
    """Get the value of column B (index 1)."""
    cell = cell_map.get(1)
    if cell:
        return get_cell_value(cell)
    return ""


def update_ods_template(template_path: Path, items: List[Dict], invoice_data: Dict, output_path: Path):
    """Update ODS template with extracted invoice data.
    
    Strategy:
    - Write NUMERIC values to SOURCE cells (c6,c9,c10,c11,c14,c15)
    - Leave FORMULA cells untouched (c7,c8 VLOOKUP; c12 V.E=J+K; c13 PESO_BR=O*R2/R4)
    - Calculate ALL totals ourselves and write them as values to totals row
    - R2 (PESO BRUT) and R4 (PESO NET) = copies of N52 and O52 totals
    """
    
    doc = load_ods(str(template_path))
    sheet = doc.spreadsheet.getElementsByType(table.Table)[0]
    rows = sheet.getElementsByType(table.TableRow)
    
    COL_CODE_G = 6      # G - Code (source for VLOOKUP)
    COL_VALOR_DUA = 9   # J - Valor DUA
    COL_OPR_MAT = 10    # K - 7009 (OPR Material Cost)
    COL_PALETS = 11     # L - Palets
    # COL 12 = M - V.E  → FORMULA: J+K (don't touch)
    COL_PESO_BR = 13    # N - Peso BR → FORMULA per row, but we WRITE the total
    COL_PESO_NET = 14   # O - Peso Net
    COL_UN = 15         # P - UN (quantity)
    COL_PESO_VAL = 17   # R - Peso Brut/Net summary values
    
    # Columns we need cell access to
    WRITABLE_COLS = {COL_CODE_G, COL_VALOR_DUA, COL_OPR_MAT, COL_PALETS,
                     COL_PESO_BR, COL_PESO_NET, COL_UN, COL_PESO_VAL}
    
    # Source columns to clear on data rows (NOT formula cols c12, c13)
    CLEAR_COLS = {COL_CODE_G, COL_VALOR_DUA, COL_OPR_MAT, COL_PALETS,
                  COL_PESO_NET, COL_UN}
    
    # --- Pre-calculate totals from items ---
    peso_brut = invoice_data.get("peso_brut") or 0
    peso_net_total = invoice_data.get("peso_net") or 0
    brut_net_ratio = peso_brut / peso_net_total if peso_net_total > 0 else 1.0
    
    total_valor = 0.0
    total_opr = 0.0
    total_pallets = 0
    total_peso_net = 0.0
    total_peso_br = 0.0
    
    for item in items:
        subtotal = item.get("subtotal_price") or 0
        opr = item.get("opr_material") or 0
        pals = item.get("project_pallets") or 0
        pw = item.get("partial_weight") or 0
        total_valor += subtotal
        total_opr += opr
        total_pallets += pals
        total_peso_net += pw
        total_peso_br += pw * brut_net_ratio
    
    item_idx = 0
    
    for row_idx, row in enumerate(rows):
        if row_idx == 0:
            continue
        
        cell_map = build_cell_map(row, WRITABLE_COLS)
        col_b = get_col_b_value(cell_map)
        
        # --- TOTALS ROW (520) ---
        if col_b == "520":
            # Clear all totals cells first (handles dirty templates)
            for cc in CLEAR_COLS:
                clear_cell(cell_map.get(cc))
            clear_cell(cell_map.get(COL_PESO_BR))
            
            # Write totals — prefer PDF-extracted values (exact), fallback to calculated
            final_valor = invoice_data.get("total_invoice") or total_valor
            final_opr = invoice_data.get("total_opr") or total_opr
            final_pallets = invoice_data.get("pallets") or total_pallets
            set_numeric_value(cell_map.get(COL_VALOR_DUA), final_valor,
                              f'{final_valor:.2f}'.replace(".", ","))
            set_numeric_value(cell_map.get(COL_OPR_MAT), final_opr,
                              f'{final_opr:.2f}'.replace(".", ","))
            set_numeric_value(cell_map.get(COL_PALETS), final_pallets, str(final_pallets))
            # PESO BR and PESO NET totals: use PDF values (exact) for consistency with R2/R4
            peso_br_total = int(invoice_data.get("peso_brut") or round(total_peso_br))
            peso_net_tot = int(invoice_data.get("peso_net") or round(total_peso_net))
            set_numeric_value(cell_map.get(COL_PESO_BR), peso_br_total, str(peso_br_total))
            set_numeric_value(cell_map.get(COL_PESO_NET), peso_net_tot, str(peso_net_tot))
            continue
        
        # --- DATA ROWS ---
        # Clear source cells
        for cc in CLEAR_COLS:
            clear_cell(cell_map.get(cc))
        
        # Write PESO BRUT/NET summary (R2, R4)
        # Use invoice PDF values directly (more accurate than summing items)
        peso_brut_pdf = invoice_data.get("peso_brut")
        peso_net_pdf = invoice_data.get("peso_net")
        if row_idx == 1 and peso_brut_pdf is not None:
            set_numeric_value(cell_map.get(COL_PESO_VAL),
                              int(peso_brut_pdf), str(int(peso_brut_pdf)))
        if row_idx == 3 and peso_net_pdf is not None:
            set_numeric_value(cell_map.get(COL_PESO_VAL),
                              int(peso_net_pdf), str(int(peso_net_pdf)))
        
        # Place next item sequentially
        if item_idx < len(items):
            item = items[item_idx]
            
            # CODE (column G) - numeric for VLOOKUP
            code = item.get("code", "")
            code_display = code.lstrip("0") or code
            try:
                set_numeric_value(cell_map.get(COL_CODE_G), int(code_display), code_display)
            except (ValueError, TypeError):
                set_cell_value(cell_map.get(COL_CODE_G), code_display)
            
            # VALOR DUA (column J)
            subtotal = item.get("subtotal_price")
            if subtotal is not None:
                set_numeric_value(cell_map.get(COL_VALOR_DUA), subtotal)
            
            # OPR Material Cost (column K) - only last item in project
            opr = item.get("opr_material")
            if opr is not None:
                set_numeric_value(cell_map.get(COL_OPR_MAT), opr)
            
            # PALETS (column L) - only last item in project
            pallets = item.get("project_pallets")
            if pallets is not None:
                set_numeric_value(cell_map.get(COL_PALETS), pallets, str(pallets))
            
            # PESO NET (column O)
            partial_weight = item.get("partial_weight")
            if partial_weight is not None:
                set_numeric_value(cell_map.get(COL_PESO_NET), partial_weight,
                                  f'{partial_weight:.2f}'.replace(".", ","))
            
            # UN / quantity (column P)
            quantity = item.get("quantity")
            if quantity is not None:
                set_numeric_value(cell_map.get(COL_UN), int(quantity), str(int(quantity)))
            
            item_idx += 1
    
    doc.save(str(output_path))


def process_invoice(pdf_path: Path, template_path: Path, output_dir: Path) -> Path:
    """Process a single invoice PDF and generate ODS output."""
    
    print(f"Processing: {pdf_path.name}")
    
    text = read_pdf_text(pdf_path)
    fmt = detect_number_format(text)
    print(f"  Number format: {fmt}")
    
    invoice_data = extract_invoice_metadata(text, fmt)
    items = extract_line_items(text, fmt)
    
    invoice_no = invoice_data.get("invoice_no", "UNKNOWN")
    print(f"  Invoice: {invoice_no}, Items: {len(items)}, Pallets: {invoice_data.get('pallets')}")
    print(f"  Total: {invoice_data.get('total_invoice')}, OPR: {invoice_data.get('total_opr')}")
    print(f"  Peso Brut: {invoice_data.get('peso_brut')}, Peso Net: {invoice_data.get('peso_net')}")
    
    for i, item in enumerate(items):
        print(f"  #{i+1}: code={item['code']} part={item['part_number']} "
              f"qty={item.get('quantity')} subtotal={item.get('subtotal_price')} "
              f"weight={item.get('partial_weight')} project={item.get('project')} "
              f"pallets={item.get('project_pallets')} opr={item.get('opr_material')}")
    
    output_file = output_dir / f"COMPLETADO_{invoice_no}.ods"
    update_ods_template(template_path, items, invoice_data, output_file)
    
    print(f"  ✔ Created: {output_file.name}")
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
    parser.add_argument("--version", action="version", version=f"%(prog)s {SCRIPT_VERSION}")
    
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
    
    print(f"\n✔ Successfully processed {len(processed_files)} invoice(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
