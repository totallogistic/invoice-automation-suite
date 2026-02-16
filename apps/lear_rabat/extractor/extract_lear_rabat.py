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
from decimal import Decimal

SCRIPT_VERSION = "2026-02-16.v1"

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


def parse_number(s: str) -> Optional[float]:
    """Parse a number string, handling spaces, commas as decimals."""
    if not s or s.strip() == "":
        return None
    # Remove spaces
    s = s.replace(" ", "").strip()
    # Handle comma as decimal separator (European format)
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def extract_invoice_number(text: str) -> str:
    """Extract invoice number from header."""
    lines = text.splitlines()
    for line in lines[:20]:
        # Match pattern like "00007297" on line after "INVOICE"
        if re.match(r'^\d{8}$', line.strip()):
            return line.strip()
    return ""


def extract_invoice_metadata(text: str) -> Dict:
    """Extract invoice header data (weights, pallets, totals)."""
    lines = text.splitlines()
    
    invoice_no = extract_invoice_number(text)
    total_invoice = None
    peso_brut = None
    peso_net = None
    pallets = None
    
    for i, line in enumerate(lines):
        # Extract total from "Total: 110 134,51"
        if "TOTAL INVOICE" in line or "Total:" in line:
            for j in range(i, min(i + 10, len(lines))):
                m = re.search(r'Total:\s+([\d\s,\.]+)', lines[j])
                if m:
                    total_invoice = parse_number(m.group(1))
                    break
        
        # Extract weights "Net: 8827 K Brut - Bruto - Gross: 10696 K"
        if "Net:" in line and "Gross:" in line:
            m = re.search(r'Net:\s*([\d\s]+)\s*K\s+.*?Gross:\s*([\d\s]+)\s*K', line)
            if m:
                peso_net = parse_number(m.group(1))
                peso_brut = parse_number(m.group(2))
        
        # Extract pallet count from "Colls - Bultos - Packs" section
        if i < 30 and ("Colls" in line or "Packs" in line):
            # Next line or same line might have the number
            for check_line in [line, lines[i+1] if i+1 < len(lines) else ""]:
                m = re.search(r'\b(\d{1,3})\s*$', check_line.strip())
                if m:
                    pallets = int(m.group(1))
                    break
    
    return {
        "invoice_no": invoice_no,
        "total_invoice": total_invoice,
        "peso_brut": peso_brut,
        "peso_net": peso_net,
        "pallets": pallets,
    }


def extract_line_items(text: str) -> List[Dict]:
    """Extract line items with project grouping."""
    lines = text.splitlines()
    items = []
    current_project = None
    
    # Track project-level data
    project_data = {}
    
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Detect project headers: "O_BCP21 Project: 7 Plts"
        project_match = re.match(r'(O_\w+)\s+Project:\s+(\d+)\s+Plts', line)
        if project_match:
            current_project = project_match.group(1)
            project_pallets = int(project_match.group(2))
            project_data[current_project] = {"pallets": project_pallets}
            i += 1
            continue
        
        # Detect FG (Finished Goods) lines
        # Pattern: "N FG PartNumber ESOPO...- CodeI HSCode Quantity"
        fg_match = re.match(r'(\d+)\s+FG\s+(\S+)\s+ESOPO\d+-\s*(\d+)I?\s+([\d]+)\s+([\d,\.]+)', line)
        if fg_match:
            line_no = fg_match.group(1)
            part_number = fg_match.group(2)
            code = fg_match.group(3)  # Remove 'I' suffix if present
            hs_code = fg_match.group(4)
            quantity = parse_number(fg_match.group(5))
            
            # Look ahead for "Subtotal FG:" and "Partial Weight:"
            subtotal_price = None
            partial_weight = None
            
            for j in range(i + 1, min(i + 30, len(lines))):
                check_line = lines[j].strip()
                
                if "Subtotal FG:" in check_line:
                    # Extract price from end of line
                    m = re.search(r'([\d\s,\.]+)$', check_line)
                    if m:
                        subtotal_price = parse_number(m.group(1))
                
                if "Partial Weight:" in check_line:
                    # Extract weight "1 034,50 K"
                    m = re.search(r'([\d\s,\.]+)\s*K', check_line)
                    if m:
                        partial_weight = parse_number(m.group(1))
                
                # Stop at next FG or project
                if re.match(r'(\d+\s+FG|O_\w+\s+Project)', check_line):
                    break
            
            # Look for project subtotals (OPR Material Cost, etc.)
            opr_material = None
            if current_project and current_project in project_data:
                # Scan for OPR Material Cost for this project
                for j in range(i, min(i + 100, len(lines))):
                    check_line = lines[j].strip()
                    if "OPR Material Cost:" in check_line:
                        m = re.search(r'([\d\s,\.]+)$', check_line)
                        if m:
                            opr_material = parse_number(m.group(1))
                            project_data[current_project]["opr_material"] = opr_material
                        break
                    # Stop at next project
                    if re.match(r'O_\w+\s+Project:', check_line) and current_project not in check_line:
                        break
            
            item_data = {
                "line_no": line_no,
                "part_number": part_number,
                "code": code,
                "hs_code": hs_code,
                "quantity": quantity,
                "subtotal_price": subtotal_price,
                "partial_weight": partial_weight,
                "project": current_project,
            }
            
            # Add project-level data
            if current_project and current_project in project_data:
                item_data["project_pallets"] = project_data[current_project].get("pallets")
                item_data["opr_material"] = project_data[current_project].get("opr_material")
            
            items.append(item_data)
        
        i += 1
    
    return items


def get_cell_value(cell):
    """Get text value from ODS cell."""
    for p in cell.getElementsByType(odftext.P):
        return str(p).strip()
    return ""


def set_cell_value(cell, value):
    """Set text value in ODS cell."""
    # Remove existing content
    for child in cell.childNodes[:]:
        cell.removeChild(child)
    # Add new content
    p = odftext.P()
    p.addText(str(value))
    cell.appendChild(p)


def update_ods_template(template_path: Path, items: List[Dict], invoice_data: Dict, output_path: Path):
    """Update ODS template with extracted invoice data."""
    
    # Load template
    doc = load_ods(str(template_path))
    sheet = doc.spreadsheet.getElementsByType(table.Table)[0]
    rows = sheet.getElementsByType(table.TableRow)
    
    # Column mapping (0-indexed)
    COL_CODE = 6  # Column G - CODIGO (to match against)
    COL_VALOR_DUA = 9  # Column J - VALOR DUA
    COL_OPR_MAT = 10  # Column K - OPR Material Cost
    COL_PALETS = 11  # Column L - PALETS
    COL_VE = 12  # Column M - V.E (calculated)
    COL_PESO_BR = 13  # Column N - PESO BR
    COL_PESO_NET = 14  # Column O - PESO NET
    COL_UN = 15  # Column P - UN (units/quantity)
    
    # Group items by code
    items_by_code = {}
    for item in items:
        code = item["code"]
        if code not in items_by_code:
            items_by_code[code] = []
        items_by_code[code].append(item)
    
    # Update data rows
    for row_idx, row in enumerate(rows):
        if row_idx == 0:  # Skip header row
            continue
        
        cells = row.getElementsByType(table.TableCell)
        if len(cells) < 17:
            continue
        
        # Get the code from column G
        code = get_cell_value(cells[COL_CODE])
        
        if code and code in items_by_code:
            # Get first matching item for this code
            item = items_by_code[code][0]
            
            # Set VALOR DUA (subtotal price)
            if item.get("subtotal_price"):
                set_cell_value(cells[COL_VALOR_DUA], f'{item["subtotal_price"]:.2f}'.replace(".", ","))
            
            # Set OPR Material Cost
            if item.get("opr_material"):
                set_cell_value(cells[COL_OPR_MAT], f'{item["opr_material"]:.2f}'.replace(".", ","))
            
            # Set PALETS (project-level)
            if item.get("project_pallets"):
                set_cell_value(cells[COL_PALETS], str(item["project_pallets"]))
            
            # Calculate and set V.E (VALOR DUA + OPR Material)
            if item.get("subtotal_price") and item.get("opr_material"):
                ve = item["subtotal_price"] + item["opr_material"]
                set_cell_value(cells[COL_VE], f'{ve:.2f}'.replace(".", ","))
            
            # Set PESO NET (partial weight)
            if item.get("partial_weight"):
                set_cell_value(cells[COL_PESO_NET], str(item["partial_weight"]))
            
            # Set UN (quantity/units)
            if item.get("quantity"):
                set_cell_value(cells[COL_UN], str(int(item["quantity"])))
    
    # Update totals row (code 520 gets invoice-level totals)
    for row_idx, row in enumerate(rows):
        cells = row.getElementsByType(table.TableCell)
        if len(cells) < 17:
            continue
        
        code_cell = get_cell_value(cells[1])  # Column B
        
        if code_cell == "520":
            # Set invoice-level totals
            if invoice_data.get("total_invoice"):
                set_cell_value(cells[COL_VALOR_DUA], f'{invoice_data["total_invoice"]:.2f}'.replace(".", ","))
            
            if invoice_data.get("pallets"):
                set_cell_value(cells[COL_PALETS], str(invoice_data["pallets"]))
            
            if invoice_data.get("peso_brut"):
                set_cell_value(cells[COL_PESO_BR], str(int(invoice_data["peso_brut"])))
            
            if invoice_data.get("peso_net"):
                set_cell_value(cells[COL_PESO_NET], str(int(invoice_data["peso_net"])))
            
            break
    
    # Save output
    doc.save(str(output_path))


def process_invoice(pdf_path: Path, template_path: Path, output_dir: Path) -> Path:
    """Process a single invoice PDF and generate ODS output."""
    
    print(f"Processing: {pdf_path.name}")
    
    # Extract data
    text = read_pdf_text(pdf_path)
    invoice_data = extract_invoice_metadata(text)
    items = extract_line_items(text)
    
    invoice_no = invoice_data.get("invoice_no", "UNKNOWN")
    print(f"  Invoice: {invoice_no}, Items: {len(items)}")
    
    # Generate output filename
    output_file = output_dir / f"COMPLETADO_{invoice_no}.ods"
    
    # Update template
    update_ods_template(template_path, items, invoice_data, output_file)
    
    print(f"  ✓ Created: {output_file.name}")
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
    
    # Resolve paths
    output_dir = Path(args.output).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Find template
    if args.template:
        template_path = Path(args.template).resolve()
    else:
        # Default: look for template in same directory as script
        template_path = Path(__file__).parent / "COMPLETADO_TEMPLATE.ods"
    
    if not template_path.exists():
        print(f"ERROR: Template not found: {template_path}", file=sys.stderr)
        print("Please provide template with -t flag or place COMPLETADO_TEMPLATE.ods in extractor directory", file=sys.stderr)
        return 2
    
    # Process each PDF
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
    
    print(f"\n✓ Successfully processed {len(processed_files)} invoice(s)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
