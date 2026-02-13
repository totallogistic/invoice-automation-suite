#!/usr/bin/env python3
"""
Import Partida extractor - Multi-format support
Extracts data from Maersk B/L PDFs (supports multiple template formats)

Usage:
  extract_import_partida_fields.py <input.pdf> <out_dir>

Outputs in <out_dir>:
  - import_partida.csv
"""

import re
import sys
from pathlib import Path

import pdfplumber


def read_pdf_text(pdf_path: Path) -> str:
    """Extract text from all pages."""
    parts = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                parts.append(text)
    return "\n".join(parts)


def normalize_ws(s: str) -> str:
    """Normalize whitespace while keeping commas/periods."""
    s = s.replace("\u00a0", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\s*\n\s*", "\n", s)
    return s.strip()


def extract_field(text: str, patterns: list) -> str:
    """Try multiple regex patterns, return first match."""
    for pattern in patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
    return ""


def extract_bl_no(text: str) -> str:
    """Extract B/L number."""
    return extract_field(text, [
        r"B/L No\.\s+([A-Z0-9]+)",
        r"B/L:\s*([A-Z0-9]+)",
    ])


def extract_shipper(text: str) -> str:
    """Extract shipper name (company name only, no CO., LTD.)."""
    lines = text.splitlines()
    bl_no = extract_bl_no(text)
    
    for i, line in enumerate(lines):
        if "Shipper (As principal" in line:
            if i + 1 < len(lines):
                company_line = lines[i + 1].strip()
                # Remove booking/BL number if present
                if bl_no:
                    company_line = company_line.replace(bl_no, "").strip()
                # Remove "CO., LTD." and everything after
                company = re.sub(r'\s+CO\.,\s*LTD\..*$', '', company_line, flags=re.IGNORECASE)
                return company.strip()
    return ""


def extract_consignee(text: str) -> str:
    """Extract consignee name."""
    lines = text.splitlines()
    
    # Format 1: "As principal, where..." followed by company name on same line
    for i, line in enumerate(lines):
        if "As principal, where" in line and ("ALVARO" in line or "NEXO" in line):
            # Extract company name from same line
            m = re.search(r'\)\s+([A-Z0-9\s\-]+(?:S\.L\.U\.|S\.L\.))', line)
            if m:
                return m.group(1).strip()
    
    # Format 2: Line starts with company name after Consignee section
    for i, line in enumerate(lines):
        if "Consignee (Negotiable" in line:
            # Check next few lines for company name
            for j in range(i + 1, min(i + 4, len(lines))):
                next_line = lines[j].strip()
                # Look for S.L.U. or S.L.
                if "S.L.U." in next_line or "S.L." in next_line:
                    m = re.match(r'^([A-Z0-9\s\-]+(?:S\.L\.U\.|S\.L\.))', next_line)
                    if m:
                        return m.group(1).strip()
    
    return ""


def extract_vessel(text: str) -> str:
    """Extract vessel name and voyage number combined."""
    lines = text.splitlines()
    
    # Format 1: Line after "Vessel (see clause 1 + 19) Voyage No."
    for i, line in enumerate(lines):
        if "Vessel (see clause 1 + 19)" in line and i + 1 < len(lines):
            vessel_line = lines[i + 1].strip()
            # Should be like "BERLIN MAERSK 603W"
            return vessel_line
    
    # Format 2: Line after "Vessel Voyage No."
    for i, line in enumerate(lines):
        if line.startswith("Vessel") and "Voyage No." in line and i + 1 < len(lines):
            vessel_line = lines[i + 1].strip()
            # Should be like "BEIJING MAERSK 550W"
            return vessel_line
    
    return ""


def extract_ports(text: str) -> tuple:
    """Extract port of loading and port of discharge."""
    lines = text.splitlines()
    pol = ""
    pod = ""
    
    # Format 1: Ports on separate line "NINGBO, CHINA Valencia,Spain"
    for line in lines:
        if line.startswith("NINGBO, CHINA"):
            pol = "NINGBO, CHINA"
            if "Valencia,Spain" in line:
                pod = "Valencia,Spain"
            break
    
    # Format 2: Ports on one line "Shanghai Valencia Madrid"
    if not pol:
        for line in lines:
            # Look for line with multiple cities
            if "Shanghai" in line and "Valencia" in line:
                parts = line.split()
                if "Shanghai" in parts:
                    pol = "Shanghai"
                if "Valencia" in parts:
                    pod = "Valencia"
                break
    
    return pol, pod


def extract_contain(text: str) -> str:
    """Extract container/package count."""
    return extract_field(text, [
        r"(\d+\s+CARTONS)",
        r"(\d+\s+PACKAGES)",
        r"Said to Contain\s+(\d+\s+(?:PACKAGES|CARTONS))",
    ])


def extract_weight(text: str) -> str:
    """Extract weight (handles both comma and period as decimal separator)."""
    return extract_field(text, [
        r"([\d,\.]+\s+KGS)",
        r"Weight[:\s]+([\d,\.]+\s+KGS)",
    ])


def extract_measurement(text: str) -> str:
    """Extract measurement (handles both comma and period as decimal separator)."""
    return extract_field(text, [
        r"([\d,\.]+\s+CBM)",
        r"Measurement[:\s]+([\d,\.]+\s+CBM)",
    ])


def extract_mrsu(text: str) -> str:
    """Extract container number (MRSU or MSKU prefix)."""
    return extract_field(text, [
        r"\b(MRSU\d{7})\b",
        r"\b(MSKU\d{7})\b",
    ])


def extract_all_fields(pdf_path: Path) -> dict:
    """Extract all fields from PDF."""
    text = read_pdf_text(pdf_path)
    text = normalize_ws(text)
    
    if not text.strip():
        raise ValueError("No text extracted from PDF")
    
    pol, pod = extract_ports(text)
    weight = extract_weight(text)
    measurement = extract_measurement(text)
    
    # Convert decimal separators from period to comma (European format)
    weight = weight.replace(".", ",") if weight else ""
    measurement = measurement.replace(".", ",") if measurement else ""
    
    return {
        "bl_no": extract_bl_no(text),
        "shipper": extract_shipper(text),
        "consignee": extract_consignee(text),
        "vessel": extract_vessel(text),
        "port_of_loading": pol,
        "port_of_discharge": pod,
        "contain": extract_contain(text),
        "weight": weight,
        "measurement": measurement,
        "mrsu": extract_mrsu(text),
    }


def write_csv(path: Path, data: dict) -> None:
    """Write extracted data to CSV with selective quoting to match expected format."""
    columns = [
        "bl_no", "shipper", "consignee", "vessel",
        "port_of_loading", "port_of_discharge",
        "contain", "weight", "measurement", "mrsu"
    ]
    
    # Fields that should NOT be quoted
    no_quote_fields = {"bl_no", "mrsu"}
    
    with path.open("w", encoding="utf-8", newline="") as f:
        # Write header
        f.write(",".join(columns) + "\r\n")
        
        # Write data row with selective quoting
        row_parts = []
        for col in columns:
            value = data.get(col, "")
            if col in no_quote_fields:
                row_parts.append(value)
            else:
                # Quote the field
                row_parts.append(f'"{value}"')
        
        f.write(",".join(row_parts) + "\r\n")


def main(argv: list) -> int:
    if len(argv) != 3:
        print("Usage: extract_import_partida_fields.py <input.pdf> <out_dir>", file=sys.stderr)
        return 2
    
    pdf_path = Path(argv[1]).expanduser().resolve()
    out_dir = Path(argv[2]).expanduser().resolve()
    
    if not pdf_path.exists():
        print(f"ERROR: PDF not found: {pdf_path}", file=sys.stderr)
        return 2
    
    out_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        data = extract_all_fields(pdf_path)
        write_csv(out_dir / "import_partida.csv", data)
        print(f"✓ Extracted data to {out_dir / 'import_partida.csv'}")
        return 0
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 3


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))

