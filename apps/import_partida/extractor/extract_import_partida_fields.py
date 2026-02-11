#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, Optional

import pdfplumber
from openpyxl import Workbook


def die(msg: str, rc: int = 2) -> int:
    print(msg, file=sys.stderr)
    return rc


def norm(s: str) -> str:
    s = s.replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def read_pdf_text(pdf_path: Path) -> str:
    chunks = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return norm("\n".join(chunks))


def first_match(text: str, pattern: str, flags: int = re.IGNORECASE | re.MULTILINE) -> Optional[str]:
    m = re.search(pattern, text, flags)
    if not m:
        return None
    return (m.group(1) or "").strip()


def clean_one_line(s: str) -> str:
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s.strip(" :")  # NO quitamos '.' aquí a propósito


def strip_trailing_punct(s: str) -> str:
    # para Port of Discharge: quitamos ':' y '.' al final
    return re.sub(r"[.:]+\s*$", "", s.strip())


def line_value_after_label(text: str, labels: list[str]) -> Optional[str]:
    for lab in labels:
        lab_re = re.escape(lab)

        # mismo renglón: "LABEL: value" o "LABEL value"
        m = re.search(rf"(?im)^\s*{lab_re}\s*[:\-]?\s*(.+?)\s*$", text)
        if m:
            val = clean_one_line(m.group(1))
            if val and val.lower() != lab.lower():
                return val

        # valor en la línea siguiente
        m2 = re.search(rf"(?im)^\s*{lab_re}\s*$\n\s*(.+?)\s*$", text)
        if m2:
            val = clean_one_line(m2.group(1))
            if val:
                return val

    return None


def block_after_label(text: str, label: str, stop_labels: list[str], first_line_only: bool = False) -> Optional[str]:
    """
    Extrae bloque tras etiqueta hasta la siguiente etiqueta de parada.
    Si first_line_only=True, devuelve SOLO la primera línea no vacía del bloque.
    """
    stop_re = "|".join([re.escape(x) for x in stop_labels])

    pat = rf"(?is){re.escape(label)}\s*(?:\([^)]*\))?\s*:?\s*\n(.*?)(?:\n(?:{stop_re})\b)"
    m = re.search(pat, text)
    if not m:
        return None

    block_raw = m.group(1)

    # Normaliza líneas (sin destruir estructura antes de sacar primera línea)
    lines = [ln.strip() for ln in block_raw.splitlines()]
    lines = [ln for ln in lines if ln]  # quita vacías
    if not lines:
        return None

    if first_line_only:
        return clean_one_line(lines[0])

    # bloque completo en una línea
    block = " ".join(lines)
    return clean_one_line(block)


def extract_fields(text: str) -> Dict[str, Any]:
    # Try to extract B/L No from "B/L No:" label first
    bl_no = first_match(text, r"(?im)B/L\s+No\.?\s*:?\s*([A-Z0-9\-]+)") or ""
    # Fallback to NPOS pattern if not found
    if not bl_no:
        bl_no = first_match(text, r"\b(NPOS\d{4,})\b") or ""

    shipper = block_after_label(
        text,
        "Shipper",
        stop_labels=["Consignee", "Notify party", "Booking No.", "Export references", "Voyage No.", "Vessel"],
        first_line_only=False,
    ) or ""

    # 👇 Consignee: SOLO razón social (primera línea)
    consignee = block_after_label(
        text,
        "Consignee",
        stop_labels=["Notify party", "Voyage No.", "Vessel", "Port of Loading", "Port of Discharge", "Place of Receipt"],
        first_line_only=True,
    ) or ""

    vessel = (
        first_match(text, r"(?im)^\s*Vessel\s*(?:\([^)]*\))?\s*:?\s*\n\s*([^\n]+)")
        or first_match(text, r"(?im)^\s*Vessel\s*(?:\([^)]*\))?\s*:?\s*([^\n]+)")
        or ""
    )
    vessel = clean_one_line(vessel)

    port_loading = line_value_after_label(
        text,
        labels=["Port of Loading", "PORT OF LOADING", "P.O.L.", "POL"],
    ) or ""
    port_loading = clean_one_line(port_loading)

    port_discharge = line_value_after_label(
        text,
        labels=["Port of Discharge", "PORT OF DISCHARGE", "P.O.D.", "POD"],
    ) or ""
    port_discharge = clean_one_line(port_discharge)

    container_no = first_match(text, r"\b(MRSU\d{7,})\b") or ""

    # Extract packages with "PACKAGES" suffix
    packages = first_match(text, r"(?im)Said to Contain\s+([\d,]+\s+PACKAGES)") or ""
    if not packages:
        packages = first_match(text, r"(?im)\b([\d,]+\s+PACKAGES)\b") or ""

    # Extract weight with "KGS" suffix
    weight_kgs = first_match(text, r"(?im)\b([\d,.]+\s*KGS)\b") or ""
    
    # Extract measurement with "CBM" suffix
    measurement_cbm = first_match(text, r"(?im)\b([\d,.]+\s*CBM)\b") or ""

    return {
        "bl_no": bl_no,
        "shipper": shipper,
        "consignee": consignee,
        "vessel": vessel,
        "port_of_loading": port_loading,
        "port_of_discharge": port_discharge,
        "packages": packages,
        "weight_kgs": weight_kgs,
        "measurement_cbm": measurement_cbm,
        "container_no": container_no,
    }


def get_headers_and_values(fields: Dict[str, Any]) -> tuple[list[str], list[str]]:
    """Returns the headers and values for output files."""
    headers = [
        "bl_no",
        "shipper",
        "consignee",
        "vessel",
        "port_of_loading",
        "port_of_discharge",
        "contain",
        "weight",
        "measurement",
        "mrsu",
    ]
    values = [
        fields.get("bl_no", ""),
        fields.get("shipper", ""),
        fields.get("consignee", ""),
        fields.get("vessel", ""),
        fields.get("port_of_loading", ""),
        fields.get("port_of_discharge", ""),
        fields.get("packages", ""),
        fields.get("weight_kgs", ""),
        fields.get("measurement_cbm", ""),
        fields.get("container_no", ""),
    ]
    return headers, values


def write_csv(out_path: Path, fields: Dict[str, Any]) -> None:
    """Write extracted fields to CSV file."""
    headers, values = get_headers_and_values(fields)
    
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        # Write headers without quotes
        writer.writerow(headers)
        # Write values with selective quoting
        # Build the row manually to control quoting
        formatted_row = values[0]  # bl_no without quotes
        for val in values[1:-1]:  # middle values with quotes
            formatted_row += ',"' + str(val) + '"'
        formatted_row += ',' + values[-1]  # mrsu without quotes
        # Use the writer to ensure consistent line endings
        f.write(formatted_row + '\r\n')


def write_xlsx(out_path: Path, fields: Dict[str, Any]) -> None:
    """Write extracted fields to XLSX file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "import_partida"

    headers, values = get_headers_and_values(fields)
    ws.append(headers)
    ws.append(values)
    wb.save(str(out_path))


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        return die("Uso: extract_import_partida_fields.py <input.pdf> <out_dir>")

    pdf_path = Path(argv[1]).expanduser().resolve()
    out_dir = Path(argv[2]).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not pdf_path.exists():
        return die(f"ERROR: no existe el PDF: {pdf_path}")
    if pdf_path.suffix.lower() != ".pdf":
        return die("ERROR: input no es .pdf")

    text = read_pdf_text(pdf_path)
    fields = extract_fields(text)

    (out_dir / "extracted.json").write_text(
        json.dumps(fields, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    write_csv(out_dir / "import_partida.csv", fields)
    write_xlsx(out_dir / "import_partida.xlsx", fields)

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
