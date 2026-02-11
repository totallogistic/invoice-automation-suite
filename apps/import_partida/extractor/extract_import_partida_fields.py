#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional

import pdfplumber


def die(msg: str, code: int = 2) -> int:
    print(msg, file=sys.stderr, flush=True)
    return code


def read_pdf_lines(pdf_path: Path) -> List[str]:
    lines: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            for ln in txt.splitlines():
                ln = ln.rstrip()
                if ln.strip():
                    lines.append(ln)
    return lines


def find_first_regex(text: str, patterns: List[str], flags: int = re.IGNORECASE) -> str:
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return (m.group(1) or "").strip()
    return ""


def next_nonempty_line(lines: List[str], idx: int) -> str:
    for j in range(idx + 1, min(idx + 8, len(lines))):
        if lines[j].strip():
            return lines[j].strip()
    return ""


def collect_block_until(lines: List[str], start_idx: int, stop_regex: str, max_lines: int = 6) -> List[str]:
    out: List[str] = []
    for j in range(start_idx + 1, min(start_idx + 1 + max_lines, len(lines))):
        if re.search(stop_regex, lines[j], re.IGNORECASE):
            break
        if lines[j].strip():
            out.append(lines[j].strip())
    return out


def clean_company(s: str) -> str:
    s = s.strip()
    # Quita dobles espacios
    s = re.sub(r"\s+", " ", s)
    # Quita puntuación final típica
    s = s.rstrip(" .,:;")
    return s.strip()


def extract_ports_from_layout(lines: List[str]) -> Dict[str, str]:
    """
    En este PDF concreto, aparece una línea:
      "Port of Loading Port of Discharge ..."
    y en la siguiente línea los valores:
      "NINGBO, CHINA Valencia,Spain"
    """
    pol = ""
    pod = ""

    for i, ln in enumerate(lines):
        if ("Port of Loading" in ln) and ("Port of Discharge" in ln):
            vals = next_nonempty_line(lines, i)
            if not vals:
                break

            # Si hay separación clara por múltiples espacios, úsala
            parts = re.split(r"\s{2,}", vals.strip())
            if len(parts) >= 2:
                pol = parts[0].strip()
                pod = parts[1].strip().rstrip(":")
                break

            # Si NO hay separación clara: heurística (como en tu ejemplo)
            # "NINGBO, CHINA Valencia,Spain" -> pol="NINGBO, CHINA" pod="Valencia,Spain"
            m = re.match(r"^(?P<pol>.+?)\s+(?P<pod>[A-Z][A-Za-z].+)$", vals.strip())
            if m:
                pol = (m.group("pol") or "").strip()
                pod = (m.group("pod") or "").strip()
                pod = pod.rstrip(":")
                break

            # Fallback: si no logramos separar, al menos volcamos todo en loading
            pol = vals.strip()
            pod = ""
            break

    return {
        "port_of_loading": clean_company(pol),
        "port_of_discharge": clean_company(pod),
    }


def main() -> int:
    if len(sys.argv) != 3:
        return die("Uso: extract_import_partida_fields.py <input.pdf> <out_dir>")

    pdf_path = Path(sys.argv[1]).expanduser().resolve()
    out_dir = Path(sys.argv[2]).expanduser().resolve()

    if not pdf_path.exists():
        return die(f"ERROR: PDF no existe: {pdf_path}")

    out_dir.mkdir(parents=True, exist_ok=True)

    lines = read_pdf_lines(pdf_path)
    full_text = "\n".join(lines)

    # 1) B/L No.
    bl_no = find_first_regex(
        full_text,
        [
            r"\bB/L\s*No\.?\s*([A-Z0-9]+)\b",
            r"\bB\/L\s*No\.?\s*([A-Z0-9]+)\b",
        ],
        flags=re.IGNORECASE,
    )

    # 2) Vessel
    vessel = find_first_regex(
        full_text,
        [
            r"\bVessel\s*([^\n]+)",
        ],
        flags=re.IGNORECASE,
    )
    vessel = clean_company(vessel)

    # 3) Shipper: en este PDF el valor viene en la línea siguiente a "Shipper"
    shipper = ""
    for i, ln in enumerate(lines):
        if re.search(r"\bShipper\b", ln, re.IGNORECASE):
            v = next_nonempty_line(lines, i)
            if v:
                # La misma línea suele traer el Booking/BL al final: "... NPOS56802"
                if bl_no and v.endswith(bl_no):
                    v = v[: -len(bl_no)].strip()
                shipper = clean_company(v)
            break

    # 4) Consignee: en tu PDF viene con el prefijo “As principal, where …)”
    consignee = ""
    for i, ln in enumerate(lines):
        if re.search(r"\bConsignee\b", ln, re.IGNORECASE):
            block = collect_block_until(lines, i, stop_regex=r"\bVessel\b", max_lines=6)
            joined = " ".join(block)
            joined = re.sub(r"\s+", " ", joined).strip()

            # Si hay un ") ..." nos quedamos con lo que va después
            if ")" in joined:
                joined = joined.split(")")[-1].strip()

            # Quita el prefijo exacto que te está ensuciando (más robusto)
            joined = re.sub(
                r"^As principal, where\s+“care of”,\s+“c/o”,\s+or other variants used\.\)\s*",
                "",
                joined,
                flags=re.IGNORECASE,
            )

            consignee = clean_company(joined)
            break

    # 5) Ports (layout)
    ports = extract_ports_from_layout(lines)

    # 6) Contain / Weight / Measurement / MRSU
    contain = find_first_regex(full_text, [r"\b(\d+\s+PACKAGES)\b"], flags=re.IGNORECASE)
    weight = find_first_regex(full_text, [r"\b(\d[\d.,]*\s*KGS)\b"], flags=re.IGNORECASE)
    measurement = find_first_regex(full_text, [r"\b(\d[\d.,]*\s*CBM)\b"], flags=re.IGNORECASE)
    mrsu = find_first_regex(full_text, [r"\b(MRSU\d+)\b"], flags=re.IGNORECASE)

    data = {
        "bl_no": clean_company(bl_no),
        "shipper": clean_company(shipper),
        "consignee": clean_company(consignee),
        "vessel": clean_company(vessel),
        "port_of_loading": clean_company(ports.get("port_of_loading", "")),
        "port_of_discharge": clean_company(ports.get("port_of_discharge", "")),
        "contain": clean_company(contain),
        "weight": clean_company(weight),
        "measurement": clean_company(measurement),
        "mrsu": clean_company(mrsu),
    }

    # JSON
    (out_dir / "extracted.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    # CSV esperado (delimiter ;)
    csv_path = out_dir / "import_partida.csv"
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
    with csv_path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=headers, delimiter=";")
        w.writeheader()
        w.writerow({k: data.get(k, "") for k in headers})

    return 0


if __name__ == "__main__":
    raise SystemExit(main())