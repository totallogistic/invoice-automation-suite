#!/usr/bin/env python3
"""
Croton packing + factura/mapeo -> Resumen_Partidas

What this version fixes
-----------------------
- Packing input can be .ods or .xlsx
- Factura/mapeo input can be .ods, .xlsx or .xls
- product_mapping.csv is OPTIONAL, not mandatory
- If there is factura, VALOR is allocated from factura totals for the whole matched
  reference bucket, avoiding mixed ODS/manual partial values and double counting
- The result is injected into a copy of the source workbook as sheet
  'Resumen_Partidas'

Accepted packing layouts
------------------------
1) Enriched/manual layout (already contains MERCANCIA + PARTIDA)
2) Raw packing-list layout (DESCRIPCION + CODIGO + NETO/BRUTO + CANT TOTAL + m2)

For raw layouts, classification is resolved using:
- optional CSV mapping if provided / present next to script
- otherwise built-in heuristics from description / invoice description
"""
from __future__ import annotations
SCRIPT_VERSION = "2026-03-28.v2"

SCRIPT_CHANGELOG = """
## 2026-03-28.v2

### Fixes respecto a v1
- CRITICO: load_factura usaba indices de columna erróneos para XLSX
  (cant=col5, neto=col7, importe=col8 en vez de col6/col8/col9)
  -> Esto hacía que VALOR = NETO (kg) en vez de EUR
- SARGA: ahora distingue fibra primaria (ALG vs POL/SINT) para separar
  partidas 5211320090 (algodón predominante) vs 5514220000 (sint predominante)
- TAFETAN/POPELIN: ahora distingue blanqueado (BLANCO) -> 5513112000
  vs teñido -> 5513210000
- PUNTO: ahora detecta tejidos sin POL (100%ALG o ALG+ACR/VISC sin POL)
  -> 6006220000, y tejidos con POL -> 6006320000
- REJILLA: siempre 5804109000 (eliminado caso especial para BLANCA)
- CANALE colores: ahora se agrupa bajo misma descripción que resto de
  6006320000, evitando filas duplicadas
- Peso palés: los kg de paletas se añaden al primer grupo en BRUTO
- Descripciones: terminología aduanera completa en lugar de abreviada
- load_factura: nuevo soporte para múltiples líneas del mismo CODIGO con
  distinta descripción (ej. RFCANALE BLANCO vs NARANJA), matching por
  descripción cuando hay más de una entrada para el mismo ref

## 2026-03-17.v1 (original)
- Primera versión funcional
"""

import argparse
import csv
import logging
import os
import re
import shutil
import zipfile
from collections import OrderedDict, defaultdict
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple
from xml.etree import ElementTree as ET

import openpyxl
from odf.opendocument import load, OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P
from odf.style import Style, TextProperties, TableCellProperties
from openpyxl.workbook.views import BookView

EXTRACTOR_DIR = Path(__file__).parent
DEFAULT_MAPPING_FILE = EXTRACTOR_DIR / "product_mapping.csv"
# Alternative names the mapping CSV may be saved under
_MAPPING_CANDIDATES = [
    "product_mapping.csv",
    "mapping_partidas_v2.csv",
    "mapping_partidas.csv",
]

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("croton")

COLOR_WORDS = {
    "NEGRO", "NEGRA", "MARINO", "AMARILLO", "AMARILLA", "AZUL", "AZULINA",
    "CELESTE", "ROJO", "ROJA", "VERDE", "GRIS", "MORADO", "MORADA", "MARRON",
    "TURQUESA", "CEREZA", "NARANJA", "BLANCO", "BLANCA", "BEIGE",
}

BUILTIN_REF_ALIASES = {
    "RF008": "RF004",
    "KL016": "KL004",
    "RF388": "RF385",
    "RFJER3": "RFJER1",
}


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def _parse_num(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("\xa0", "")
    if text in ("", "#¡DIV/0!", "#DIV/0!", "#N/A", "#REF!"):
        return None
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def _fmt(value: float) -> str:
    if value in (None, 0, 0.0):
        return ""
    s = f"{value:.4f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def _safe_str(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _norm_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip())


def _norm_ref(ref: str) -> str:
    return re.sub(r"\s+", "", (ref or "").upper())


def _normalize_desc_for_match(text: str) -> str:
    t = _norm_spaces((text or "").upper())
    t = re.sub(r"A\s*1[,\.]\d+", "", t)
    # Normalise spaces around % before other substitutions: "65% POL" → "65%POL"
    t = re.sub(r"(\d)\s*%\s*([A-Z])", r"\1%\2", t)
    t = re.sub(r"\b\d+[%]?[A-Z]*\b", lambda m: m.group(0).replace(" ", ""), t)
    for color in COLOR_WORDS:
        t = re.sub(rf"\b{re.escape(color)}\b", "COLORES", t)
    t = re.sub(r"\bCOLORES(?:\s+COLORES)+\b", "COLORES", t)
    t = re.sub(r"[^A-Z0-9% ]+", " ", t)
    return _norm_spaces(t)


def _extract_fiber_pcts(desc: str) -> Dict[str, float]:
    """
    Extract fiber percentages from a description like '65%POL 35%ALG' or '50%POL-50%ALG'.
    Returns dict mapping fiber code (ALG, POL, ACR, VISC, PA, ELAST...) to percentage.
    """
    result: Dict[str, float] = {}
    # Match patterns like "67%POL", "33% ALG", "50 % POLIAM", "2%ELAST", "1%ANTIEST"
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*%\s*([A-Z]+(?:\s*MOD)?)", desc.upper()):
        pct = float(m.group(1).replace(",", "."))
        fiber = re.sub(r"\s+", "", m.group(2))
        # Normalize known aliases
        fiber = {"POLIAM": "PA", "POLAM": "PA", "POLYAM": "PA",
                 "ANTIEST": "ANTIEST", "ELASTANO": "ELAST",
                 "ACRMOD": "ACRMOD"}.get(fiber, fiber)
        result[fiber] = result.get(fiber, 0.0) + pct
    return result


def _is_synthetic_dominant(fibers: Dict[str, float]) -> bool:
    """
    Returns True if synthetic fibers (POL, PA, ACR, ACRMOD) dominate over ALG.
    Viscose/Modal (VISC, MODAL) are artificial fibres, not synthetic for this purpose.
    """
    alg = fibers.get("ALG", 0.0)
    # Synthetic fibers: polyester (POL), polyamide (PA/POLIAM), acrylic (ACR/ACRMOD)
    sint = sum(v for k, v in fibers.items() if k in ("POL", "PA", "ACR", "ACRMOD"))
    return sint > alg


def _has_pol(fibers: Dict[str, float]) -> bool:
    """Returns True if polyester (POL) is present in fabric."""
    return fibers.get("POL", 0.0) > 0


def _is_blanqueado(desc: str) -> bool:
    up = desc.upper()
    return "BLANCO" in up or "BLANCA" in up


def _extract_width(desc: str) -> float:
    """Extract fabric width in metres from description, e.g. 'A 1,60' or 'A 100' (cm)."""
    m = re.search(r"\bA\s+(\d+(?:[.,]\d+)?)\b", (desc or "").upper())
    if m:
        val = float(m.group(1).replace(",", "."))
        return val / 100.0 if val >= 10 else val   # >=10 → centimetres
    return 0.0


def _calc_gramaje(neto_kg: float, m2: float, cant_total: float = 0.0, desc: str = "") -> Optional[float]:
    """Return gramaje in g/m².  Uses m2 if available, else cant_total × width."""
    if m2 and m2 > 0:
        return neto_kg / m2 * 1000.0
    if cant_total and cant_total > 0:
        ancho = _extract_width(desc)
        if ancho > 0:
            return neto_kg / (cant_total * ancho) * 1000.0
    return None


# ---------------------------------------------------------------------------
# Read ODS sheets
# ---------------------------------------------------------------------------

def _ods_sheet_names(path: str) -> List[str]:
    with zipfile.ZipFile(path) as zf:
        with zf.open("content.xml") as f:
            tree = ET.parse(f)
    ns = {"table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0"}
    return [
        s.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name")
        for s in tree.getroot().findall(".//table:table", ns)
    ]


def _read_ods_sheet(path: str, sheet_name: str) -> List[List[str]]:
    with zipfile.ZipFile(path) as zf:
        with zf.open("content.xml") as f:
            tree = ET.parse(f)
    root = tree.getroot()
    ns = {
        "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
        "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    }
    for sheet in root.findall(".//table:table", ns):
        name = sheet.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}name")
        if name != sheet_name:
            continue
        rows: List[List[str]] = []
        for row in sheet.findall("table:table-row", ns):
            vals = []
            for cell in row.findall("table:table-cell", ns):
                rep = cell.get("{urn:oasis:names:tc:opendocument:xmlns:table:1.0}number-columns-repeated")
                txt = "".join(p.text or "" for p in cell.findall(".//text:p", ns))
                n = int(rep) if rep and int(rep) < 100 else 1
                vals.extend([txt] * n)
            rows.append(vals)
        return rows
    raise ValueError(f"Sheet '{sheet_name}' not found in {path}")


# ---------------------------------------------------------------------------
# Read XLSX sheets
# ---------------------------------------------------------------------------

def _xlsx_sheet_names(path: str) -> List[str]:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        return wb.sheetnames[:]
    finally:
        wb.close()


def _read_xlsx_sheet(path: str, sheet_name: str) -> List[List[Any]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[sheet_name]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _sheet_names(path: str) -> List[str]:
    ext = Path(path).suffix.lower()
    if ext == ".ods":
        return _ods_sheet_names(path)
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        return _xlsx_sheet_names(path)
    raise ValueError(f"Unsupported packing extension: {ext}")


def _read_sheet(path: str, sheet_name: str) -> List[List[Any]]:
    ext = Path(path).suffix.lower()
    if ext == ".ods":
        return _read_ods_sheet(path, sheet_name)
    return _read_xlsx_sheet(path, sheet_name)


# ---------------------------------------------------------------------------
# Detect packing layout and parse groups
# ---------------------------------------------------------------------------

def _detect_packing_sheet(path: str) -> str:
    names = _sheet_names(path)
    preferred = []
    for n in names:
        up = n.upper()
        if "1_2" in up or "1.2" in up:
            return n
        if up in ("HOJA1", "HOJA5"):
            preferred.append(n)
    if preferred:
        preferred_sorted = sorted(preferred, key=lambda x: 0 if x.upper() == "HOJA1" else 1)
        return preferred_sorted[0]
    return names[0]


def _row_has(values: List[Any], idx: int, needle: str) -> bool:
    return len(values) > idx and needle.upper() in _safe_str(values[idx]).upper()


def _detect_layout(rows: List[List[Any]]) -> str:
    scan = rows[:12]
    for r in scan:
        if _row_has(r, 1, "MERCANCIA") and _row_has(r, 2, "PARTIDA"):
            return "enriched_classic"
        if _row_has(r, 0, "Nº BULTO") and _row_has(r, 2, "CODIGO") and _row_has(r, 7, "CANT. TOTAL"):
            return "enriched_headerless"
        if _row_has(r, 0, "LINEA") and _row_has(r, 2, "DESCRIPCION") and _row_has(r, 3, "CODIGO"):
            return "raw_packing"
    for r in rows[:10]:
        vals = [_safe_str(v).upper() for v in r[:5]]
        if len(vals) >= 3 and vals[1] == "MERCANCIA" and vals[2] == "PARTIDA":
            return "enriched_classic"
    return "raw_packing"


def _parse_enriched_classic(rows: List[List[Any]]) -> Tuple[List[Dict[str, Any]], float]:
    result = []
    for row in rows:
        row = list(row)
        while len(row) < 15:
            row.append(None)
        bulto = _safe_str(row[0])
        mercancia = _safe_str(row[1])
        partida = _safe_str(row[2])
        if not bulto or bulto.upper() == "Nº BULTO" or not mercancia or mercancia.upper() == "MERCANCIA":
            continue
        if not partida or not partida.replace(".", "").isdigit():
            continue
        cant_total = _parse_num(row[7])
        if cant_total is None:
            continue
        result.append({
            "referencia": partida,
            "descripcion": mercancia,
            "mercancia": mercancia,
            "partida_arancel": partida,
            "cant_total": cant_total,
            "m2": _parse_num(row[8]) or 0.0,
            "neto": _parse_num(row[10]) or 0.0,
            "bruto": _parse_num(row[11]) or 0.0,
            "bultos": _parse_num(row[12]) or 0.0,
            "valor": _parse_num(row[13]),
            "source_layout": "enriched_classic",
        })
    return result, 0.0


def _parse_enriched_headerless(rows: List[List[Any]]) -> Tuple[List[Dict[str, Any]], float]:
    result, _ = _parse_enriched_classic(rows)
    return result, 0.0


def _parse_raw_packing(rows: List[List[Any]]) -> Tuple[List[Dict[str, Any]], float]:
    """
    Raw packing-list layout:
    col 0: LINEA, 1: Nº BULTO, 2: DESCRIPCION, 3: CODIGO, 4: NºPALET,
    col 5: NETO, 6: BRUTO, 7: CANT, 8: CANT TOTAL, 9: M2

    Returns (lineas, pallet_bruto_kg) where pallet_bruto_kg is the extra weight
    of pallets found in the footer rows (e.g. "7 PALETS  133 kg").
    """
    result = []
    cur_rows = []
    cur_neto = 0.0
    cur_bruto = 0.0
    pallet_bruto = 0.0

    start = 0
    for i, r in enumerate(rows):
        if _row_has(r, 0, "LINEA") and _row_has(r, 2, "DESCRIPCION"):
            start = i + 1
            break

    for row in rows[start:]:
        row = list(row)
        while len(row) < 10:
            row.append(None)

        col0 = _safe_str(row[0]).upper()
        col2 = _safe_str(row[2]).upper()

        # Detect pallet weight row: "7 PALETS" text in col2 or col0, weight in col6
        if "PALET" in col0 or "PALET" in col2:
            bruto_val = _parse_num(row[6])
            if bruto_val:
                pallet_bruto += bruto_val
                log.info("Detected pallet weight: %.2f kg from '%s'", bruto_val, row[2] or row[0])
            continue

        # Skip TOTAL rows
        if "TOTAL" in col0:
            continue

        desc = _safe_str(row[2])
        code = _safe_str(row[3])
        if not desc or not code:
            continue

        cur_rows.append(row)
        cur_neto += _parse_num(row[5]) or 0.0
        cur_bruto += _parse_num(row[6]) or 0.0

        cant_total = _parse_num(row[8])
        if cant_total is None:
            continue

        # End of a LINEA group
        result.append({
            "referencia": code,
            "descripcion": desc,
            "cant_total": cant_total,
            "m2": _parse_num(row[9]) or 0.0,
            "neto": round(cur_neto, 2),
            "bruto": round(cur_bruto, 2),
            "bultos": float(len(cur_rows)),
            "valor": None,
            "source_layout": "raw_packing",
        })
        cur_rows = []
        cur_neto = 0.0
        cur_bruto = 0.0

    return result, pallet_bruto


def read_packing(path: str) -> Tuple[List[Dict[str, Any]], str, str, float]:
    sheet_name = _detect_packing_sheet(path)
    rows = _read_sheet(path, sheet_name)
    layout = _detect_layout(rows)
    log.info("Packing '%s' -> sheet '%s' layout=%s", path, sheet_name, layout)
    if layout == "enriched_classic":
        lineas, pallet_bruto = _parse_enriched_classic(rows)
    elif layout == "enriched_headerless":
        lineas, pallet_bruto = _parse_enriched_headerless(rows)
    else:
        lineas, pallet_bruto = _parse_raw_packing(rows)
    return lineas, sheet_name, layout, pallet_bruto


# ---------------------------------------------------------------------------
# Optional CSV mapping
# ---------------------------------------------------------------------------

def load_mapping(csv_path: str) -> List[Dict[str, str]]:
    """
    Load product mapping CSV.  Accepts two column layouts:

    Classic layout:  CODIGO | DESCRIPCION_CONTAINS | MERCANCIA | PARTIDA
    Extended layout: CODIGO | DESCRIPCION | … | PARTIDA | DESC_ADUANERA | …

    Both layouts are auto-detected by checking the header row.
    """
    rules = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = [c.strip().upper() for c in (reader.fieldnames or [])]
        # Detect layout
        has_extended = "DESC_ADUANERA" in fieldnames
        for row in reader:
            codigo = row.get("CODIGO", "").strip()
            if not codigo or codigo.startswith("#") or codigo.startswith("?"):
                continue
            if has_extended:
                # Extended layout: CODIGO is the unique key, no description filter needed.
                # DESC_ADUANERA is the customs description to use.
                mercancia = row.get("DESC_ADUANERA", "").strip()
                contains  = ""   # match by CODIGO only
            else:
                mercancia = row.get("MERCANCIA", "").strip()
                contains  = row.get("DESCRIPCION_CONTAINS", "").strip().upper()
            partida = row.get("PARTIDA", "").strip()
            if not partida or not mercancia:
                continue
            rules.append({
                "codigo":   _norm_ref(codigo),
                "contains": contains,
                "mercancia": mercancia,
                "partida":  partida,
            })
    log.info("Loaded %d mapping rules from %s", len(rules), csv_path)
    return rules


def lookup_mapping(referencia: str, descripcion: str, rules: List[Dict[str, str]]) -> Optional[Tuple[str, str]]:
    ref_up = _norm_ref(referencia)
    desc_up = _safe_str(descripcion).upper()
    for rule in rules:
        if rule["codigo"] == ref_up and rule["contains"] and rule["contains"] in desc_up:
            return rule["mercancia"], rule["partida"]
    for rule in rules:
        if rule["codigo"] == ref_up and not rule["contains"]:
            return rule["mercancia"], rule["partida"]
    return None


# ---------------------------------------------------------------------------
# Factura loading and matching
# ---------------------------------------------------------------------------

def _detect_factura_columns(rows: List[List[Any]]) -> Tuple[int, int, int, int]:
    """
    Auto-detect column positions for: ref, desc, cant, importe in a factura sheet.
    Returns (ref_col, desc_col, cant_col, importe_col).

    Known XLSX layout (Mendez & Croton):
      col 0: REFERENCIA
      col 1: DESCRIPCIÓN
      col 5: CANTIDAD MTS/KGS
      col 6: PRECIO MT/KG
      col 7: PESO NETO KGS
      col 8: IMPORTE (EUR)
      col 9: KGS (repeat)

    ODS / legacy layout may differ - try header detection.
    """
    # Try to find the header row
    for r in rows[:20]:
        row_strs = [_safe_str(v).upper() for v in r]
        row_joined = " ".join(row_strs)
        if "REFERENCIA" in row_joined and ("CANTIDAD" in row_joined or "IMPORTE" in row_joined):
            # Found header row - scan for key columns
            ref_col = desc_col = cant_col = importe_col = None
            for i, cell in enumerate(row_strs):
                if "REFERENCIA" in cell and ref_col is None:
                    ref_col = i
                elif "DESCRIPCI" in cell and desc_col is None:
                    desc_col = i
                elif ("CANTIDAD" in cell or "CANT" in cell) and cant_col is None:
                    cant_col = i
                elif "IMPORTE" in cell and importe_col is None:
                    importe_col = i
            if ref_col is not None and importe_col is not None:
                log.info("Factura columns auto-detected: ref=%d desc=%d cant=%d importe=%d",
                         ref_col, desc_col or 1, cant_col or 5, importe_col)
                return ref_col, desc_col or 1, cant_col or 5, importe_col

    # Fallback to known XLSX layout
    log.info("Factura columns: using default XLSX layout (ref=0, desc=1, cant=5, importe=8)")
    return 0, 1, 5, 8


def load_factura(path: str) -> Dict[str, List[Dict[str, Any]]]:
    """
    Load factura. Returns dict mapping normalised reference -> list of entries.
    Multiple entries with the same reference are kept separate (they may have
    different descriptions, e.g. RFCANALE BLANCO vs RFCANALE NARANJA).
    """
    ext = Path(path).suffix.lower()
    if ext == ".ods":
        names = _ods_sheet_names(path)
        sheet_name = next((n for n in names if "TOTAL FRA" in n.upper()), names[0])
        rows: List[List[Any]] = _read_ods_sheet(path, sheet_name)
    else:
        wb = openpyxl.load_workbook(path, data_only=True)
        try:
            ws = wb["TOTAL FRA."] if "TOTAL FRA." in wb.sheetnames else wb.active
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
        finally:
            wb.close()

    ref_col, desc_col, cant_col, importe_col = _detect_factura_columns(rows)

    # result: ref -> list of entry dicts (NOT merged, kept per factura line)
    result: Dict[str, List[Dict[str, Any]]] = OrderedDict()

    for row in rows:
        while len(row) <= max(ref_col, desc_col, cant_col, importe_col):
            row.append(None)
        ref = _safe_str(row[ref_col])
        desc = _safe_str(row[desc_col])
        cant = _parse_num(row[cant_col])
        importe = _parse_num(row[importe_col])
        if not ref or ref.upper() == "REFERENCIA" or cant is None or importe is None:
            continue
        refn = _norm_ref(ref)
        entry = {
            "referencia": refn,
            "descripcion": desc,
            "cant_total": cant,
            "importe_total": importe,
            "desc_norm": _normalize_desc_for_match(desc),
        }
        result.setdefault(refn, []).append(entry)

    total_importe = sum(e["importe_total"] for entries in result.values() for e in entries)
    log.info("Loaded FACTURA refs=%d total=%.2f EUR", len(result), total_importe)
    return result


def infer_invoice_ref(linea: Dict[str, Any], factura: Dict[str, List[Dict[str, Any]]]) -> Optional[Tuple[str, int]]:
    """
    Returns (ref_key, entry_index) for the best factura match, or None.
    If a ref has multiple entries, tries to pick by description similarity.
    """
    ref = _norm_ref(linea.get("referencia", ""))

    def _best_entry_idx(refn: str) -> int:
        entries = factura[refn]
        if len(entries) == 1:
            return 0
        # Multiple entries: try description match using normalised words first,
        # then break ties using colour words from the ORIGINAL (non-normalised)
        # description so that e.g. "MARINO" vs "AZULINA" is disambiguated.
        desc_norm = _normalize_desc_for_match(linea.get("descripcion", ""))
        orig_upper = (linea.get("descripcion", "") or "").upper()
        best = 0
        best_score = (-1, -1)
        for i, e in enumerate(entries):
            a_words = set(desc_norm.split())
            b_words = set(e["desc_norm"].split())
            score_norm = len(a_words & b_words)
            # Secondary: count color/detail words in common with original descriptions
            orig_entry_upper = (e.get("descripcion", "") or "").upper()
            orig_a_words = set(orig_upper.split())
            orig_b_words = set(orig_entry_upper.split())
            score_color = len(orig_a_words & orig_b_words)
            score = (score_norm, score_color)
            if score > best_score:
                best_score = score
                best = i
        return best

    if ref in factura:
        return ref, _best_entry_idx(ref)

    alias = BUILTIN_REF_ALIASES.get(ref)
    if alias and alias in factura:
        return alias, _best_entry_idx(alias)

    # Try description-only match
    desc_norm = _normalize_desc_for_match(linea.get("descripcion", ""))
    candidates = []
    for k, entries in factura.items():
        for i, e in enumerate(entries):
            if e.get("desc_norm") == desc_norm:
                candidates.append((k, i))
    if len(candidates) == 1:
        return candidates[0]
    return None


# ---------------------------------------------------------------------------
# Classification without mandatory CSV
# ---------------------------------------------------------------------------

def _contains(text: str, *needles: str) -> bool:
    up = (text or "").upper()
    return any(n.upper() in up for n in needles)


def classify_from_text(descripcion: str, referencia: str = "", factura_desc: str = "",
                       gramaje: Optional[float] = None) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    Returns (mercancia_desc, partida_arancelaria, error_msg_or_None).

    EU CN textile classification using fiber composition + gramaje (g/m²):

    SARGAS (twill weave):
      Sint ≥85%, >170 g/m²  → 5512.xx
      Sint <85% dominant, >170 g/m² → 5514 (blanqueada=5514120000, teñida=5514220000)
      Sint <85% dominant, ≤170 g/m² → 5513 (blanqueada=5513112000, teñida=5513210000)
      ALG ≥85%, >200 g/m²  → 5209 (blanqueada=5209220000, teñida=5209320000)
      ALG ≥85%, ≤200 g/m²  → 5208.xx (rare)
      ALG dom. <85% mixed, >200 g/m² → 5211320090
      ALG dom. <85% mixed, ≤200 g/m² → 5210320090

    TAFETAN/PLANA (plain weave), ≤170 g/m²:
      Sint ≥85% → 5512199000 (teñida)
      Sint <85% dominant → 5513210000 (teñida), 5513112000 (blanqueada)
      ALG dom. → 5210/5211/5208 (uncommon in Croton)

    PUNTO (knitted): no gramaje thresholds, fiber dominance only.
    FELPA: fiber dominance determines 6001920000 (sint) vs 6001910000 (alg).
    REJILLA: fiber dominance determines 5804101000 (sint) vs 5804109000 (alg/other).
    """
    desc = (factura_desc or descripcion or "").upper()
    fibers = _extract_fiber_pcts(factura_desc or descripcion or "")
    blanqueado = _is_blanqueado(factura_desc or descripcion or "")

    alg  = fibers.get("ALG", 0.0)
    pol  = fibers.get("POL", 0.0)
    # man-made = synthetic (POL, PA, ACR, ACRMOD) + artificial (VISC, MODAL)
    man_made = sum(v for k, v in fibers.items() if k in ("POL", "PA", "ACR", "ACRMOD", "VISC", "MODAL"))
    # pure synthetic (not artificial): POL, PA, ACR, ACRMOD
    sint = sum(v for k, v in fibers.items() if k in ("POL", "PA", "ACR", "ACRMOD"))
    sint_dominant = man_made > alg  # man-made > cotton → synthetic chapter
    alg_high = alg >= 85.0          # ≥85% cotton → pure-cotton chapters (5208/5209)
    sint_high = sint >= 85.0        # ≥85% synthetic → high-sint chapters (5512)

    g = gramaje  # may be None when not calculable

    # ---- Accessories / non-fabric ----------------------------------------
    if _contains(desc, "CREMALLERA"):
        if _contains(desc, "METAL"):
            return "CREMALLERA DIENTE METAL", "9607110000", None
        return "CREMALLERA DIENTE PLASTICO", "9607190000", None
    if _contains(desc, "TRANSFER"):
        return "TRANSFER", "5807909000", None
    if _contains(desc, "ANAGRAMA"):
        return "ANAGRAMAS", "5807101000", None
    if _contains(desc, "ETIQUETA"):
        if _contains(desc, "CARTON"):
            return "ETIQUETA CARTON", "4821109000", None
        return "ETIQUETAS COSER", "5807101000", None
    if _contains(desc, "REFLECTANTE"):
        return "REFLECTANTES", "3920610090", None

    # ---- Rejilla (lace / mesh) -------------------------------------------
    if _contains(desc, "REJILLA"):
        # 5804101000 = man-made fibres; 5804109000 = other (cotton-dominant)
        if sint_dominant:
            return "TEJIDOS DE REJILLA DE FIBRAS SINTETICAS", "5804101000", None
        return "TEJIDOS DE REJILLA", "5804109000", None

    # ---- Felpa (terry / velour) ------------------------------------------
    if _contains(desc, "FELPA"):
        # Tie (man_made = alg) → synthetic chapter (EU CN last-chapter rule)
        if man_made >= alg and man_made > 0:
            return "TEJIDOS DE FELPA DE FIBRAS SINTETICAS", "6001920000", None
        return "TEJIDOS DE FELPA PRED EL ALGODON", "6001910000", None

    # ---- Punto / tejido de punto (knitted) ------------------------------
    if _contains(desc, "PUNTO", "PIQUE", "CANALE"):
        if blanqueado:
            return "TEJIDOS DE PUNTO BLANQUEADOS DE FIBRAS SINTETICAS", "6006310000", None
        if man_made >= alg and man_made > 0:
            return "TEJIDOS TEÑIDOS DE PUNTO DE FIBRAS SINTETICAS MEZCLADOS CON ALGODON", "6006320000", None
        return "TEJIDOS TEÑIDOS DE PUNTO DE ALGODON", "6006220000", None

    # ---- Woven SARGA (twill) ---------------------------------------------
    if _contains(desc, "SARGA"):
        # Thresholds: 5513/5514 split at 170 g/m²; 5208/5209 vs 5210/5211 at 200 g/m²
        g_hi = g is not None and g > 200.0   # >200 g/m²
        g_mid = g is not None and g > 170.0  # >170 g/m²

        if blanqueado:
            if sint_high:                        # ≥85% sint, blanqueada
                return "TEJIDOS BLANQUEADOS DE SARGA DE FIBRAS SINTETICAS", "5512120000", None
            if sint_dominant:
                return "TEJIDOS BLANQUEADOS DE SARGA DE FIBRAS SINTETICAS MEZCLADOS CON ALGODON", "5514120000", None
            if alg_high:
                return "TEJIDOS BLANQUEADOS DE SARGA DE ALGODON", "5209220000", None
            # ALG dom. <85% mixed, blanqueada
            if g_hi:
                return "TEJIDOS BLANQUEADOS DE SARGA DE ALGODON MEZCLADOS CON FIBRAS SINTETICAS", "5211120090", None
            return "TEJIDOS BLANQUEADOS DE SARGA DE ALGODON MEZCLADOS CON FIBRAS SINTETICAS", "5210120090", None

        # Teñida
        if sint_high:                            # ≥85% sint, teñida
            return "TEJIDOS TEÑIDOS DE SARGA DE FIBRAS SINTETICAS", "5512199000", None
        if sint_dominant:
            # Sint <85% dominant: 5514 if >170 g/m², 5513 if ≤170 g/m²
            # Default to 5514 when gramaje unknown (most common case)
            if g is None or g_mid:
                return "TEJIDOS TEÑIDOS DE SARGA DE FIBRAS SINTETICAS MEZCLADAS CON ALGODON", "5514220000", None
            return "TEJIDOS TEÑIDOS DE SARGA DE FIBRAS SINTETICAS MEZCLADAS CON ALGODON", "5513210000", None
        if alg_high:
            return "TEJIDOS TEÑIDOS DE SARGA DE ALGODON", "5209320000", None
        # ALG dom. <85% mixed: 5211 if >200, 5210 if ≤200
        if g is None or g_hi:
            return "TEJIDOS TEÑIDOS DE SARGA DE ALGODON MEZCLADOS CON FIBRAS SINTETICAS", "5211320090", None
        return "TEJIDOS TEÑIDOS DE SARGA DE ALGODON MEZCLADOS CON FIBRAS SINTETICAS", "5210320090", None

    # ---- Woven TAFETAN / POPELIN / PLANA (plain weave) -------------------
    if _contains(desc, "POPELIN", "TAFETAN", "PLANA"):
        # Tafetán/plana are typically ≤170 g/m² → chapters 5512/5513
        if blanqueado:
            if sint_dominant:
                return "TEJIDOS BLANQUEADOS DE TAFETAN DE FIBRAS SINTETICAS MEZCLADOS CON ALGODON", "5513112000", None
            return "TEJIDOS BLANQUEADOS DE TAFETAN DE ALGODON", "5208110000", None

        if sint_high:                            # ≥85% sint, teñida
            return "TEJIDOS TEÑIDOS DE TAFETAN DE FIBRAS SINTETICAS", "5512199000", None
        if sint_dominant:
            return "TEJIDOS TEÑIDOS DE TAFETAN DE FIBRAS SINTETICAS MEZCLADOS CON ALGODON", "5513210000", None
        # Specific case from Croton: 65%POL 35%VISC → 5514301000 (plana)
        visc = fibers.get("VISC", 0.0)
        if visc > 0 and pol > 0 and not alg:
            return "TEJIDOS PLANOS DE FIBRAS SINTETICAS", "5514301000", None
        # ALG dominant tafetan (uncommon, e.g. 100%ALG POPELIN)
        return "TEJIDOS TEÑIDOS DE TAFETAN DE ALGODON", "5209310000", None

    return None, None, "No classification rule matched"


def apply_classification(lineas: List[Dict[str, Any]], factura: Optional[Dict[str, List[Dict[str, Any]]]] = None,
                         rules: Optional[List[Dict[str, str]]] = None) -> List[str]:
    issues: List[str] = []
    for l in lineas:
        if l.get("mercancia") and l.get("partida_arancel"):
            continue
        mapped = None
        if rules:
            mapped = lookup_mapping(l.get("referencia", ""), l.get("descripcion", ""), rules)
        if mapped:
            l["mercancia"], l["partida_arancel"] = mapped
            l["clasif_fuente"] = "CSV"
            continue
        match = infer_invoice_ref(l, factura) if factura else None
        l["invoice_ref"] = match[0] if match else None
        l["invoice_ref_idx"] = match[1] if match else None
        inv_desc = ""
        if factura and match:
            entries = factura.get(match[0], [])
            if match[1] < len(entries):
                inv_desc = entries[match[1]].get("descripcion", "")
        merc, part, err = classify_from_text(
            l.get("descripcion", ""), l.get("referencia", ""), inv_desc,
            gramaje=l.get("gramaje")
        )
        if merc and part:
            l["mercancia"] = merc
            l["partida_arancel"] = part
            l["clasif_fuente"] = "HEURISTIC"
        else:
            issues.append(f"Unclassified referencia={l.get('referencia')} descripcion={l.get('descripcion')}")
            l["mercancia"] = l.get("descripcion") or l.get("referencia")
            l["partida_arancel"] = "PENDIENTE"
            l["clasif_fuente"] = "UNCLASSIFIED"
    return issues


def apply_manual_business_rules(lineas: List[Dict[str, Any]]) -> None:
    """
    Croton-specific normalisation.
    Handles edge cases not captured by the general heuristic.
    """
    for l in lineas:
        desc = _norm_spaces(_safe_str(l.get("descripcion", "")).upper())
        inv_desc = _norm_spaces(_safe_str(l.get("_inv_desc", "")).upper())

        # PLANA BLANCA CUADRO VERDE - keep as its own row
        if desc == "PLANA BLANCA CUADRO VERDE 60%ALG 40%POL A 1,50":
            l["mercancia"] = "PLANA BLANCA CUADRO VERDE 60%ALG 40%POL A 1,50"
            l["partida_arancel"] = "5513210000"
            l["clasif_fuente"] = "MANUAL_RULE"
            l["keep_literal"] = True

        # Note: CANALE handling is now done in classify_from_text.
        # No override needed here - colored CANALE correctly goes to 6006320000
        # with the same description as other PUNTO TEÑIDO F SINT fabrics.


# ---------------------------------------------------------------------------
# VALOR allocation from factura
# ---------------------------------------------------------------------------

def enrich_valor_from_factura(lineas: List[Dict[str, Any]], factura: Dict[str, List[Dict[str, Any]]]) -> List[str]:
    """
    Assign EUR VALOR to each packing line from the factura.
    When a ref has multiple factura entries, match by description to assign
    exact import values rather than prorating.
    """
    issues: List[str] = []
    # Group packing lines by (inv_ref, inv_ref_idx) for prorating
    groups: Dict[Tuple[str, int], List[Dict[str, Any]]] = defaultdict(list)

    for l in lineas:
        match = infer_invoice_ref(l, factura)
        l["invoice_ref"] = match[0] if match else None
        l["invoice_ref_idx"] = match[1] if match else None
        if match:
            groups[(match[0], match[1])].append(l)
        else:
            if l.get("valor") is None:
                l["valor"] = 0.0
            l["valor_fuente"] = "NO_MATCH"
            issues.append(f"No factura match for referencia={l.get('referencia')} descripcion={l.get('descripcion')}")

    for (inv_ref, idx), items in groups.items():
        entries = factura.get(inv_ref, [])
        if idx >= len(entries):
            for i in items:
                i["valor"] = 0.0
                i["valor_fuente"] = "FACTURA_MISSING"
            continue
        inv = entries[idx]
        total_qty = sum((i.get("cant_total") or 0.0) for i in items)
        inv_qty = inv.get("cant_total") or 0.0
        base_qty = total_qty if total_qty > 0 else inv_qty
        importe_total = inv.get("importe_total") or 0.0

        if base_qty <= 0:
            for i in items:
                i["valor"] = 0.0
                i["valor_fuente"] = "FACTURA_ZERO"
            issues.append(f"Zero quantity for matched factura ref {inv_ref}")
            continue

        running = 0.0
        for pos, item in enumerate(items, start=1):
            qty = item.get("cant_total") or 0.0
            if pos < len(items):
                value = round((qty / base_qty) * importe_total, 2)
                running += value
            else:
                value = round(importe_total - running, 2)
            item["valor"] = value
            item["valor_fuente"] = f"FACTURA:{inv_ref}[{idx}]"
    return issues


# ---------------------------------------------------------------------------
# Aggregate
# ---------------------------------------------------------------------------

def aggregate(lineas: List[Dict[str, Any]], pallet_bruto: float = 0.0) -> List[Dict[str, Any]]:
    """
    Aggregate packing lines by (mercancia, partida_arancel).
    If pallet_bruto > 0, adds that weight to the BRUTO of the first group
    (following the manual operator practice of attributing pallet weight to
    the first line item).
    """
    seen: "OrderedDict[Tuple[str, str], Dict[str, Any]]" = OrderedDict()

    for l in lineas:
        key = (l["mercancia"], l["partida_arancel"])
        if key not in seen:
            seen[key] = {
                "mercancia": l["mercancia"],
                "partida": l["partida_arancel"],
                "bx": 0.0,
                "valor": 0.0,
                "valor_present": False,
                "bruto": 0.0,
                "neto": 0.0,
                "m2": 0.0,
            }
        g = seen[key]
        g["bx"] += l.get("bultos") or 0.0
        val = l.get("valor")
        if val is not None:
            g["valor"] += float(val or 0.0)
            g["valor_present"] = True
        g["bruto"] += l.get("bruto") or 0.0
        g["neto"] += l.get("neto") or 0.0
        g["m2"] += l.get("m2") or 0.0

    out = []
    first = True
    for g in seen.values():
        g["bx"] = int(round(g["bx"]))
        g["valor"] = round(g["valor"], 2) if g["valor_present"] else None
        g["bruto"] = round(g["bruto"], 2)
        g["neto"] = round(g["neto"], 2)
        g["m2"] = round(g["m2"], 2)
        g.pop("valor_present", None)
        # Add pallet weight to first fabric group
        if first and pallet_bruto > 0:
            g["bruto"] = round(g["bruto"] + pallet_bruto, 2)
            first = False
        elif first:
            first = False
        out.append(g)
    return out


# ---------------------------------------------------------------------------
# Output writers
# ---------------------------------------------------------------------------

def _make_cell(doc, value: str, bold: bool = False, bg: str = None) -> TableCell:
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


def _row_from_values(doc, values: List[Any], bold: bool = False, bg: Optional[str] = None) -> TableRow:
    row = TableRow()
    for v in values:
        row.addElement(_make_cell(doc, str(v) if v is not None else "", bold=bold, bg=bg))
    return row


def _summary_headers() -> List[str]:
    return ["MERCANCIA", "PARTIDA", "Suma - BX", "Suma - VALOR", "Suma - BRUTO", "Suma - NETO", "Suma - M22"]


def inject_summary_into_ods(summary: List[Dict[str, Any]], source_ods: str, output_path: str) -> None:
    shutil.copy2(source_ods, output_path)
    doc = load(output_path)
    for existing in doc.spreadsheet.getElementsByType(Table):
        if existing.getAttribute("name") == "Resumen_Partidas":
            doc.spreadsheet.removeChild(existing)
            break
    sheet = Table(name="Resumen_Partidas")
    doc.spreadsheet.addElement(sheet)
    sheet.addElement(_row_from_values(doc, _summary_headers(), bold=True, bg="#C0C0C0"))
    for r in summary:
        vals = [r["mercancia"], r["partida"], r["bx"], _fmt(r["valor"]), _fmt(r["bruto"]), _fmt(r["neto"]), _fmt(r["m2"])]
        sheet.addElement(_row_from_values(doc, vals))
    doc.save(output_path)


def inject_summary_into_xlsx(summary: List[Dict[str, Any]], source_xlsx: str, output_path: str,
                             issues: Optional[List[str]] = None,
                             detail: Optional[List[Dict[str, Any]]] = None) -> None:
    shutil.copy2(source_xlsx, output_path)
    wb = openpyxl.load_workbook(output_path)
    for name in ("Resumen_Partidas", "Issues", "Detalle_Extractor"):
        if name in wb.sheetnames:
            del wb[name]
    ws = wb.create_sheet("Resumen_Partidas")
    ws.append(_summary_headers())
    for r in summary:
        ws.append([r["mercancia"], r["partida"], r["bx"], r["valor"], r["bruto"], r["neto"], r["m2"]])
    wb.active = wb.worksheets[0]
    if wb.views:
        for bv in wb.views:
            bv.showSheetTabs = True
    else:
        wb.views.append(BookView(showSheetTabs=True))
    wb.save(output_path)
    wb.close()


def write_output(summary: List[Dict[str, Any]], output_path: str) -> None:
    ext = Path(output_path).suffix.lower()
    if ext == ".ods":
        doc = OpenDocumentSpreadsheet()
        sheet = Table(name="Resumen_Partidas")
        doc.spreadsheet.addElement(sheet)
        sheet.addElement(_row_from_values(doc, _summary_headers(), bold=True, bg="#C0C0C0"))
        for r in summary:
            vals = [r["mercancia"], r["partida"], r["bx"], _fmt(r["valor"]), _fmt(r["bruto"]), _fmt(r["neto"]), _fmt(r["m2"])]
            sheet.addElement(_row_from_values(doc, vals))
        doc.save(output_path)
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Resumen_Partidas"
        ws.append(_summary_headers())
        for r in summary:
            ws.append([r["mercancia"], r["partida"], r["bx"], r["valor"], r["bruto"], r["neto"], r["m2"]])
        wb.save(output_path)
        wb.close()


# ---------------------------------------------------------------------------
# Public pipeline
# ---------------------------------------------------------------------------

def process(packing_path: str, output_path: str, factura_path: Optional[str] = None,
            mapping_path: Optional[str] = None, inject: bool = True) -> List[Dict[str, Any]]:
    if not os.path.exists(packing_path):
        raise FileNotFoundError(f"Packing not found: {packing_path}")
    if factura_path and not os.path.exists(factura_path):
        raise FileNotFoundError(f"FACTURA not found: {factura_path}")

    lineas, sheet_name, layout, pallet_bruto = read_packing(packing_path)
    log.info("Parsed %d line groups from packing (pallet_bruto=%.2f kg)", len(lineas), pallet_bruto)

    # ── Compute gramaje (g/m²) for each line group ──────────────────────────
    for l in lineas:
        g = _calc_gramaje(
            neto_kg=l.get("neto") or 0.0,
            m2=l.get("m2") or 0.0,
            cant_total=l.get("cant_total") or 0.0,
            desc=l.get("descripcion") or "",
        )
        if g and 10.0 < g < 2000.0:   # sanity range
            l["gramaje"] = round(g, 1)
        else:
            l["gramaje"] = None

    rules = None
    if mapping_path and os.path.exists(mapping_path):
        csv_path = mapping_path
    else:
        # Auto-detect: look for any known mapping filename next to the script
        csv_path = next(
            (str(EXTRACTOR_DIR / name) for name in _MAPPING_CANDIDATES
             if (EXTRACTOR_DIR / name).exists()),
            None
        )
    if csv_path:
        rules = load_mapping(csv_path)
    else:
        log.info("No product mapping CSV provided/found. Continuing with built-in heuristics.")

    factura = load_factura(factura_path) if factura_path else {}
    issues: List[str] = []
    issues.extend(apply_classification(lineas, factura=factura, rules=rules))
    apply_manual_business_rules(lineas)

    if factura:
        issues.extend(enrich_valor_from_factura(lineas, factura))
    else:
        for l in lineas:
            if l.get("valor") is None:
                l["valor"] = 0.0
                l["valor_fuente"] = "NONE"

    summary = aggregate(lineas, pallet_bruto=pallet_bruto)

    ext = Path(output_path).suffix.lower()
    if inject:
        source_ext = Path(packing_path).suffix.lower()
        if source_ext == ".ods":
            inject_summary_into_ods(summary, packing_path, output_path)
        else:
            inject_summary_into_xlsx(summary, packing_path, output_path, issues=issues, detail=lineas)
    else:
        write_output(summary, output_path)

    log.info(
        "Done — %d partidas | BX=%d VALOR=%.2f EUR BRUTO=%.2f NETO=%.2f M2=%.2f",
        len(summary), sum(r["bx"] for r in summary), sum((r["valor"] or 0.0) for r in summary),
        sum(r["bruto"] for r in summary), sum(r["neto"] for r in summary), sum(r["m2"] for r in summary),
    )
    if issues:
        log.warning("Issues detected: %d", len(issues))
        for iss in issues:
            log.warning("  -> %s", iss)
    return summary


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Croton packing (.ods/.xlsx) + factura/mapeo (.ods/.xlsx/.xls) -> Resumen_Partidas")
    parser.add_argument("input_packing", help="Input packing file (.ods or .xlsx)")
    parser.add_argument("-o", "--output", dest="output_file", required=True, help="Output file path")
    parser.add_argument("--factura", default=None, help="Factura / mapeo de productos path (.ods, .xlsx or .xls)")
    parser.add_argument("--mapping", default=None, help="Optional product_mapping.csv path")
    parser.add_argument("--no-inject", action="store_true", default=False, help="Write standalone output instead of injecting into a copy of the source workbook")
    parser.add_argument("--inject", action="store_true", default=False, help="Legacy flag; injection is default")
    args = parser.parse_args()

    summary = process(
        packing_path=args.input_packing,
        output_path=args.output_file,
        factura_path=args.factura,
        mapping_path=args.mapping,
        inject=not args.no_inject,
    )
    print(f"Done — {len(summary)} partidas -> {args.output_file}")
    print(
        f"VALOR={sum((r.get('valor') or 0) for r in summary):.2f} EUR  "
        f"BRUTO={sum((r.get('bruto') or 0) for r in summary):.2f}kg  "
        f"NETO={sum((r.get('neto') or 0) for r in summary):.2f}kg  "
        f"M2={sum((r.get('m2') or 0) for r in summary):.2f}"
    )


if __name__ == "__main__":
    main()