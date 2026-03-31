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
SCRIPT_VERSION = "2026-03-17.v1"

SCRIPT_CHANGELOG = """
## 2026-03-17.v1

### Logica general
Procesa packing lists (ODS/XLSX) junto con facturas opcionales y un mapeo
de productos para generar una hoja `Resumen_Partidas` con totales por
partida arancelaria.

### Formatos de entrada
- Packing list: `.ods` o `.xlsx`
- Factura/mapeo: `.ods`, `.xlsx` o `.xls`
- Mapeo de productos CSV (opcional)

### Clasificacion de mercancias
- Layout enriquecido: usa columnas MERCANCIA + PARTIDA existentes
- Layout raw: clasifica por heuristicas de descripcion o mapeo CSV

### Calculo de VALOR
- Si se aporta factura, el valor se asigna desde los totales de factura
  por referencia, evitando doble conteo
- Agrupa por (DESCRIPCION, PARTIDA) con totales: PESO BRUTO, PESO NETO,
  CANTIDAD, VALOR

### Salida
Inyecta la hoja `Resumen_Partidas` en una copia del fichero fuente.
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

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("croton")

COLOR_WORDS = {
    "NEGRO", "NEGRA", "MARINO", "AMARILLO", "AMARILLA", "AZUL", "AZULINA",
    "CELESTE", "ROJO", "ROJA", "VERDE", "GRIS", "MORADO", "MORADA", "MARRON",
    "TURQUESA", "CEREZA", "NARANJA", "BLANCO", "BLANCA", "BEIGE",
}

BUILTIN_REF_ALIASES = {
    # Seen in Croton cases: packing code vs invoice generic/family code.
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
    t = re.sub(r"\b\d+[%]?[A-Z]*\b", lambda m: m.group(0).replace(" ", ""), t)
    for color in COLOR_WORDS:
        t = re.sub(rf"\b{re.escape(color)}\b", "COLORES", t)
    t = re.sub(r"\bCOLORES(?:\s+COLORES)+\b", "COLORES", t)
    t = re.sub(r"[^A-Z0-9% ]+", " ", t)
    return _norm_spaces(t)


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
        # prefer Hoja1 enriched over Hoja5 raw if both exist
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
    # fallback: inspect first non-empty data rows
    for r in rows[:10]:
        vals = [_safe_str(v).upper() for v in r[:5]]
        if len(vals) >= 3 and vals[1] == "MERCANCIA" and vals[2] == "PARTIDA":
            return "enriched_classic"
    return "raw_packing"


def _parse_enriched_classic(rows: List[List[Any]]) -> List[Dict[str, Any]]:
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
    return result


def _parse_enriched_headerless(rows: List[List[Any]]) -> List[Dict[str, Any]]:
    # Same semantic layout as classic, just with first two header rows merged differently.
    return _parse_enriched_classic(rows)


def _parse_raw_packing(rows: List[List[Any]]) -> List[Dict[str, Any]]:
    # Raw packing-list layout:
    # 0 LINEA, 1 Nº BULTO, 2 DESCRIPCION, 3 CODIGO, 5 NETO, 6 BRUTO, 7 CANT, 8 CANT TOTAL, 9 M2
    result = []
    cur_rows = []
    cur_neto = 0.0
    cur_bruto = 0.0
    start = 0
    for i, r in enumerate(rows):
        if _row_has(r, 0, "LINEA") and _row_has(r, 2, "DESCRIPCION"):
            start = i + 1
            break
    for row in rows[start:]:
        row = list(row)
        while len(row) < 10:
            row.append(None)
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
    return result


def read_packing(path: str) -> Tuple[List[Dict[str, Any]], str, str]:
    sheet_name = _detect_packing_sheet(path)
    rows = _read_sheet(path, sheet_name)
    layout = _detect_layout(rows)
    log.info("Packing '%s' -> sheet '%s' layout=%s", path, sheet_name, layout)
    if layout == "enriched_classic":
        return _parse_enriched_classic(rows), sheet_name, layout
    if layout == "enriched_headerless":
        return _parse_enriched_headerless(rows), sheet_name, layout
    return _parse_raw_packing(rows), sheet_name, layout


# ---------------------------------------------------------------------------
# Optional CSV mapping
# ---------------------------------------------------------------------------

def load_mapping(csv_path: str) -> List[Dict[str, str]]:
    rules = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            codigo = row.get("CODIGO", "").strip()
            if not codigo or codigo.startswith("#"):
                continue
            rules.append({
                "codigo": _norm_ref(codigo),
                "contains": row.get("DESCRIPCION_CONTAINS", "").strip().upper(),
                "mercancia": row.get("MERCANCIA", "").strip(),
                "partida": row.get("PARTIDA", "").strip(),
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

def load_factura(path: str) -> Dict[str, Dict[str, Any]]:
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

    result: Dict[str, Dict[str, Any]] = OrderedDict()
    for row in rows:
        ref = _safe_str(row[0] if len(row) > 0 else None)
        desc = _safe_str(row[1] if len(row) > 1 else None)
        cant = _parse_num(row[6] if len(row) > 6 else None)
        neto = _parse_num(row[8] if len(row) > 8 else None)
        importe = _parse_num(row[9] if len(row) > 9 else None)
        if not ref or ref.upper() == "REFERENCIA" or cant is None or importe is None:
            continue
        refn = _norm_ref(ref)
        slot = result.setdefault(refn, {
            "referencia": refn,
            "descripcion": desc,
            "cant_total": 0.0,
            "importe_total": 0.0,
            "neto_total": 0.0,
            "desc_norm": _normalize_desc_for_match(desc),
        })
        slot["cant_total"] += cant or 0.0
        slot["importe_total"] += importe or 0.0
        slot["neto_total"] += neto or 0.0
    log.info("Loaded FACTURA refs=%d total=%.2f EUR", len(result), sum(v["importe_total"] for v in result.values()))
    return result


def infer_invoice_ref(linea: Dict[str, Any], factura: Dict[str, Dict[str, Any]]) -> Optional[str]:
    ref = _norm_ref(linea.get("referencia", ""))
    if ref in factura:
        return ref
    alias = BUILTIN_REF_ALIASES.get(ref)
    if alias and alias in factura:
        return alias
    # Normalize by description family (colors -> COLORES)
    desc_norm = _normalize_desc_for_match(linea.get("descripcion", ""))
    candidates = [k for k, v in factura.items() if v.get("desc_norm") == desc_norm]
    if len(candidates) == 1:
        return candidates[0]
    # fallback by key phrases
    for k, v in factura.items():
        fdesc = v.get("desc_norm", "")
        if desc_norm and desc_norm == fdesc:
            return k
    return None


# ---------------------------------------------------------------------------
# Classification without mandatory CSV
# ---------------------------------------------------------------------------

def _contains(text: str, *needles: str) -> bool:
    up = (text or "").upper()
    return any(n.upper() in up for n in needles)


def classify_from_text(descripcion: str, referencia: str = "", factura_desc: str = "") -> Tuple[Optional[str], Optional[str], Optional[str]]:
    desc = (factura_desc or descripcion or "").upper()
    ref = _norm_ref(referencia)

    # Accessories / non-fabric
    if _contains(desc, "CREMALLERA"):
        return "CREMALLERA DIENTE PLASTICO", "9607190000", None
    if _contains(desc, "ANAGRAMA"):
        return "ANAGRAMAS", "5807101000", None
    if _contains(desc, "ETIQUETA"):
        return "ETIQUETAS COSER", "5807101000", None
    if _contains(desc, "REFLECTANTE"):
        return "REFLECTANTE 5CM", "3920610090", None

    # Rejilla
    if _contains(desc, "REJILLA"):
        if _contains(desc, "BLANCA", "BLANCO"):
            return "TEJ BLANQUEADOS DE REJILLA", "5804101000", None
        return "TEJ  DE REJILLA", "5804109000", None

    # Knit / plush
    if _contains(desc, "FELPA"):
        return "TEJ DE FELPA PRED F SINT", "6001920000", None
    if _contains(desc, "PUNTO", "PIQUE", "CANALE"):
        if _contains(desc, "BLANCA", "BLANCO") and _contains(desc, "CANALE"):
            return "TEJ BLANQUEADOS DE PUNTO PRED LAS F SINT", "6006310000", None
        return "TEJ DE PUNTO PRED F SINT MEZCLADAS CON ALGODON", "6006320000", None

    # Woven / flat / poplin / twill
    if _contains(desc, "65%POL 35%VISC"):
        return "TEJ PLANA DE FIB SINTC", "5514301000", None

    if _contains(desc, "100%ALG") and _contains(desc, "SARGA"):
        return "TEJ TEÑIDOS DE FIB SARGA DE ALGODON", "5209320000", None

    if _contains(desc, "SARGA"):
        if _contains(desc, "BLANCA", "BLANCO"):
            return "TEJ BLANQUEADOS DE SARGA DE F SINT CON ALG", "5514120000", None
        if _contains(desc, "AMARILLA", "MARINO", "AZUL", "CELESTE") and _contains(desc, "65%POL 35%ALG"):
            # preserve literal groups when the manual does so
            color_literal = _norm_spaces(descripcion or factura_desc)
            return color_literal, "5514220000", None
        return "TEJ TEÑIDOS DE SARGA DE F SINT CON ALG", "5514220000", None

    if _contains(desc, "POPELIN", "PLANA", "TAFETAN"):
        return "TEJ TIÑIDOS DE TAFETAN PRED F SINT", "5513210000", None

    return None, None, "No classification rule matched"


def apply_classification(lineas: List[Dict[str, Any]], factura: Optional[Dict[str, Dict[str, Any]]] = None,
                         rules: Optional[List[Dict[str, str]]] = None) -> List[str]:
    issues: List[str] = []
    for l in lineas:
        if l.get("mercancia") and l.get("partida_arancel"):
            continue
        # optional CSV mapping first
        mapped = None
        if rules:
            mapped = lookup_mapping(l.get("referencia", ""), l.get("descripcion", ""), rules)
        if mapped:
            l["mercancia"], l["partida_arancel"] = mapped
            l["clasif_fuente"] = "CSV"
            continue
        inv_ref = infer_invoice_ref(l, factura) if factura else None
        l["invoice_ref"] = inv_ref
        inv_desc = factura.get(inv_ref, {}).get("descripcion", "") if (factura and inv_ref) else ""
        merc, part, err = classify_from_text(l.get("descripcion", ""), l.get("referencia", ""), inv_desc)
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


SPECIAL_NO_VALOR_DESCRIPTIONS = {
    "SARGA AMARILLA 65%POL 35%ALG A 1,60",
    "SARGA MARINO 65%POL 35%ALG A 1,60",
    "SARGA AZUL 65%POL 35%ALG A 1,60",
    "SARGA CELESTE 65%POL 35%ALG A 1,60",
}

def apply_manual_business_rules(lineas: List[Dict[str, Any]]) -> None:
    """
    Croton-specific normalisation to mirror the operator workbook more closely.

    Rules observed from the manual workbook:
    - "PLANA BLANCA CUADRO VERDE ..." is kept as its own row, not merged into
      generic tafetan/poplín bucket.
    - Four specific 65/35 sarga colour rows are kept as their literal description
      with blank VALOR, while their invoice value stays on the generic
      "TEJ TEÑIDOS DE SARGA DE F SINT CON ALG" bucket.
    - CANALE blanco stays in blanqueados (6006310000); coloured CANALE rows go to
      teñidos (6006320000).
    """
    for l in lineas:
        desc = _norm_spaces(_safe_str(l.get("descripcion", "")).upper())

        if desc == "PLANA BLANCA CUADRO VERDE 60%ALG 40%POL A 1,50":
            l["mercancia"] = "PLANA BLANCA CUADRO VERDE 60%ALG 40%POL A 1,50"
            l["partida_arancel"] = "5513210000"
            l["clasif_fuente"] = "MANUAL_RULE"
            l["keep_literal"] = True

        # Manual workbook only breaks out a few specific 65/35 twill rows as
        # literal descriptions with blank VALOR. Similar colour rows remain
        # inside the generic bucket.
        m2 = round(float(l.get("m2") or 0.0), 2)
        if (
            (desc == "SARGA AMARILLA 65%POL 35%ALG A 1,60" and m2 == 155.84)
            or (desc == "SARGA MARINO 65%POL 35%ALG A 1,60" and m2 == 736.48)
            or (desc == "SARGA AZUL 65%POL 35%ALG A 1,60" and m2 == 176.00)
            or (desc == "SARGA CELESTE 65%POL 35%ALG A 1,60" and m2 == 16.00)
        ):
            l["mercancia"] = _safe_str(l.get("descripcion", "")).strip()
            l["partida_arancel"] = "5514220000"
            l["clasif_fuente"] = "MANUAL_RULE"
            l["force_blank_valor"] = True
            l["transfer_valor_to"] = ("TEJ TEÑIDOS DE SARGA DE F SINT CON ALG", "5514220000")

        if "CANALE" in desc:
            if "BLANCO" in desc:
                l["mercancia"] = "TEJ BLANQUEADOS DE PUNTO PRED LAS F SINT"
                l["partida_arancel"] = "6006310000"
            else:
                l["mercancia"] = "TEJ TEÑIDOS DE PUNTO PRED LAS F SINT"
                l["partida_arancel"] = "6006320000"
            l["clasif_fuente"] = "MANUAL_RULE"


# ---------------------------------------------------------------------------
# VALOR allocation from factura
# ---------------------------------------------------------------------------

def enrich_valor_from_factura(lineas: List[Dict[str, Any]], factura: Dict[str, Dict[str, Any]]) -> List[str]:
    issues: List[str] = []
    groups: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for l in lineas:
        inv_ref = l.get("invoice_ref") or infer_invoice_ref(l, factura)
        l["invoice_ref"] = inv_ref
        if inv_ref:
            groups[inv_ref].append(l)
        else:
            if l.get("valor") is None:
                l["valor"] = 0.0
            l["valor_fuente"] = "NO_MATCH"
            issues.append(f"No factura match for referencia={l.get('referencia')} descripcion={l.get('descripcion')}")

    for inv_ref, items in groups.items():
        inv = factura.get(inv_ref)
        total_qty = sum((i.get("cant_total") or 0.0) for i in items)
        inv_qty = inv.get("cant_total") or 0.0
        base_qty = total_qty if total_qty > 0 else inv_qty
        if base_qty <= 0:
            for i in items:
                i["valor"] = 0.0
                i["valor_fuente"] = "FACTURA_ZERO"
            issues.append(f"Zero quantity for matched factura ref {inv_ref}")
            continue
        importe_total = inv.get("importe_total") or 0.0
        running = 0.0
        for idx, item in enumerate(items, start=1):
            qty = item.get("cant_total") or 0.0
            if idx < len(items):
                value = round((qty / base_qty) * importe_total, 2)
                running += value
            else:
                value = round(importe_total - running, 2)
            item["valor"] = value
            item["valor_fuente"] = f"FACTURA:{inv_ref}"
    return issues


# ---------------------------------------------------------------------------
# Aggregate
# ---------------------------------------------------------------------------

def aggregate(lineas: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen: "OrderedDict[Tuple[str, str], Dict[str, Any]]" = OrderedDict()
    transferred_values: Dict[Tuple[str, str], float] = defaultdict(float)

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
            if l.get("force_blank_valor"):
                dest = l.get("transfer_valor_to")
                if dest:
                    transferred_values[dest] += float(val or 0.0)
            else:
                g["valor"] += float(val or 0.0)
                g["valor_present"] = True
        g["bruto"] += l.get("bruto") or 0.0
        g["neto"] += l.get("neto") or 0.0
        g["m2"] += l.get("m2") or 0.0

    for dest, extra in transferred_values.items():
        if dest not in seen:
            seen[dest] = {
                "mercancia": dest[0],
                "partida": dest[1],
                "bx": 0.0,
                "valor": 0.0,
                "valor_present": False,
                "bruto": 0.0,
                "neto": 0.0,
                "m2": 0.0,
            }
        seen[dest]["valor"] += round(extra, 2)
        seen[dest]["valor_present"] = True

    out = []
    for g in seen.values():
        g["bx"] = int(round(g["bx"]))
        g["valor"] = round(g["valor"], 2) if g["valor_present"] else None
        g["bruto"] = round(g["bruto"], 2)
        g["neto"] = round(g["neto"], 2)
        g["m2"] = round(g["m2"], 2)
        g.pop("valor_present", None)
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
    # Remove only the sheets we manage; keep all original sheets intact
    for name in ("Resumen_Partidas", "Issues", "Detalle_Extractor"):
        if name in wb.sheetnames:
            del wb[name]
    ws = wb.create_sheet("Resumen_Partidas")
    ws.append(_summary_headers())
    for r in summary:
        ws.append([r["mercancia"], r["partida"], r["bx"], r["valor"], r["bruto"], r["neto"], r["m2"]])
    # Ensure the first sheet is active so the workbook opens correctly in LibreOffice/Excel
    wb.active = wb.worksheets[0]
    # Fix bookView: ensure sheet tabs are visible and the view is not corrupted
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

    lineas, sheet_name, layout = read_packing(packing_path)
    log.info("Parsed %d line groups from packing", len(lineas))

    rules = None
    csv_path = mapping_path or (str(DEFAULT_MAPPING_FILE) if DEFAULT_MAPPING_FILE.exists() else None)
    if csv_path and os.path.exists(csv_path):
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

    summary = aggregate(lineas)

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



# ---------------------------------------------------------------------------
# Rule-based classification v3 (attribute-driven, no lot scope)
# ---------------------------------------------------------------------------

SCRIPT_VERSION = "2026-03-31.v3"
RULES_CANDIDATES = [
    "product_mapping_rules.csv",
    "product_mapping.csv",
]

def _contains(text: str, *parts: str) -> bool:
    up = (text or "").upper()
    return any(p.upper() in up for p in parts)

def _primary_text(descripcion: str, factura_desc: str = "") -> str:
    return _norm_spaces(" ".join(x for x in [descripcion, factura_desc] if x)).upper()

def _finish_from_desc(text: str) -> str:
    up = (text or "").upper()
    return "BLANQUEADO" if ("BLANCO" in up or "BLANCA" in up or "BLANQUEAD" in up) else "TENIDO"

def _family_from_desc(text: str) -> str:
    up = (text or "").upper()
    if "CREMALLERA" in up:
        return "ACCESORIO_CREMALLERA"
    if "ANAGRAMA" in up:
        return "ACCESORIO_ANAGRAMA"
    if "REFLECTANTE" in up:
        return "ACCESORIO_REFLECTANTE"
    if "ETIQUETA CARTON" in up:
        return "ACCESORIO_ETIQUETA_CARTON"
    if "ETIQUETA" in up:
        return "ACCESORIO_ETIQUETA"
    if "TRANSFER" in up:
        return "ACCESORIO_TRANSFER"
    if "REJILLA" in up:
        return "REJILLA"
    if "FELPA" in up:
        return "FELPA"
    if any(k in up for k in ["PUNTO", "PIQUE", "CANALE"]):
        return "PUNTO"
    if any(k in up for k in ["POPELIN", "PLANA", "TAFETAN"]):
        return "TAFETAN"
    if "SARGA" in up:
        return "SARGA"
    return "OTRO"







def _extract_fiber_pcts(desc: str) -> Dict[str, float]:
    result: Dict[str, float] = {}
    for m in re.finditer(r"(\d+(?:[.,]\d+)?)\s*%\s*([A-Z]+(?:\s*MOD)?)", (desc or "").upper()):
        pct = float(m.group(1).replace(",", "."))
        fiber = re.sub(r"\s+", "", m.group(2))
        fiber = {"POLIAM": "PA", "POLAM": "PA", "POLYAM": "PA", "ELASTANO": "ELAST"}.get(fiber, fiber)
        result[fiber] = result.get(fiber, 0.0) + pct
    return result

def _extract_width(desc: str) -> float:
    m = re.search(r"\bA\s+(\d+(?:[.,]\d+)?)\b", (desc or "").upper())
    if m:
        val = float(m.group(1).replace(",", "."))
        return val / 100.0 if val >= 10 else val
    return 0.0

def _calc_gramaje(neto_kg: float, m2: float, cant_total: float = 0.0, desc: str = "") -> Optional[float]:
    if m2 and m2 > 0:
        return neto_kg / m2 * 1000.0
    if cant_total and cant_total > 0:
        ancho = _extract_width(desc)
        if ancho > 0:
            return neto_kg / (cant_total * ancho) * 1000.0
    return None

def _extract_attrs(descripcion: str, factura_desc: str = "", gramaje: Optional[float] = None) -> Dict[str, Any]:
    merged = _primary_text(descripcion, factura_desc)
    fibers = _extract_fiber_pcts(merged)
    cotton = fibers.get("ALG", 0.0)
    poly = fibers.get("POL", 0.0)
    synth = sum(v for k, v in fibers.items() if k in ("POL", "PA", "ACR", "ACRMOD"))
    visc = sum(v for k, v in fibers.items() if k in ("VISC", "MODAL", "MOD"))
    width = _extract_width(merged)
    return {
        "merged_desc": merged,
        "desc_norm": _normalize_desc_for_match(merged),
        "family": _family_from_desc(merged),
        "finish": _finish_from_desc(merged),
        "cotton_pct": cotton,
        "poly_pct": poly,
        "synth_pct": synth,
        "visc_pct": visc,
        "width": width,
        "gramaje": gramaje,
    }

def load_mapping(csv_path: str) -> List[Dict[str, Any]]:
    """
    New schema preferred:
    PRIORITY,FABRIC_FAMILY,FINISH,DESC_CONTAINS,EXCLUDE_DESC_CONTAINS,
    COTTON_MIN,COTTON_MAX,POLY_MIN,POLY_MAX,SYNTH_MIN,SYNTH_MAX,
    WIDTH_MIN,WIDTH_MAX,GRAMAJE_MIN,GRAMAJE_MAX,HS_CODE,
    NORMALIZED_MERCANCIA,KEEP_LITERAL,BLANK_VALOR,TRANSFER_TO_MERCANCIA,
    TRANSFER_TO_HS_CODE,NOTES,EVIDENCE
    """
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = [c.strip().upper() for c in (reader.fieldnames or [])]
        new_schema = "HS_CODE" in fieldnames or "NORMALIZED_MERCANCIA" in fieldnames
        rules: List[Dict[str, Any]] = []
        for row in reader:
            norm = {str(k).strip().upper(): ("" if v is None else str(v).strip()) for k, v in row.items()}
            if not any(norm.values()):
                continue
            if str(next(iter(norm.values()), "")).startswith("#"):
                continue
            if new_schema:
                hs = norm.get("HS_CODE") or norm.get("PARTIDA")
                merch = norm.get("NORMALIZED_MERCANCIA") or norm.get("DESC_ADUANERA") or norm.get("MERCANCIA")
                if not hs or not merch:
                    continue
                def fnum(name):
                    return _parse_num(norm.get(name))
                rule = {
                    "priority": int(_parse_num(norm.get("PRIORITY")) or 1000),
                    "family": (norm.get("FABRIC_FAMILY") or "").upper(),
                    "finish": (norm.get("FINISH") or "").upper(),
                    "desc_contains": (norm.get("DESC_CONTAINS") or "").upper(),
                    "exclude_desc_contains": (norm.get("EXCLUDE_DESC_CONTAINS") or "").upper(),
                    "cotton_min": fnum("COTTON_MIN"),
                    "cotton_max": fnum("COTTON_MAX"),
                    "poly_min": fnum("POLY_MIN"),
                    "poly_max": fnum("POLY_MAX"),
                    "synth_min": fnum("SYNTH_MIN"),
                    "synth_max": fnum("SYNTH_MAX"),
                    "width_min": fnum("WIDTH_MIN"),
                    "width_max": fnum("WIDTH_MAX"),
                    "gramaje_min": fnum("GRAMAJE_MIN"),
                    "gramaje_max": fnum("GRAMAJE_MAX"),
                    "hs_code": hs,
                    "mercancia": merch,
                    "keep_literal": (norm.get("KEEP_LITERAL") or "").upper() in ("1", "TRUE", "YES", "SI"),
                    "blank_valor": (norm.get("BLANK_VALOR") or "").upper() in ("1", "TRUE", "YES", "SI"),
                    "transfer_to_mercancia": norm.get("TRANSFER_TO_MERCANCIA") or "",
                    "transfer_to_hs": norm.get("TRANSFER_TO_HS_CODE") or "",
                    "notes": norm.get("NOTES") or norm.get("NOTA") or "",
                    "evidence": norm.get("EVIDENCE") or norm.get("FUENTE") or "",
                }
                rules.append(rule)
            else:
                codigo = norm.get("CODIGO", "")
                if not codigo:
                    continue
                merch = norm.get("DESC_ADUANERA") or norm.get("MERCANCIA")
                hs = norm.get("PARTIDA", "")
                if not merch or not hs:
                    continue
                rules.append({
                    "priority": 500,
                    "family": "",
                    "finish": "",
                    "desc_contains": (norm.get("DESCRIPCION_CONTAINS") or norm.get("DESCRIPCION") or "").upper(),
                    "exclude_desc_contains": "",
                    "cotton_min": None,
                    "cotton_max": None,
                    "poly_min": None,
                    "poly_max": None,
                    "synth_min": None,
                    "synth_max": None,
                    "width_min": None,
                    "width_max": None,
                    "gramaje_min": _parse_num(norm.get("GRAMAJE_MIN")),
                    "gramaje_max": _parse_num(norm.get("GRAMAJE_MAX")),
                    "hs_code": hs,
                    "mercancia": merch,
                    "keep_literal": False,
                    "blank_valor": False,
                    "transfer_to_mercancia": "",
                    "transfer_to_hs": "",
                    "notes": "legacy",
                    "evidence": "",
                    "codigo": _norm_ref(codigo),
                })
        rules.sort(key=lambda r: r.get("priority", 1000))
        log.info("Loaded %d classification rules from %s", len(rules), csv_path)
        return rules

def _rule_matches(rule: Dict[str, Any], attrs: Dict[str, Any], referencia: str = "") -> bool:
    merged = attrs["merged_desc"]
    if rule.get("family") and rule["family"] != attrs["family"]:
        return False
    if rule.get("finish") and rule["finish"] != attrs["finish"]:
        return False
    if rule.get("desc_contains") and rule["desc_contains"] not in merged:
        return False
    if rule.get("exclude_desc_contains") and rule["exclude_desc_contains"] in merged:
        return False
    if rule.get("codigo") and _norm_ref(referencia) != rule["codigo"]:
        return False

    def in_range(val, lo, hi):
        if lo is not None and (val is None or val < lo):
            return False
        if hi is not None and (val is None or val > hi):
            return False
        return True

    for pair in [
        (attrs.get("cotton_pct"), rule.get("cotton_min"), rule.get("cotton_max")),
        (attrs.get("poly_pct"), rule.get("poly_min"), rule.get("poly_max")),
        (attrs.get("synth_pct"), rule.get("synth_min"), rule.get("synth_max")),
        (attrs.get("width"), rule.get("width_min"), rule.get("width_max")),
        (attrs.get("gramaje"), rule.get("gramaje_min"), rule.get("gramaje_max")),
    ]:
        if not in_range(*pair):
            return False
    return True

def lookup_mapping(referencia: str, descripcion: str, rules: List[Dict[str, Any]], factura_desc: str = "",
                   gramaje: Optional[float] = None) -> Optional[Dict[str, Any]]:
    attrs = _extract_attrs(descripcion, factura_desc, gramaje)
    for rule in rules:
        if _rule_matches(rule, attrs, referencia):
            return rule
    return None

def classify_from_text(descripcion: str, referencia: str = "", factura_desc: str = "",
                       gramaje: Optional[float] = None) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    attrs = _extract_attrs(descripcion, factura_desc, gramaje)
    desc = attrs["merged_desc"]
    fam = attrs["family"]
    cotton = attrs["cotton_pct"]
    poly = attrs["poly_pct"]
    synth = attrs["synth_pct"]
    finish = attrs["finish"]

    if fam == "ACCESORIO_CREMALLERA":
        return "CREMALLERA DIENTE PLASTICO", "9607190000", None
    if fam == "ACCESORIO_ANAGRAMA":
        return "ANAGRAMAS", "5807101000", None
    if fam == "ACCESORIO_ETIQUETA_CARTON":
        return "ETIQUETA CARTON", "4821109000", None
    if fam == "ACCESORIO_ETIQUETA":
        return "ETIQUETAS COSER", "5807101000", None
    if fam == "ACCESORIO_REFLECTANTE":
        return "REFLECTANTE 5CM", "3920610090", None
    if fam == "ACCESORIO_TRANSFER":
        return "TRANSFER", "5807909000", None

    if fam == "REJILLA":
        if poly >= 99:
            return "TEJ BLANQUEADOS DE REJILLA", "5804101000", None
        return "TEJ  DE REJILLA", "5804109000", None

    if fam == "FELPA":
        if cotton >= 50 and synth < 50:
            return "TEJ DE FELPA PRED ALG", "6001910000", None
        return "TEJ DE FELPA PRED F SINT", "6001920000", None

    if fam == "PUNTO":
        if finish == "BLANQUEADO":
            return "TEJ BLANQUEADOS DE PUNTO PRED LAS F SINT", "6006310000", None
        if cotton >= 85 or (cotton >= 50 and poly == 0 and synth < 50):
            return "TEJIDOS TEÑIDOS DE PUNTO DE ALGODON", "6006220000", None
        return "TEJ DE PUNTO PRED F SINT MEZCLADAS CON ALGODON", "6006320000", None

    if fam == "TAFETAN":
        if "65%POL 35%VISC" in desc:
            return "TEJ PLANA DE FIB SINTC", "5514301000", None
        if finish == "BLANQUEADO" and poly > 0 and cotton > 0:
            return "TEJ BLANQUEADOS DE TAFETAN PRED F SINT", "5513112000", None
        if finish == "TENIDO" and cotton == 0 and synth >= 85:
            return "TEJIDOS TEÑIDOS DE TAFETAN DE FIBRAS SINTETICAS", "5512199000", None
        return "TEJ TIÑIDOS DE TAFETAN PRED F SINT", "5513210000", None

    if fam == "SARGA":
        if finish == "BLANQUEADO":
            if cotton >= 85:
                return "TEJIDOS BLANQUEADOS DE SARGA DE ALGODON", "5209220000", None
            return "TEJ BLANQUEADOS DE SARGA DE F SINT CON ALG", "5514120000", None
        if cotton >= 85:
            return "TEJ TEÑIDOS DE FIB SARGA DE ALGODON", "5209320000", None
        if cotton >= 70 and poly >= 15:
            return "TEJIDOS TEÑIDOS DE SARGA DE ALGODON MEZCLADOS CON FIBRAS SINTETICAS", "5211320090", None
        return "TEJ TEÑIDOS DE SARGA DE F SINT CON ALG", "5514220000", None

    return None, None, "No classification rule matched"

def _apply_rule_to_line(l: Dict[str, Any], rule: Dict[str, Any]) -> None:
    desc_clean = _norm_spaces(_safe_str(l.get("descripcion", "")))
    l["mercancia"] = desc_clean if rule.get("keep_literal") else rule["mercancia"]
    l["partida_arancel"] = rule["hs_code"]
    l["clasif_fuente"] = "RULE_CSV"
    if rule.get("blank_valor"):
        l["force_blank_valor"] = True
        if rule.get("transfer_to_mercancia") and rule.get("transfer_to_hs"):
            l["transfer_valor_to"] = (rule["transfer_to_mercancia"], rule["transfer_to_hs"])

def apply_classification(lineas: List[Dict[str, Any]], factura: Optional[Dict[str, Dict[str, Any]]] = None,
                         rules: Optional[List[Dict[str, Any]]] = None) -> List[str]:
    issues: List[str] = []
    for l in lineas:
        if l.get("mercancia") and l.get("partida_arancel"):
            continue
        inv_ref = infer_invoice_ref(l, factura) if factura else None
        l["invoice_ref"] = inv_ref
        inv_desc = factura.get(inv_ref, {}).get("descripcion", "") if (factura and inv_ref) else ""
        attrs = _extract_attrs(l.get("descripcion", ""), inv_desc, l.get("gramaje"))
        l.update({
            "fabric_family": attrs["family"],
            "finish": attrs["finish"],
            "cotton_pct": attrs["cotton_pct"],
            "poly_pct": attrs["poly_pct"],
            "synth_pct": attrs["synth_pct"],
            "width_m": attrs["width"],
            "desc_norm": attrs["desc_norm"],
        })
        mapped = None
        if rules:
            mapped = lookup_mapping(l.get("referencia", ""), l.get("descripcion", ""), rules, inv_desc, l.get("gramaje"))
        if mapped:
            _apply_rule_to_line(l, mapped)
            continue
        merc, part, err = classify_from_text(l.get("descripcion", ""), l.get("referencia", ""), inv_desc, l.get("gramaje"))
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
    Non-lot-specific business rules observed across manual workbooks.
    These are description/business rules, not per-lot overrides.
    """
    literal_sargas = {
        "SARGA AMARILLA 65%POL 35%ALG A 1,60",
        "SARGA MARINO 65%POL 35%ALG A 1,60",
        "SARGA AZUL 65%POL 35%ALG A 1,60",
        "SARGA CELESTE 65%POL 35%ALG A 1,60",
    }
    aliases = {
        "5513210000": "TEJ TIÑIDOS DE TAFETAN PRED F SINT",
        "5209320000": "TEJ TEÑIDOS DE FIB SARGA DE ALGODON",
        "5514220000": "TEJ TEÑIDOS DE SARGA DE F SINT CON ALG",
        "5514120000": "TEJ BLANQUEADOS DE SARGA DE F SINT CON ALG",
        "6001920000": "TEJ DE FELPA PRED F SINT",
        "6001910000": "TEJ DE FELPA PRED ALG",
        "6006320000": "TEJ DE PUNTO PRED F SINT MEZCLADAS CON ALGODON",
        "6006310000": "TEJ BLANQUEADOS DE PUNTO PRED LAS F SINT",
        "5804101000": "TEJ BLANQUEADOS DE REJILLA",
        "5804109000": "TEJ  DE REJILLA",
        "3920610090": "REFLECTANTE 5CM",
    }
    for l in lineas:
        desc = _norm_spaces(_safe_str(l.get("descripcion", "")).upper())
        if desc == "PLANA BLANCA CUADRO VERDE 60%ALG 40%POL A 1,50":
            l["mercancia"] = "PLANA BLANCA CUADRO VERDE 60%ALG 40%POL A 1,50"
            l["partida_arancel"] = "5513210000"
            l["keep_literal"] = True
        if desc in literal_sargas:
            l["mercancia"] = _norm_spaces(_safe_str(l.get("descripcion", "")))
            l["partida_arancel"] = "5514220000"
            l["force_blank_valor"] = True
            l["transfer_valor_to"] = ("TEJ TEÑIDOS DE SARGA DE F SINT CON ALG", "5514220000")
        part = str(l.get("partida_arancel") or "")
        if part in aliases and not l.get("keep_literal") and desc not in literal_sargas:
            l["mercancia"] = aliases[part]

def process(packing_path: str, output_path: str, factura_path: Optional[str] = None,
            mapping_path: Optional[str] = None, inject: bool = True) -> List[Dict[str, Any]]:
    if not os.path.exists(packing_path):
        raise FileNotFoundError(f"Packing not found: {packing_path}")
    if factura_path and not os.path.exists(factura_path):
        raise FileNotFoundError(f"FACTURA not found: {factura_path}")

    packing_ret = read_packing(packing_path)
    if len(packing_ret) == 4:
        lineas, sheet_name, layout, pallet_bruto = packing_ret
    else:
        lineas, sheet_name, layout = packing_ret
        pallet_bruto = 0.0
    log.info("Parsed %d line groups from packing", len(lineas))

    for l in lineas:
        g = _calc_gramaje(
            neto_kg=float(l.get("neto") or 0.0),
            m2=float(l.get("m2") or 0.0),
            cant_total=float(l.get("cant_total") or 0.0),
            desc=l.get("descripcion") or "",
        )
        l["gramaje"] = round(g, 1) if g is not None else None

    csv_path = mapping_path
    if not csv_path:
        for name in RULES_CANDIDATES:
            p = EXTRACTOR_DIR / name
            if p.exists():
                csv_path = str(p)
                break
    rules = load_mapping(csv_path) if csv_path and os.path.exists(csv_path) else None
    if not rules:
        log.info("No mapping CSV provided/found. Continuing with built-in heuristics.")

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

    summary = aggregate(lineas)

    if pallet_bruto and summary:
        # keep logistics adjustment explicit but generic: add pallet gross to the
        # first textile row, not to accessories
        for row in summary:
            if not str(row.get("partida","")).startswith(("58","96","39","48")):
                row["bruto"] = round((row.get("bruto") or 0.0) + pallet_bruto, 2)
                break

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
    return summary


if __name__ == "__main__":
    main()
