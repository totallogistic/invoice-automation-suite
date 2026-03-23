#!/usr/bin/env python3
"""
camion_export_processor.py – Camión Export unified processor
=====================================================
Workflow integrado:
  1. Conserva la hoja original del XLSX del cliente.
  2. Si se pasa --doc, añade una hoja PDF_VALIDADO con la validación XLSX vs PDF.
  3. Procesa los T1 y añade una hoja PACKING_LIST_RESULT con el resumen final.
  4. Añade una hoja T1_SUMMARY con lo extraído de los PDFs T1.

Uso:
    # Modo terminal
    python camion_export_processor.py --xlsx PL.xlsx --t1 T1_1.pdf T1_2.pdf --doc DOC.pdf --output PL-PROCESSED.xlsx

    # Modo web (llamado por el stack)
    python camion_export_processor.py --xlsx PL.xlsx --t1 T1_1.pdf T1_2.pdf --doc DOC.pdf -o /data/camion/out/batch_id

Si no se pasa --output/-o, genera automáticamente:
    <stem-del-xlsx>-PROCESSED.xlsx
"""

from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
from datetime import datetime
import math
import re
import unicodedata
from collections import defaultdict
from pathlib import Path
from typing import Iterable, Optional

import fitz  # PyMuPDF
import openpyxl
import pdfplumber
import pytesseract
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image
from pypdf import PdfReader

# ============================================================================
# Shared styling
# ============================================================================

MISSING_FILL = PatternFill(fill_type="solid", fgColor="FFF2F2")
CONFIRMED_FILL = PatternFill(fill_type="solid", fgColor="F2FFF2")

_GREY_HEADER = 'FFF0F0F0'
_GREEN_ROW   = 'FF66FF66'
_CYAN_TOTALS = 'FFCCFFFF'
_YELLOW_DATA = 'FFFFFFBB'


def _fill(hex_rgb: str) -> PatternFill:
    return PatternFill('solid', start_color=hex_rgb, end_color=hex_rgb)


# ============================================================================
# Validator logic (XLSX vs DOC PDF)
# ============================================================================


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip().lower()
    return text


def normalize_token(value: object) -> str:
    text = normalize_text(value)
    return re.sub(r"[^a-z0-9]", "", text)


def float_candidates(value: Optional[float]) -> list[str]:
    if value is None:
        return []
    candidates = set()
    v = float(value)
    candidates.add(str(int(round(v))) if v.is_integer() else f"{v:.2f}")
    candidates.add(f"{v:.2f}")
    candidates.add(f"{v:.1f}")
    for item in list(candidates):
        candidates.add(item.replace(".", ","))
        candidates.add(item.replace(",", "."))
        candidates.add(item.replace(".", ""))
        candidates.add(item.replace(",", ""))
    return sorted(x for x in candidates if x)


def find_present_values(haystack: str, needles: Iterable[str]) -> list[str]:
    found = []
    for n in needles:
        if n and n in haystack:
            found.append(n)
    return found


def safe_float(value: object) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except Exception:
        return None


@dataclasses.dataclass
class RowEntry:
    excel_row: int
    tour: Optional[str]
    date: Optional[str]
    trailer: Optional[str]
    to_code: Optional[str]
    shipper_name: Optional[str]
    shipper_iso: Optional[str]
    shipper_city: Optional[str]
    shipper_code: Optional[str]
    recipient_name: Optional[str]
    recipient_iso: Optional[str]
    recipient_city: Optional[str]
    recipient_code: Optional[str]
    vol: Optional[float]
    mrn_invoice: Optional[str]
    mrn_detail: Optional[str]
    value_eur: Optional[float]
    value_usd: Optional[float]
    hu: Optional[float]
    peso_bruto: Optional[float]
    peso_neto: Optional[float]
    pk: Optional[float]
    cl: Optional[float]

    @property
    def invoice_tokens(self) -> list[str]:
        tokens: list[str] = []
        for raw in [self.mrn_invoice, self.mrn_detail]:
            if not raw:
                continue
            for part in re.split(r"[\n,/;]+", str(raw)):
                part = part.strip()
                if not part:
                    continue
                part = re.sub(r"^(EX|T1)\s*:\s*", "", part, flags=re.I)
                token = normalize_token(part)
                if token:
                    tokens.append(token)
        return sorted(set(tokens))

    @property
    def shipper_tokens(self) -> list[str]:
        raw_parts = [self.shipper_name, self.shipper_city, self.shipper_code, self.shipper_iso]
        tokens = []
        for raw in raw_parts:
            tok = normalize_token(raw)
            if tok and len(tok) >= 2:
                tokens.append(tok)
        return sorted(set(tokens))

    @property
    def recipient_tokens(self) -> list[str]:
        raw_parts = [self.recipient_name, self.recipient_city, self.recipient_code, self.recipient_iso]
        tokens = []
        for raw in raw_parts:
            tok = normalize_token(raw)
            if tok and len(tok) >= 2:
                tokens.append(tok)
        return sorted(set(tokens))


@dataclasses.dataclass
class PageMatch:
    page: int
    score: int
    reasons: list[str]
    matched_tokens: dict[str, list[str] | bool | int]


def load_entries_from_xlsx(xlsx_path: Path) -> list[RowEntry]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws = wb[wb.sheetnames[0]]

    entries: list[RowEntry] = []
    current_tour = None
    current_date = None
    current_trailer = None

    for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
        if not any(cell is not None and cell != "" for cell in row):
            continue

        tour, date_val, trailer = row[0], row[1], row[2]
        if tour:
            current_tour = str(tour)
        if date_val:
            if isinstance(date_val, (dt.datetime, dt.date)):
                current_date = date_val.strftime("%Y-%m-%d")
            else:
                current_date = str(date_val)
        if trailer:
            current_trailer = str(trailer)

        entries.append(RowEntry(
            excel_row=row_idx,
            tour=current_tour,
            date=current_date,
            trailer=current_trailer,
            to_code=str(row[3]).strip() if row[3] else None,
            shipper_name=str(row[4]).strip() if row[4] else None,
            shipper_iso=str(row[5]).strip() if row[5] else None,
            shipper_city=str(row[6]).strip() if row[6] else None,
            shipper_code=str(row[7]).strip() if row[7] else None,
            recipient_name=str(row[8]).strip() if row[8] else None,
            recipient_iso=str(row[9]).strip() if row[9] else None,
            recipient_city=str(row[10]).strip() if row[10] else None,
            recipient_code=str(row[11]).strip() if row[11] else None,
            vol=safe_float(row[12]),
            mrn_invoice=str(row[13]).strip() if row[13] else None,
            mrn_detail=str(row[14]).strip() if row[14] else None,
            value_eur=safe_float(row[15]),
            value_usd=safe_float(row[16]),
            hu=safe_float(row[17]),
            peso_bruto=safe_float(row[18]),
            peso_neto=safe_float(row[19]),
            pk=safe_float(row[20]),
            cl=safe_float(row[21]),
        ))

    return entries


def extract_pdf_text(pdf_path: Path, dpi: int = 150) -> dict[int, str]:
    page_texts: dict[int, str] = {}

    reader = PdfReader(str(pdf_path))
    native_texts: dict[int, str] = {}
    for i, page in enumerate(reader.pages, start=1):
        try:
            native_texts[i] = page.extract_text() or ""
        except Exception:
            native_texts[i] = ""

    doc = fitz.open(str(pdf_path))
    for page_no in range(1, len(doc) + 1):
        native = native_texts.get(page_no, "")
        if len(normalize_text(native)) >= 80:
            page_texts[page_no] = native
            continue

        txt_path = cache_dir / f"page-{page_no:03d}.txt"
        if txt_path.exists():
            page_texts[page_no] = txt_path.read_text(encoding="utf-8", errors="ignore")
            continue

        page = doc.load_page(page_no - 1)
        scale = dpi / 72.0
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        ocr_text = pytesseract.image_to_string(img, lang="eng")
        txt_path.write_text(ocr_text, encoding="utf-8")
        page_texts[page_no] = ocr_text

    return page_texts


def score_entry_on_page(entry: RowEntry, page_text: str) -> PageMatch | None:
    tok = normalize_token(page_text)
    score = 0
    reasons: list[str] = []
    matched: dict[str, list[str] | bool | int] = {}

    to_token = normalize_token(entry.to_code)
    if to_token and to_token in tok:
        score += 50
        reasons.append("TO")
        matched["to"] = True

    invoice_hits = find_present_values(tok, entry.invoice_tokens)
    if invoice_hits:
        score += 35 + 10 * min(len(invoice_hits), 3)
        reasons.append("invoice/mrn")
        matched["invoice_tokens"] = invoice_hits

    shipper_hits = [s for s in entry.shipper_tokens if len(s) >= 4 and s in tok]
    if shipper_hits:
        score += 8 * min(len(shipper_hits), 4)
        reasons.append("shipper")
        matched["shipper_tokens"] = shipper_hits

    recipient_hits = [s for s in entry.recipient_tokens if len(s) >= 4 and s in tok]
    if recipient_hits:
        score += 6 * min(len(recipient_hits), 4)
        reasons.append("recipient")
        matched["recipient_tokens"] = recipient_hits

    for key, value, pts in [
        ("gross", entry.peso_bruto, 8),
        ("net", entry.peso_neto, 8),
        ("hu", entry.hu, 4),
        ("vol", entry.vol, 5),
        ("pk", entry.pk, 4),
        ("cl", entry.cl, 4),
    ]:
        hits = find_present_values(tok, [normalize_token(x) for x in float_candidates(value)])
        if hits:
            score += pts
            reasons.append(key)
            matched[key] = hits

    value_hits = []
    for val in [entry.value_eur, entry.value_usd]:
        value_hits.extend(find_present_values(tok, [normalize_token(x) for x in float_candidates(val)]))
    if value_hits:
        score += 10
        reasons.append("value")
        matched["value"] = sorted(set(value_hits))

    if score == 0:
        return None

    return PageMatch(page=-1, score=score, reasons=reasons, matched_tokens=matched)


def find_best_pages(entry: RowEntry, page_texts: dict[int, str], top_n: int = 6) -> list[PageMatch]:
    matches: list[PageMatch] = []
    for page_no, text in page_texts.items():
        pm = score_entry_on_page(entry, text)
        if pm is None:
            continue
        pm.page = page_no
        matches.append(pm)
    matches.sort(key=lambda x: (-x.score, x.page))
    return matches[:top_n]


def infer_status(entry: RowEntry, matches: list[PageMatch]) -> tuple[str, list[str]]:
    issues: list[str] = []
    if not matches:
        return "NO_MATCH", ["No se localizaron paginas candidatas"]

    best = matches[0]
    if best.score < 35:
        issues.append(f"Score bajo: {best.score}")

    has_to = bool(best.matched_tokens.get("to")) if entry.to_code else False
    has_invoice = bool(best.matched_tokens.get("invoice_tokens")) if entry.invoice_tokens else False

    if entry.to_code and not has_to:
        issues.append("No aparece el TO en la mejor pagina candidata")
    if entry.invoice_tokens and not has_invoice:
        issues.append("No aparece invoice/MRN detail en la mejor pagina candidata")

    return ("REVIEW", issues) if issues else ("OK", [])


def gather_context_text(matches: list[PageMatch], page_texts: dict[int, str], extra_window: int = 1) -> tuple[str, list[int]]:
    if not matches:
        return "", []
    page_set = set()
    for m in matches[:3]:
        for p in range(max(1, m.page - extra_window), min(max(page_texts), m.page + extra_window) + 1):
            page_set.add(p)
    pages = sorted(page_set)
    text = "\n\n".join(page_texts[p] for p in pages)
    return normalize_token(text), pages


def field_confirmation(entry: RowEntry, matches: list[PageMatch], page_texts: dict[int, str]) -> dict[str, bool | list[int]]:
    ctx_token, ctx_pages = gather_context_text(matches, page_texts)

    def token_present(value: object) -> bool:
        tok = normalize_token(value)
        return bool(tok) and tok in ctx_token

    def any_float_present(value: Optional[float]) -> bool:
        if value is None:
            return False
        candidates = [normalize_token(x) for x in float_candidates(value)]
        return any(c and c in ctx_token for c in candidates)

    mrn_invoice_ok = all(normalize_token(t) in ctx_token for t in entry.invoice_tokens[:1]) if entry.mrn_invoice else False
    mrn_detail_ok = all(normalize_token(t) in ctx_token for t in entry.invoice_tokens[1:]) if entry.mrn_detail else False

    return {
        "context_pages": ctx_pages,
        "A": False,
        "B": False,
        "C": False,
        "D": token_present(entry.to_code),
        "E": token_present(entry.shipper_name),
        "F": token_present(entry.shipper_iso),
        "G": token_present(entry.shipper_city),
        "H": token_present(entry.shipper_code),
        "I": token_present(entry.recipient_name),
        "J": token_present(entry.recipient_iso),
        "K": token_present(entry.recipient_city),
        "L": token_present(entry.recipient_code),
        "M": any_float_present(entry.vol),
        "N": mrn_invoice_ok,
        "O": mrn_detail_ok,
        "P": any_float_present(entry.value_eur),
        "Q": any_float_present(entry.value_usd),
        "R": any_float_present(entry.hu),
        "S": any_float_present(entry.peso_bruto),
        "T": any_float_present(entry.peso_neto),
        "U": any_float_present(entry.pk),
        "V": any_float_present(entry.cl),
    }


def build_report(entries: list[RowEntry], page_texts: dict[int, str]) -> dict:
    results = []
    ok = review = no_match = 0

    for entry in entries:
        matches = find_best_pages(entry, page_texts)
        status, issues = infer_status(entry, matches)
        confirmations = field_confirmation(entry, matches, page_texts) if matches else {}
        if status == "OK":
            ok += 1
        elif status == "REVIEW":
            review += 1
        else:
            no_match += 1

        results.append({
            "excel_row": entry.excel_row,
            "tour": entry.tour,
            "date": entry.date,
            "trailer": entry.trailer,
            "to": entry.to_code,
            "mrn_invoice": entry.mrn_invoice,
            "mrn_detail": entry.mrn_detail,
            "shipper_name": entry.shipper_name,
            "peso_bruto": entry.peso_bruto,
            "peso_neto": entry.peso_neto,
            "value_eur": entry.value_eur,
            "value_usd": entry.value_usd,
            "status": status,
            "issues": issues,
            "best_pages": [m.page for m in matches],
            "best_match": dataclasses.asdict(matches[0]) if matches else None,
            "candidates": [dataclasses.asdict(m) for m in matches],
            "confirmations": confirmations,
        })

    return {
        "summary": {
            "entries": len(entries),
            "ok": ok,
            "review": review,
            "no_match": no_match,
            "pdf_pages": len(page_texts),
        },
        "results": results,
    }


def add_validated_sheet(wb: openpyxl.Workbook, report: dict, validated_sheet_name: str = "PDF_VALIDADO") -> None:
    src_ws = _find_source_worksheet(wb)

    if validated_sheet_name in wb.sheetnames:
        del wb[validated_sheet_name]
    dst_ws = wb.copy_worksheet(src_ws)
    dst_ws.title = validated_sheet_name

    row_map = {item["excel_row"]: item for item in report["results"]}
    review_col = 23  # W
    pages_col = 24   # X

    dst_ws.cell(1, review_col).value = "Validation Status"
    dst_ws.cell(1, pages_col).value = "PDF Pages"
    dst_ws.cell(2, review_col).value = ""
    dst_ws.cell(2, pages_col).value = ""

    for row_idx in range(5, dst_ws.max_row + 1):
        item = row_map.get(row_idx)
        if not item:
            continue
        conf = item.get("confirmations", {})

        for col_idx in range(1, 23):
            col_letter = get_column_letter(col_idx)
            cell = dst_ws.cell(row_idx, col_idx)
            original_has_value = cell.value not in (None, "")
            keep = bool(conf.get(col_letter, False))

            if original_has_value and not keep:
                cell.value = None
                cell.fill = MISSING_FILL
            elif original_has_value and keep:
                cell.fill = CONFIRMED_FILL

        dst_ws.cell(row_idx, review_col).value = item.get("status")
        dst_ws.cell(row_idx, pages_col).value = ", ".join(map(str, conf.get("context_pages", [])))


# ============================================================================
# Existing T1 processor logic
# ============================================================================


def _parse_european_number(s: str) -> float:
    s = s.strip()
    if ',' in s:
        return float(s.replace('.', '').replace(',', '.'))
    parts = s.split('.')
    if len(parts) == 2 and len(parts[1]) == 3:
        return float(s.replace('.', ''))
    return float(s)



def _safe_idx(row, idx, default=None):
    if row is None:
        return default
    if idx < 0 or idx >= len(row):
        return default
    return row[idx]

def extract_t1_info(pdf_path: str) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text() or ""

    mrn_m = re.search(r'(\d{2}[A-Z]{2}\d{12}[A-Z0-9]+)', text)
    mrn = mrn_m.group(1) if mrn_m else None

    gross_m = re.search(r'[A-Z ]+SL\s+(\d+)\s+(\d+)\s+([\d.,]+)\s+\d', text)
    if gross_m:
        items = int(gross_m.group(1))
        packages = int(gross_m.group(2))
        gross_kg = _parse_european_number(gross_m.group(3))
    else:
        items = packages = 0
        gross_kg = 0.0

    deadline_m = re.search(r'Plazo.*?(\d{2}-\d{2}-\d{4})', text)
    deadline = deadline_m.group(1) if deadline_m else None

    return {
        'mrn': mrn,
        'gross_kg': gross_kg,
        'packages': packages,
        'items': items,
        'deadline': deadline,
        'source_file': Path(pdf_path).name,
    }


_COL = {
    'tour': 0,
    'date': 1,
    'trailer': 2,
    'to': 3,
    'shipper_name': 4,
    'shipper_iso': 5,
    'shipper_city': 6,
    'shipper_code': 7,
    'recipient_name': 8,
    'recipient_iso': 9,
    'recipient_city': 10,
    'recipient_code': 11,
    'vol': 12,
    'mrn_invoice': 13,
    'mrn_detail': 14,
    'value_eur': 15,
    'value_usd': 16,
    'hu': 17,
    'peso_bruto': 18,
    'peso_neto': 19,
    'pk': 20,
    'cl': 21,
}


def _find_source_worksheet(wb: openpyxl.Workbook) -> openpyxl.worksheet.worksheet.Worksheet:
    """
    Devuelve la primera hoja que parece ser el packing list del cliente,
    sin depender del nombre de la hoja.
    """
    for ws in wb.worksheets:
        raw = list(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 6), values_only=True))
        if len(raw) < 4:
            continue
        try:
            for i in range(2, min(6, len(raw))):
                row = raw[i]
                if row is None:
                    continue
                if len(row) <= _COL['cl']:
                    continue
                if row[_COL['peso_bruto']] is not None:
                    return ws
        except Exception:
            continue
    raise ValueError('No se encontró ninguna hoja con formato válido de packing list')


def read_sheet1(xlsx_path: str) -> tuple[list[dict], dict]:
    wb = openpyxl.load_workbook(xlsx_path)
    ws = _find_source_worksheet(wb)
    raw_rows = list(ws.iter_rows(values_only=False))
    raw = [tuple(c.value for c in r) for r in raw_rows]

    summary_row_idx = next(i for i in range(2, 6) if raw[i][_COL['peso_bruto']] is not None)
    s = raw[summary_row_idx]
    summary = {
        'total_peso_bruto': s[_COL['peso_bruto']],
        'total_peso_neto': s[_COL['peso_neto']],
        'total_pk': s[_COL['pk']],
        'total_cl': s[_COL['cl']],
    }

    rows = []
    for src_row_obj, raw_row in zip(raw_rows[summary_row_idx + 1:], raw[summary_row_idx + 1:]):
        if all(v is None for v in raw_row):
            continue
        has_yellow = any(c.fill.fgColor.rgb == 'FFFFFFBB' for c in src_row_obj if c.fill)
        row_dict = {k: raw_row[i] for k, i in _COL.items()}
        row_dict['_src_yellow'] = has_yellow
        rows.append(row_dict)

    return rows, summary


def _row_type(row: dict) -> str:
    mrn = str(row.get('mrn_invoice') or '')
    if mrn.startswith('T1:'):
        return 'T1'
    if mrn.startswith('EX:'):
        return 'EX'
    if row.get('to') is None and row.get('shipper_name') is None:
        return 'orphan'
    return 'regular'


def _group_key(row: dict) -> str:
    mrn = str(row.get('mrn_invoice') or '')
    rtype = row.get('_type', _row_type(row))
    if rtype == 'T1':
        return mrn
    if rtype == 'EX':
        return 'EX'
    if rtype == 'orphan':
        return 'ORPHAN'
    return 'REGULAR'


def process_packing_list(rows: list[dict], t1_map: dict) -> list[dict]:
    for row in rows:
        row['_type'] = _row_type(row)

    group_id = 0
    parent_group_id = {}
    for idx, row in enumerate(rows):
        if row['_type'] != 'orphan':
            group_id += 1
        parent_group_id[idx] = group_id

    orphan_weight_by_group: dict[int, float] = defaultdict(float)
    for idx, row in enumerate(rows):
        if row['_type'] == 'orphan':
            orphan_weight_by_group[parent_group_id[idx]] += row['peso_bruto'] or 0.0

    separator_before: set[int] = set()
    prev_gkey = None
    prev_type = None
    for idx, row in enumerate(rows):
        gkey = _group_key(row)
        rtype = row['_type']
        if prev_gkey is not None and gkey != prev_gkey:
            skip = rtype == 'orphan' or (prev_type == 'regular' and rtype in ('EX', 'regular'))
            if not skip:
                separator_before.add(idx)
        prev_gkey = gkey
        prev_type = rtype

    t1_mrn_seen: set[str] = set()
    result: list[dict] = []

    for idx, row in enumerate(rows):
        rtype = row['_type']
        gid = parent_group_id[idx]
        peso_bruto = row['peso_bruto']

        if rtype == 'T1':
            t1_mrn = str(row['mrn_invoice']).replace('T1:', '').strip()
            if t1_mrn not in t1_mrn_seen:
                t1_mrn_seen.add(t1_mrn)
                gross = t1_map.get(t1_mrn)
                peso_bruto = int(gross) if gross is not None else peso_bruto
            else:
                peso_bruto = None
        elif rtype == 'EX':
            if peso_bruto is not None:
                peso_bruto = math.ceil(peso_bruto)
        elif rtype == 'orphan':
            peso_bruto = None
        elif rtype == 'regular':
            extra = orphan_weight_by_group.get(gid, 0.0)
            if extra and peso_bruto is not None:
                peso_bruto = int(peso_bruto + extra)

        result.append({
            'tour': row['tour'],
            'date': row['date'],
            'trailer': row['trailer'],
            'shipper_name': row['shipper_name'],
            'mrn_invoice': row['mrn_invoice'],
            'peso_bruto': peso_bruto,
            'peso_neto': row['peso_neto'],
            'pk': row['pk'],
            'cl': row['cl'],
            '_type': rtype,
            '_src_yellow': row.get('_src_yellow', False),
            '_separator_before': idx in separator_before,
        })

    return result


def add_processed_sheet(wb: openpyxl.Workbook, result_rows: list[dict], summary: dict, processed_sheet_name: str = 'PACKING_LIST_RESULT') -> None:
    if processed_sheet_name in wb.sheetnames:
        del wb[processed_sheet_name]
    ws = wb.create_sheet(processed_sheet_name)

    arial10 = Font(name='Arial', size=10)
    arial10_bold = Font(name='Arial', size=10, bold=True)

    headers = ['Tour', 'Date', 'Trailer', 'Shipper', 'MRN / Invoice', 'Peso Bruto', 'Peso Neto', 'PK', 'CL']

    ws.append(headers)
    for cell in ws[1]:
        cell.font = arial10_bold
        cell.fill = _fill(_GREY_HEADER)
        cell.alignment = Alignment(horizontal='center')

    ws.append([None, None, None, 'Name', None, None, None, None, None])
    for cell in ws[2]:
        cell.font = arial10
        cell.fill = _fill(_GREY_HEADER)

    ws.append([None] * 9)
    for cell in ws[3]:
        cell.fill = _fill(_GREEN_ROW)

    ws.append([None, None, None, None, None, summary['total_peso_bruto'], summary['total_peso_neto'], summary['total_pk'], summary['total_cl']])
    for cell in ws[4]:
        cell.font = arial10_bold
        cell.fill = _fill(_CYAN_TOTALS)

    data_start_row = ws.max_row + 1
    for row in result_rows:
        if row['_separator_before']:
            sep_idx = ws.max_row + 1
            ws.append([None] * 9)
            for cell in ws[sep_idx]:
                cell.fill = _fill(_YELLOW_DATA)

        excel_row_idx = ws.max_row + 1
        ws.append([
            row['tour'], row['date'], row['trailer'], row['shipper_name'], row['mrn_invoice'],
            row['peso_bruto'], row['peso_neto'], row['pk'], row['cl'],
        ])

        row_fill = _fill(_YELLOW_DATA) if row['_src_yellow'] else None
        for cell in ws[excel_row_idx]:
            cell.font = arial10
            if row_fill:
                cell.fill = row_fill

        date_cell = ws.cell(excel_row_idx, 2)
        if date_cell.value and isinstance(date_cell.value, datetime):
            date_cell.number_format = 'DD.MM.YYYY'

    data_end_row = ws.max_row

    sep_idx = ws.max_row + 1
    ws.append([None] * 9)
    for cell in ws[sep_idx]:
        cell.fill = _fill(_YELLOW_DATA)

    coincide_row = ws.max_row + 1
    ws.append([
        None, None, None, None, None,
        f'=SUM(F{data_start_row}:F{data_end_row})',
        'COINCIDE ± CON PESO BRUTO',
        None,
        '=SUM(H4,I4)',
    ])
    for cell in ws[coincide_row]:
        cell.font = arial10

    redondeo_row = ws.max_row + 1
    ws.append([None, None, None, 'REDONDEO EL PESO BRUTO POR MRN', None, None, None, None, None])
    for cell in ws[redondeo_row]:
        cell.font = arial10

    for col, width in zip('ABCDEFGHI', [14, 12, 12, 35, 32, 12, 12, 8, 8]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = 'A5'


def add_t1_summary_sheet(wb: openpyxl.Workbook, t1_info: list[dict], sheet_name: str = 'T1_SUMMARY') -> None:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    arial10_bold = Font(name='Arial', size=10, bold=True)

    ws.append(['Source File', 'MRN', 'Gross Weight (kg)', 'Packages', 'Items', 'Deadline'])
    for cell in ws[1]:
        cell.font = arial10_bold
        cell.fill = _fill(_GREY_HEADER)

    for info in t1_info:
        ws.append([info['source_file'], info['mrn'], info['gross_kg'], info['packages'], info['items'], info['deadline']])

    for col, width in zip('ABCDEF', [30, 22, 18, 12, 10, 14]):
        ws.column_dimensions[col].width = width


# ============================================================================
# CLI
# ============================================================================


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='Kenitra unified packing-list processor')
    parser.add_argument('--xlsx', required=True, help='Input packing-list Excel (Sheet1 tab)')
    parser.add_argument('--t1', required=True, nargs='+', help='T1 transit PDF files')
    parser.add_argument('--doc', dest='doc', help='DOC PDF para validar el XLSX antes de procesar')
    parser.add_argument('--pdf', dest='doc', help=argparse.SUPPRESS)
    parser.add_argument('--output', default=None, help='Output Excel path')
    parser.add_argument('--verbose', action='store_true', help='Print row-by-row detail')
    parser.add_argument('--dpi', type=int, default=150, help='DPI para OCR del PDF DOC')
    return parser.parse_args()


def derive_output_path(xlsx_path: str, output: Optional[str]) -> Path:
    if output:
        return Path(output)
    src = Path(xlsx_path)
    return src.with_name(f'{src.stem}-PROCESSED.xlsx')


def main() -> int:
    args = parse_args()
    output_path = derive_output_path(args.xlsx, args.output)

    print('\n=== Camión Export Processor ===\n')

    print('📄 Extracting T1 data...')
    t1_info_list = []
    t1_map: dict[str, float] = {}
    for pdf_path in args.t1:
        info = extract_t1_info(pdf_path)
        t1_info_list.append(info)
        if info['mrn']:
            t1_map[info['mrn']] = info['gross_kg']
            print(f"  ✓ {info['source_file']}  →  MRN: {info['mrn']} | Gross: {info['gross_kg']:.0f} kg | Pkgs: {info['packages']} | Deadline: {info['deadline']}")
        else:
            print(f"  ⚠ Could not extract MRN from: {pdf_path}")

    print(f'\n📊 Reading packing list: {args.xlsx}')
    rows, summary = read_sheet1(args.xlsx)
    print(f'  ✓ {len(rows)} data rows')
    print(f'  ✓ Totals: gross={summary["total_peso_bruto"]} net={summary["total_peso_neto"]} PK={summary["total_pk"]} CL={summary["total_cl"]}')

    report = None
    if args.doc:
        print(f'\n🔎 Validating XLSX against PDF: {args.doc}')
        xlsx_path = Path(args.xlsx)
        pdf_path = Path(args.doc)

        entries = load_entries_from_xlsx(xlsx_path)
        page_texts = extract_pdf_text(pdf_path, dpi=args.dpi)
        report = build_report(entries, page_texts)

        print(f"  ✓ Validation summary: {json.dumps(report['summary'], ensure_ascii=False)}")

    print('\n⚙️ Applying transformation rules...')
    result = process_packing_list(rows, t1_map)

    if args.verbose:
        print('\n  Row detail:')
        print(f'  {"Sep":<4} {"Yellow":<7} {"Type":<8} {"MRN/Invoice":<35} {"PB":>8}')
        print('  ' + '-' * 68)
        for r in result:
            sep = '↑' if r['_separator_before'] else ''
            ylw = '●' if r['_src_yellow'] else ''
            pb = str(r['peso_bruto']) if r['peso_bruto'] is not None else '–'
            print(f'  {sep:<4} {ylw:<7} {r["_type"]:<8} {str(r["mrn_invoice"] or ""):<35} {pb:>8}')

    print(f'\n💾 Writing output workbook: {output_path}')
    wb = openpyxl.load_workbook(args.xlsx)
    if report is not None:
        add_validated_sheet(wb, report, validated_sheet_name='PDF_VALIDADO')
    add_processed_sheet(wb, result, summary, processed_sheet_name='PACKING_LIST_RESULT')
    add_t1_summary_sheet(wb, t1_info_list, sheet_name='T1_SUMMARY')
    wb.save(output_path)

    print('\n✅ Done!')
    print(f'  ✓ XLSX: {output_path}')
    print()
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
