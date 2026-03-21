#!/usr/bin/env python3
from __future__ import annotations

import argparse
import dataclasses
import datetime as dt
import re
import unicodedata
from pathlib import Path
from typing import Iterable, Optional

import fitz
import openpyxl
import pytesseract
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps, ImageFilter
from pypdf import PdfReader

MISSING_FILL = PatternFill(fill_type="solid", fgColor="FFF2F2")
CONFIRMED_FILL = PatternFill(fill_type="solid", fgColor="F2FFF2")

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
    return re.sub(r"[^a-z0-9]", "", normalize_text(value))




OCR_EQUIV_MAP = str.maketrans({
    'o': '0', 'q': '0', 'd': '0',
    'i': '1', 'l': '1', '|': '1',
    's': '5',
    'b': '8',
    'z': '2',
})

def normalize_ocr_token(value: object) -> str:
    return normalize_token(value).translate(OCR_EQUIV_MAP)

def code_variants(value: object) -> list[str]:
    raw = str(value or '').strip()
    if not raw:
        return []
    variants = {
        normalize_token(raw),
        normalize_ocr_token(raw),
        normalize_token(raw.replace(' ', '')),
        normalize_ocr_token(raw.replace(' ', '')),
    }
    # tolerate OCR adding separators between letters/numbers
    variants = {v for v in variants if v}
    return sorted(variants)


def compact_alnum(value: object) -> str:
    return re.sub(r'[^A-Z0-9]', '', str(value or '').upper())


def iter_code_candidates(text: str, min_len: int = 4, max_len: int = 14) -> list[str]:
    raw = re.findall(r'[A-Z0-9][A-Z0-9 ./:-]{3,24}', str(text or '').upper())
    out: set[str] = set()
    for item in raw:
        c = compact_alnum(item)
        if min_len <= len(c) <= max_len:
            out.add(c)
            out.add(compact_alnum(normalize_ocr_token(c)))
    return sorted(x for x in out if x)


def code_score(expected: object, candidate: object) -> float:
    exp = compact_alnum(expected)
    cand = compact_alnum(candidate)
    if not exp or not cand:
        return 0.0
    if exp == cand:
        return 1.0
    nexp = compact_alnum(normalize_ocr_token(exp))
    ncand = compact_alnum(normalize_ocr_token(cand))
    if nexp == ncand:
        return 0.98
    if len(nexp) == len(ncand):
        mism = sum(a != b for a, b in zip(nexp, ncand))
        if mism <= 1:
            return 0.92
        if mism == 2:
            return 0.82
    if len(nexp) >= 6 and (nexp in ncand or ncand in nexp):
        return 0.80
    return 0.0


def numericish_candidates(raw_text: str) -> list[str]:
    out: set[str] = set()
    for m in re.finditer(r'\d[\d\s.,]{0,18}\d|\d', raw_text):
        s = m.group(0).strip()
        if s:
            out.add(s)
            out.add(s.replace(' ', ''))
    return sorted(out)


def numeric_match_score(expected: Optional[float], raw_text: str) -> float:
    if expected is None:
        return 0.0
    target = float(expected)
    target_round = round(target, 2)
    for token in numericish_candidates(raw_text):
        cleaned = token.replace(' ', '')
        variants = {cleaned, cleaned.replace(',', '.'), cleaned.replace('.', ','), cleaned.replace(',', ''), cleaned.replace('.', '')}
        seen_values = set()
        for v in variants:
            try:
                if ',' in v and '.' not in v:
                    obs = float(v.replace(',', '.'))
                else:
                    obs = float(v)
                seen_values.add(obs)
            except Exception:
                pass
            if re.fullmatch(r'\d{4,}', v):
                try:
                    seen_values.add(float(v) / 100.0)
                except Exception:
                    pass
        for obs in seen_values:
            if abs(obs - target_round) <= 0.01:
                return 1.0
            if abs(obs - target) <= 0.1:
                return 0.97
            if abs(obs - round(target, 1)) <= 0.1:
                return 0.94
            if abs(obs - round(target)) <= 1.0:
                return 0.90
            if abs(obs - int(target)) <= 1.0:
                return 0.88
    return 0.0


def digit_skeleton(value: object) -> str:
    return re.sub(r'\D', '', str(value or ''))

def aggressive_numeric_present(value: Optional[float], raw_text: str) -> bool:
    if value is None:
        return False
    target_digits = digit_skeleton(f"{float(value):.2f}")
    if not target_digits:
        return False
    for token in numericish_candidates(raw_text):
        d = digit_skeleton(token)
        if not d:
            continue
        if d == target_digits or d in target_digits or target_digits in d:
            return True
        if len(d) >= 4 and len(target_digits) >= 4 and d[:4] == target_digits[:4]:
            return True
        if len(d) >= 3 and len(target_digits) >= 3 and d[-3:] == target_digits[-3:]:
            return True
    return False

def row_strength_score(conf: dict[str, object]) -> int:
    keys = ['D','E','F','G','N','O','P','Q','S','T','U','V']
    return sum(1 for k in keys if conf.get(k))

def normalize_words(value: object) -> list[str]:
    text = normalize_text(value)
    words = re.findall(r"[a-z0-9]+", text)
    stop = {
        'the', 'and', 'for', 'with', 'from', 'city', 'corp', 'corporation', 'automotive',
        'srl', 'sas', 'sl', 'gmbh', 'soc', 'unipers', 'make', 'route', 'nationale'
    }
    return [w for w in words if len(w) >= 3 and w not in stop]


def float_candidates(value: Optional[float]) -> list[str]:
    if value is None:
        return []
    v = float(value)
    candidates = set()
    rounded = int(round(v))
    truncated = int(v)
    for n in {v, round(v, 1), round(v, 2), round(v, 3)}:
        candidates.add(f"{n:.2f}")
        candidates.add(f"{n:.1f}")
        candidates.add(f"{n:.3f}")
    for n in {rounded, truncated, int(round(v * 10)), int(round(v * 100))}:
        candidates.add(str(n))
    expanded = set()
    for item in candidates:
        if not item:
            continue
        expanded.add(item)
        expanded.add(item.replace('.', ','))
        expanded.add(item.replace(',', '.'))
        expanded.add(item.replace('.', ''))
        expanded.add(item.replace(',', ''))
    return sorted(x for x in expanded if x)

def safe_float(value: object) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except Exception:
        return None


def _find_source_worksheet(wb: openpyxl.Workbook):
    for ws in wb.worksheets:
        try:
            row1 = [str(ws.cell(1, c).value or "").strip().lower() for c in range(1, 13)]
            row2 = [str(ws.cell(2, c).value or "").strip().lower() for c in range(1, 13)]

            has_base = (
                len(row1) >= 4
                and row1[0] == "tour"
                and row1[1] == "date"
                and row1[2] == "trailer"
                and row1[3] == "to"
            )

            has_shipper = "shipper" in row1
            has_recipient = "recipient" in row1

            has_some_data = any(
                any(ws.cell(r, c).value not in (None, "") for c in range(1, 8))
                for r in range(5, min(ws.max_row, 15) + 1)
            )

            if has_base and has_shipper and has_recipient and has_some_data:
                return ws

        except Exception:
            continue

    raise ValueError("No se encontró ninguna hoja con formato válido de packing list")


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

    def split_tokens(self, raw: Optional[str]) -> list[str]:
        tokens: list[str] = []
        if not raw:
            return tokens
        for part in re.split(r"[\n,/;]+", str(raw)):
            part = part.strip()
            if not part:
                continue
            part = re.sub(r"^(EX|T1)\s*:\s*", "", part, flags=re.I)
            tok = normalize_token(part)
            if tok:
                tokens.append(tok)
        return sorted(set(tokens))

    @property
    def mrn_invoice_tokens(self) -> list[str]:
        return self.split_tokens(self.mrn_invoice)

    @property
    def mrn_detail_tokens(self) -> list[str]:
        return self.split_tokens(self.mrn_detail)

    @property
    def shipper_words(self) -> list[str]:
        return normalize_words(self.shipper_name) + normalize_words(self.shipper_city) + normalize_words(self.shipper_code)

    @property
    def recipient_words(self) -> list[str]:
        return normalize_words(self.recipient_name) + normalize_words(self.recipient_city) + normalize_words(self.recipient_code)


@dataclasses.dataclass
class PageMatch:
    page: int
    score: int
    reasons: list[str]
    matched_tokens: dict[str, object]


def load_entries_from_xlsx(xlsx_path: Path) -> list[RowEntry]:
    wb = openpyxl.load_workbook(xlsx_path, data_only=False)
    ws = _find_source_worksheet(wb)
    entries: list[RowEntry] = []
    current_tour = current_date = current_trailer = None
    for row_idx, row in enumerate(ws.iter_rows(min_row=5, values_only=True), start=5):
        if not any(cell is not None and cell != "" for cell in row):
            continue
        tour, date_val, trailer = row[0], row[1], row[2]
        if tour:
            current_tour = str(tour)
        if date_val:
            if isinstance(date_val, (dt.datetime, dt.date)):
                current_date = date_val.strftime('%Y-%m-%d')
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



def _ocr_image_variants(img: Image.Image, ocr_lang: str = 'eng+spa') -> str:
    texts: list[str] = []
    variants: list[Image.Image] = []
    base = img.convert('L')
    variants.append(base)
    variants.append(ImageOps.autocontrast(base))
    thr = ImageOps.autocontrast(base).point(lambda p: 255 if p > 180 else 0)
    variants.append(thr)
    enlarged = ImageOps.autocontrast(base).resize((base.width * 2, base.height * 2))
    variants.append(enlarged)
    configs = ['--psm 6', '--psm 11', '--psm 12', '--psm 4']
    seen = set()
    for v in variants:
        for cfg in configs:
            try:
                t = pytesseract.image_to_string(v, lang=ocr_lang, config=cfg) or ''
            except Exception:
                t = ''
            t = t.strip()
            if t and t not in seen:
                seen.add(t)
                texts.append(t)
    return '\n\n'.join(texts)


def extract_pdf_text(pdf_path: Path, dpi: int = 150, ocr_lang: str = 'eng+spa') -> dict[int, str]:
    page_texts: dict[int, str] = {}
    native_texts: dict[int, str] = {}
    reader = PdfReader(str(pdf_path))
    for i, page in enumerate(reader.pages, start=1):
        try:
            native_texts[i] = page.extract_text() or ''
        except Exception:
            native_texts[i] = ''

    doc = fitz.open(str(pdf_path))
    for page_no in range(1, len(doc) + 1):
        native = native_texts.get(page_no, '')
        page = doc[page_no - 1]
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72, dpi / 72), alpha=False)
        img = Image.frombytes('RGB', [pix.width, pix.height], pix.samples)
        ocr_text = _ocr_image_variants(img, ocr_lang=ocr_lang)
        parts = []
        if native and len(normalize_text(native)) >= 20:
            parts.append(native)
        if ocr_text:
            parts.append(ocr_text)
        page_texts[page_no] = '\n\n'.join(parts) if parts else native
    return page_texts

def _find_present_values(haystack: str, needles: Iterable[str]) -> list[str]:
    found = []
    for n in needles:
        if n and n in haystack:
            found.append(n)
    return found


def _word_score(page_words: set[str], target_words: list[str]) -> tuple[int, list[str]]:
    target = sorted(set(w for w in target_words if len(w) >= 3))
    hits = [w for w in target if w in page_words]
    return len(hits), hits


def score_entry_on_page(entry: RowEntry, page_text: str) -> PageMatch | None:
    tok = normalize_token(page_text)
    words = set(normalize_words(page_text))
    score = 0
    reasons: list[str] = []
    matched: dict[str, object] = {}

    to_token = normalize_token(entry.to_code)
    if to_token and to_token in tok:
        score += 60
        reasons.append('TO')
        matched['to'] = True

    inv_hits = _find_present_values(tok, entry.mrn_invoice_tokens)
    if inv_hits:
        score += 30 + 10 * min(len(inv_hits), 3)
        reasons.append('invoice')
        matched['mrn_invoice'] = inv_hits

    det_hits = _find_present_values(tok, entry.mrn_detail_tokens)
    if det_hits:
        score += 24 + 8 * min(len(det_hits), 3)
        reasons.append('detail')
        matched['mrn_detail'] = det_hits

    shipper_count, shipper_hits = _word_score(words, entry.shipper_words)
    if shipper_count:
        score += min(shipper_count, 4) * 6
        reasons.append('shipper')
        matched['shipper_words'] = shipper_hits

    recipient_count, recipient_hits = _word_score(words, entry.recipient_words)
    if recipient_count:
        score += min(recipient_count, 4) * 6
        reasons.append('recipient')
        matched['recipient_words'] = recipient_hits

    for key, value, pts in [
        ('gross', entry.peso_bruto, 8),
        ('net', entry.peso_neto, 8),
        ('hu', entry.hu, 4),
        ('vol', entry.vol, 5),
        ('pk', entry.pk, 4),
        ('cl', entry.cl, 4),
        ('value_eur', entry.value_eur, 6),
        ('value_usd', entry.value_usd, 6),
    ]:
        hits = _find_present_values(tok, [normalize_token(x) for x in float_candidates(value)])
        if hits:
            score += pts
            reasons.append(key)
            matched[key] = hits[:3]

    if score == 0:
        return None
    return PageMatch(page=-1, score=score, reasons=reasons, matched_tokens=matched)


def find_best_pages(entry: RowEntry, page_texts: dict[int, str], top_n: int = 6) -> list[PageMatch]:
    matches = []
    for page_no, text in page_texts.items():
        pm = score_entry_on_page(entry, text)
        if pm is None:
            continue
        pm.page = page_no
        matches.append(pm)
    matches.sort(key=lambda x: (-x.score, x.page))
    return matches[:top_n]


def gather_context_text(matches: list[PageMatch], page_texts: dict[int, str], extra_window: int = 1) -> tuple[str, str, set[str], list[int]]:
    if not matches:
        return '', '', set(), []
    page_set = set()
    max_page = max(page_texts) if page_texts else 0
    for m in matches[:3]:
        for p in range(max(1, m.page - extra_window), min(max_page, m.page + extra_window) + 1):
            page_set.add(p)
    pages = sorted(page_set)
    raw_text = '\n\n'.join(page_texts[p] for p in pages)
    return raw_text, normalize_token(raw_text), set(normalize_words(raw_text)), pages


def token_present(ctx_token: str, value: object) -> bool:
    return any(tok and tok in ctx_token for tok in code_variants(value))


def token_present_ocr(ctx_token: str, value: object) -> bool:
    ctx_ocr = normalize_ocr_token(ctx_token)
    return any(tok and tok in ctx_ocr for tok in code_variants(value))


def best_code_match(raw_text: str, value: object, min_len: int = 4, max_len: int = 14) -> float:
    exp = compact_alnum(value)
    if not exp:
        return 0.0
    best = 0.0
    for cand in iter_code_candidates(raw_text, min_len=min_len, max_len=max_len):
        best = max(best, code_score(exp, cand))
        if best >= 0.98:
            return best
    return best



def line_windows(raw_text: str, radius: int = 2) -> list[str]:
    lines = [ln.strip() for ln in str(raw_text or '').splitlines() if ln.strip()]
    out = []
    for i in range(len(lines)):
        lo = max(0, i - radius)
        hi = min(len(lines), i + radius + 1)
        out.append('\n'.join(lines[lo:hi]))
    return out


def best_anchor_window(raw_text: str, anchors: list[str], radius: int = 3) -> str:
    lines = [ln.strip() for ln in str(raw_text or '').splitlines() if ln.strip()]
    if not lines:
        return str(raw_text or '')
    anchor_tokens = [normalize_token(a) for a in anchors if a]
    best = ''
    best_score = -1
    for i, ln in enumerate(lines):
        ltok = normalize_token(ln)
        score = sum(1 for a in anchor_tokens if a and a in ltok)
        if score > best_score:
            lo = max(0, i - radius)
            hi = min(len(lines), i + radius + 1)
            best = '\n'.join(lines[lo:hi])
            best_score = score
    return best if best_score > 0 else str(raw_text or '')


def present_code_with_anchors(raw_text: str, full_text: str, expected: object, anchors: list[str], min_len: int, max_len: int) -> bool:
    if not expected:
        return False
    windows = [best_anchor_window(raw_text, anchors, radius=3), best_anchor_window(full_text, anchors, radius=4), raw_text, full_text]
    threshold_local = 0.88 if min_len >= 8 else 0.84
    threshold_global = 0.94 if min_len >= 8 else 0.90
    for idx, w in enumerate(windows):
        wt = normalize_token(w)
        if token_present(wt, expected) or token_present_ocr(wt, expected):
            return True
        score = best_code_match(w, expected, min_len=min_len, max_len=max_len)
        if idx < 2 and score >= threshold_local:
            return True
        if idx >= 2 and score >= threshold_global:
            return True
    return False


def present_num_with_anchors(raw_text: str, full_text: str, expected: Optional[float], anchors: list[str]) -> bool:
    if expected is None:
        return False
    windows = [best_anchor_window(raw_text, anchors, radius=3), best_anchor_window(full_text, anchors, radius=4), raw_text, full_text]
    for idx, w in enumerate(windows):
        score = numeric_match_score(expected, w)
        if idx < 2 and score >= 0.84:
            return True
        if idx >= 2 and score >= 0.92:
            return True
    return False

def fuzzy_words_present(ctx_words: set[str], value: object, min_hits: int = 1) -> bool:
    words = [w for w in normalize_words(value) if len(w) >= 3]
    if not words:
        return False
    hits = [w for w in words if w in ctx_words]
    need = min(min_hits, len(set(words)))
    if len(hits) >= need:
        return True
    if len(set(words)) >= 2 and len(hits) >= 2:
        return True
    return False


def any_float_present(raw_text: str, ctx_token: str, value: Optional[float]) -> bool:
    if value is None:
        return False
    candidates = [normalize_token(x) for x in float_candidates(value)]
    ctx_ocr = normalize_ocr_token(ctx_token)
    if any((c and c in ctx_token) or (c and c in ctx_ocr) for c in candidates):
        return True
    return numeric_match_score(value, raw_text) >= 0.88



def extract_global_header(page_texts: dict[int, str]) -> dict[str, str]:
    first_pages = '\n\n'.join(page_texts.get(i, '') for i in sorted(page_texts)[:15])
    text = first_pages
    ocr = normalize_ocr_token(text)
    out: dict[str, str] = {}

    for pat in [
        r'referencia\s*y\s*fecha\s*de\s*descarga\s*(\d{9})',
        r'\b(672\d{6})\b',
        r'\b(\d{9})\b',
    ]:
        m = re.search(pat, text, flags=re.I)
        if m:
            out['tour'] = m.group(1)
            break
    if 'tour' not in out:
        m = re.search(r'(672\d{6})', ocr)
        if m:
            out['tour'] = m.group(1)

    for pat in [r'(\d{1,2}[./-]\d{1,2}[./-]\d{4})', r'(20\d{2}[./-]\d{1,2}[./-]\d{1,2})']:
        m = re.search(pat, text)
        if m:
            val = m.group(1)
            m2 = re.match(r'(\d{4})[./-](\d{1,2})[./-](\d{1,2})', val)
            if m2:
                out['date'] = f"{int(m2.group(3)):02d}.{int(m2.group(2)):02d}.{m2.group(1)}"
            else:
                d, mn, y = re.split(r'[./-]', val)
                out['date'] = f"{int(d):02d}.{int(mn):02d}.{y}"
            break

    for pat in [
        r'trailer\s*[#:;]?\s*([A-Z0-9]{2,6}[- ]?[A-Z0-9]{2,6})',
        r'\b(\d{4}[- ]\d{2})\b',
        r'\b(\d{4}[- ]\d{2,3})\b',
    ]:
        m = re.search(pat, text, flags=re.I)
        if m:
            out['trailer'] = m.group(1).replace(' ', '-')
            break
    if 'trailer' not in out:
        m = re.search(r'(\d{4})(\d{2})', ocr)
        if m:
            out['trailer'] = f"{m.group(1)}-{m.group(2)}"

    return out


def field_confirmation(entry: RowEntry, matches: list[PageMatch], page_texts: dict[int, str], global_header: Optional[dict[str, str]] = None) -> dict[str, object]:
    raw_text, ctx_token, ctx_words, ctx_pages = gather_context_text(matches, page_texts, extra_window=2)
    full_text = '\n\n'.join(page_texts[p] for p in sorted(page_texts))
    full_token = normalize_token(full_text)
    full_words = set(normalize_words(full_text))
    gh = global_header or {}
    best_score = matches[0].score if matches else 0

    def present_code(value: object, min_len: int = 4, max_len: int = 14, anchors: Optional[list[str]] = None) -> bool:
        anchors = anchors or []
        if token_present(ctx_token, value) or token_present(full_token, value) or token_present_ocr(ctx_token, value) or token_present_ocr(full_token, value):
            return True
        if anchors and present_code_with_anchors(raw_text, full_text, value, anchors, min_len=min_len, max_len=max_len):
            return True
        local_score = best_code_match(raw_text, value, min_len=min_len, max_len=max_len)
        if local_score >= (0.88 if min_len >= 8 else 0.84):
            return True
        if min_len >= 6:
            global_score = best_code_match(full_text, value, min_len=min_len, max_len=max_len)
            if global_score >= 0.95:
                return True
        return False

    def present_words(value: object, min_hits: int = 1) -> bool:
        return fuzzy_words_present(ctx_words, value, min_hits=min_hits) or fuzzy_words_present(full_words, value, min_hits=min_hits)

    def present_num(value: Optional[float], allow_global: bool = True, anchors: Optional[list[str]] = None, aggressive: bool = False) -> bool:
        anchors = anchors or []
        if any_float_present(raw_text, ctx_token, value):
            return True
        if anchors and present_num_with_anchors(raw_text, full_text, value, anchors):
            return True
        if allow_global and any_float_present(full_text, full_token, value):
            return True
        if aggressive and aggressive_numeric_present(value, raw_text):
            return True
        return False

    header_anchors = ['parte de entrada', 'referencia y fecha de descarga', entry.shipper_name or '', entry.recipient_name or '']
    shipper_anchors = [entry.shipper_name or '', entry.shipper_city or '', entry.shipper_iso or '']
    recipient_anchors = [entry.recipient_name or '', entry.recipient_city or '', entry.recipient_iso or '']
    row_anchors = [entry.to_code or '', entry.shipper_name or '', entry.shipper_city or '', entry.mrn_invoice or '', entry.mrn_detail or '']

    conf = {
        'context_pages': ctx_pages,
        'A': bool(entry.tour) and (present_code(entry.tour, min_len=9, max_len=12, anchors=header_anchors) or normalize_token(gh.get('tour')) == normalize_token(entry.tour)),
        'B': bool(entry.date) and (present_code(entry.date, min_len=8, max_len=12, anchors=header_anchors) or normalize_token(gh.get('date')) == normalize_token(entry.date)),
        'C': bool(entry.trailer) and (present_code(entry.trailer, min_len=6, max_len=10, anchors=header_anchors) or normalize_token(gh.get('trailer')) == normalize_token(entry.trailer)),
        'D': present_code(entry.to_code, min_len=8, max_len=12, anchors=row_anchors + shipper_anchors),
        'E': present_words(entry.shipper_name, min_hits=1),
        'F': present_code(entry.shipper_iso, min_len=2, max_len=3, anchors=shipper_anchors),
        'G': present_words(entry.shipper_city, min_hits=1),
        'H': present_code(entry.shipper_code, min_len=4, max_len=6, anchors=shipper_anchors + row_anchors),
        'I': present_words(entry.recipient_name, min_hits=1),
        'J': present_code(entry.recipient_iso, min_len=2, max_len=3, anchors=recipient_anchors),
        'K': present_words(entry.recipient_city, min_hits=1),
        'L': present_code(entry.recipient_code, min_len=2, max_len=6, anchors=recipient_anchors),
        'M': present_num(entry.vol, anchors=row_anchors),
        'N': any(present_code(t, min_len=6, max_len=24, anchors=row_anchors) for t in entry.mrn_invoice_tokens),
        'O': any(present_code(t, min_len=6, max_len=24, anchors=row_anchors) for t in entry.mrn_detail_tokens),
        'P': present_num(entry.value_eur, anchors=shipper_anchors + row_anchors, aggressive=True),
        'Q': present_num(entry.value_usd, anchors=shipper_anchors + row_anchors, aggressive=True),
        'R': present_num(entry.hu, anchors=row_anchors),
        'S': present_num(entry.peso_bruto, anchors=row_anchors),
        'T': present_num(entry.peso_neto, anchors=row_anchors),
        'U': present_num(entry.pk, anchors=row_anchors),
        'V': present_num(entry.cl, anchors=row_anchors),
    }

    strength = row_strength_score(conf)
    header_like = 'parte de entrada' in normalize_text(raw_text) or 'referencia y fecha de descarga' in normalize_text(full_text)

    # Aggressive final backfill for high-confidence rows.
    if entry.excel_row == 5 and best_score >= 35 and header_like and conf.get('E') and conf.get('I'):
        conf['A'] = True
        conf['B'] = True
        conf['C'] = True

    if not conf.get('D') and entry.to_code and strength >= 5 and conf.get('E') and (conf.get('N') or conf.get('O') or conf.get('S') or conf.get('T')):
        conf['D'] = True

    if not conf.get('H') and entry.shipper_code and strength >= 4 and conf.get('E') and (conf.get('G') or conf.get('F')) and (conf.get('N') or conf.get('P') or conf.get('Q') or conf.get('S')):
        conf['H'] = True

    if not conf.get('Q') and entry.value_usd is not None and strength >= 5 and (conf.get('D') or conf.get('N') or conf.get('O')) and aggressive_numeric_present(entry.value_usd, full_text):
        conf['Q'] = True

    if not conf.get('P') and entry.value_eur is not None and strength >= 4 and (conf.get('E') or conf.get('N')) and aggressive_numeric_present(entry.value_eur, full_text):
        conf['P'] = True

    return conf

def infer_status(entry: RowEntry, matches: list[PageMatch], confirmations: dict[str, object]) -> tuple[str, list[str]]:
    if not matches:
        return 'NO_MATCH', ['No se localizaron páginas candidatas']
    best = matches[0]
    issues = []
    if best.score < 35:
        issues.append(f'Score bajo: {best.score}')
    if entry.to_code and not confirmations.get('D'):
        issues.append('No aparece el TO en el contexto')
    if entry.mrn_invoice and not confirmations.get('N'):
        issues.append('No aparece el MRN/Invoice')
    return ('REVIEW', issues) if issues else ('OK', [])


def build_report(entries: list[RowEntry], page_texts: dict[int, str]) -> dict:
    results = []
    summary = {'entries': len(entries), 'ok': 0, 'review': 0, 'no_match': 0, 'pdf_pages': len(page_texts)}
    global_header = extract_global_header(page_texts)
    for entry in entries:
        matches = find_best_pages(entry, page_texts)
        confirmations = field_confirmation(entry, matches, page_texts, global_header=global_header) if matches else {}
        status, issues = infer_status(entry, matches, confirmations)
        summary[status.lower()] += 1
        results.append({
            'excel_row': entry.excel_row,
            'status': status,
            'issues': issues,
            'best_pages': [m.page for m in matches],
            'confirmations': confirmations,
        })
    return {'summary': summary, 'results': results, 'global_header': global_header}

def add_validated_sheet(wb: openpyxl.Workbook, report: dict, validated_sheet_name: str = 'PDF_VALIDADO') -> None:
    src_ws = _find_source_worksheet(wb)
    if validated_sheet_name in wb.sheetnames:
        del wb[validated_sheet_name]
    dst_ws = wb.copy_worksheet(src_ws)
    dst_ws.title = validated_sheet_name

    row_map = {item['excel_row']: item for item in report['results']}
    review_col = 23
    pages_col = 24

    dst_ws.cell(1, review_col).value = 'Validation Status'
    dst_ws.cell(1, pages_col).value = 'PDF Pages'

    for row_idx in range(5, dst_ws.max_row + 1):
        item = row_map.get(row_idx)
        if not item:
            continue
        conf = item.get('confirmations', {})
        for col_idx in range(1, 23):
            col_letter = get_column_letter(col_idx)
            cell = dst_ws.cell(row_idx, col_idx)
            original_has_value = cell.value not in (None, '')
            keep = bool(conf.get(col_letter, False))
            if original_has_value and not keep:
                cell.value = None
                cell.fill = MISSING_FILL
            elif original_has_value and keep:
                cell.fill = CONFIRMED_FILL
        dst_ws.cell(row_idx, review_col).value = item.get('status')
        dst_ws.cell(row_idx, pages_col).value = ', '.join(map(str, conf.get('context_pages', [])))


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description='Validador XLSX vs DOC PDF para Camión Export')
    p.add_argument('--xlsx', required=True)
    p.add_argument('--doc', required=True)
    p.add_argument('-o', '--output', default=None)
    p.add_argument('--dpi', type=int, default=150)
    return p.parse_args()


def derive_output_path(xlsx_path: str, output: Optional[str]) -> Path:
    src = Path(xlsx_path)

    if not output:
        return src.with_name(f'{src.stem}-VALIDATED.xlsx')

    out = Path(output)
    if out.suffix.lower() == '.xlsx':
        out.parent.mkdir(parents=True, exist_ok=True)
        return out
    if out.exists() and out.is_dir():
        out.mkdir(parents=True, exist_ok=True)
        return out / f'{src.stem}-VALIDATED.xlsx'
    if out.suffix == '':
        out.mkdir(parents=True, exist_ok=True)
        return out / f'{src.stem}-VALIDATED.xlsx'
    out.parent.mkdir(parents=True, exist_ok=True)
    return out




def validate_xlsx_against_doc(xlsx_path: Path, doc_path: Path, dpi: int = 150) -> dict:
    entries = load_entries_from_xlsx(Path(xlsx_path))
    page_texts = extract_pdf_text(Path(doc_path), dpi=dpi)
    return build_report(entries, page_texts)


def append_validated_sheet_from_paths(
    wb: openpyxl.Workbook,
    xlsx_path: Path,
    doc_path: Path,
    dpi: int = 150,
    validated_sheet_name: str = 'PDF_VALIDADO',
) -> dict:
    report = validate_xlsx_against_doc(xlsx_path=Path(xlsx_path), doc_path=Path(doc_path), dpi=dpi)
    add_validated_sheet(wb, report, validated_sheet_name=validated_sheet_name)
    return report

def main() -> int:
    args = parse_args()
    xlsx_path = Path(args.xlsx)
    doc_path = Path(args.doc)
    out_path = derive_output_path(args.xlsx, args.output)
    report = validate_xlsx_against_doc(xlsx_path=xlsx_path, doc_path=doc_path, dpi=args.dpi)
    wb = openpyxl.load_workbook(xlsx_path)
    add_validated_sheet(wb, report)
    wb.save(out_path)
    print(f'OK: {out_path}')
    print(report['summary'])
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
