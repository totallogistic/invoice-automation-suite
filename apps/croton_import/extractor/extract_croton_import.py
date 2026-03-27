#!/usr/bin/env python3
"""
ifortex_processor.py  –  IFORTEX Factura PDF → Excel
=====================================================
Extrae datos de la factura IFORTEX (PDF imagen) y rellena
la plantilla IFORTEX_TEMPLATE.xlsx con fórmulas activas.
Solo falta rellenar la columna BRUTO (amarillo) a mano.

Uso:
    python ifortex_processor.py --pdf FACTURA_N_20.pdf
    python ifortex_processor.py --pdf FACTURA_N_20.pdf --template IFORTEX_TEMPLATE.xlsx --output resultado.xlsx
"""
from __future__ import annotations

SCRIPT_VERSION = "2026-03-26.v1"

import argparse, io, re, sys
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path

import pytesseract
from PIL import Image
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Partidas arancelarias ────────────────────────────────────────────────────
PARTIDAS = {
    'JERSEY POLAR CBRO':   6110309100,
    'JERSEY POLAR SRA':    6110309900,
    'SOFTSHELL SRA':       6201939010,
    'BLUSA SRA':           6206100000,
    'SUDADERA SR':         6110309100,
    'POLO SR':             6105201000,
    'CAMISETA SR':         6109902000,
    'CAMISA SR':           6105209000,
    'PANTALON SR':         6203429000,
    'PANTALON SRA':        6204629000,
    'PANTALON CBRO':       6203429000,
    'CHAQUETA SR':         6201939010,
    'CASACA SR':           6211339000,
    'BATA SR':             6211339000,
    'DELANTAL':            6211339000,
    'COFIA':               6505009090,
    'CHALECO POLAR CBRO':  6110309100,
    'CHALECO  POLAR CBRO': 6110309100,
    'AMERICANA SRA':       6204310000,
}

KNOWN_DESC = list(PARTIDAS.keys())

# ── Styles ───────────────────────────────────────────────────────────────────
def _fill(rgb):
    return PatternFill('solid', start_color=rgb, end_color=rgb)
def _fn(bold=False, size=10, italic=False, color='FF000000'):
    return Font(bold=bold, size=size, italic=italic, name='Calibri', color=color)
def _border():
    s = Side(style='thin', color='FF000000')
    return Border(left=s, right=s, top=s, bottom=s)
def _right():  return Alignment(horizontal='right',  vertical='center')
def _center(): return Alignment(horizontal='center', vertical='center')
def _wrap():   return Alignment(horizontal='left',   vertical='center', wrap_text=True)

BRUTO_FILL   = _fill('FFFFF2CC')
INPUT_FILL   = _fill('FFE8F4E8')
FORMULA_FILL = _fill('FFEBF3FB')
TOTAL_FILL   = _fill('FFD9E1F2')
RESUMEN_FILL = _fill('FFFFE0CC')
HOJA6_FILL   = _fill('FFE2EFDA')
HEADER_FILL  = _fill('FF4472C4')
HEADER_FONT  = Font(bold=True, size=10, color='FFFFFFFF', name='Calibri')

# ── OCR ──────────────────────────────────────────────────────────────────────
def _ocr_pages(pdf_path: str) -> list[str]:
    reader = PdfReader(pdf_path)
    results = []
    for page in reader.pages:
        tiffs = [(img, len(img.data)) for img in page.images
                 if '.tiff' in img.name.lower()]
        if not tiffs:
            results.append('')
            continue
        main = max(tiffs, key=lambda x: x[1])[0]
        img  = Image.open(io.BytesIO(main.data))
        text = pytesseract.image_to_string(img, config='--psm 6')
        results.append(text)
    return results

# ── Parsers ───────────────────────────────────────────────────────────────────
def _eu(s: str) -> float:
    s = s.strip()
    if ',' in s and '.' in s:
        return float(s.replace('.', '').replace(',', '.'))
    if ',' in s:
        return float(s.replace(',', '.'))
    return float(s)

_num_re  = re.compile(r'\d+(?:[.,]\d+)*')
_of_re   = re.compile(r'^([A-Z]{0,2}-?\d{2}/\d{4,6})')
_ref_re  = re.compile(r'\b[A-Z]{1,5}\d+/\S+\b')
_comp_re = re.compile(r'\d{1,3}%')

def _best_match(ocr_text: str) -> str:
    t = re.sub(r'\s+', ' ', ocr_text.upper().strip())
    if not t:
        return ''
    best, best_score = None, 0
    for desc in KNOWN_DESC:
        d = desc.upper()
        ocr_words  = set(t.split())
        desc_words = set(d.split())
        word_score = len(ocr_words & desc_words)
        substr_score = sum(
            1 for w in ocr_words for dw in desc_words
            if len(w) >= 3 and (dw.endswith(w) or dw[1:] == w or dw == w)
        )
        ratio = SequenceMatcher(None, t, d).ratio()
        score = word_score * 3 + substr_score * 2 + ratio
        if score > best_score:
            best, best_score = desc, score
    return best or t

def _parse_composition(line: str) -> str:
    """Extract COMPOSITION text (contains % signs)."""
    m = _comp_re.search(line)
    if not m:
        return ''
    # From first % back to start of that token, forward until numbers begin
    start = max(0, line.rfind(' ', 0, m.start()) + 1)
    # Find end: where plain numbers start after composition
    rest = line[start:]
    nums = list(_num_re.finditer(rest))
    # Composition ends where we have 4 consecutive number tokens = PU, TOTAL, PROM, CONS
    comp_end = len(rest)
    for i in range(len(nums)-3):
        comp_end = nums[i].start()
        break
    return re.sub(r'\s+', ' ', rest[:comp_end]).strip()

def parse_invoice(text: str) -> tuple[list[dict], dict]:
    rows = []
    orden = 0
    for line in text.splitlines():
        line = line.strip()
        m_of = _of_re.match(line)
        if not m_of:
            continue
        of_raw = m_of.group(1)
        of = 'MP-' + re.sub(r'^[A-Z]*-?', '', of_raw)

        # Remove OF and REF
        rest = line[m_of.end():]
        rest = _ref_re.sub('', rest).strip().lstrip('=,—- ')

        # Description: before first XX%
        m_c = _comp_re.search(rest)
        desc_raw = rest[:m_c.start()].strip() if m_c else ''
        desc = _best_match(desc_raw)

        # Composition: from XX% until numbers start
        comp = _parse_composition(rest)

        # Last 4 numbers = PU, TOTAL, PROM, CONS
        nums = [_eu(m.group()) for m in _num_re.finditer(rest)]
        if len(nums) < 4:
            continue
        try:
            pu, total, prom, cons = nums[-4], nums[-3], nums[-2], nums[-1]
            qte = round(total / pu) if pu > 0 else 0
        except Exception:
            continue

        if qte <= 0 or pu <= 0:
            continue

        orden += 1
        rows.append({
            'of': of, 'orden': orden,
            'descripcion': desc, 'composicion': comp,
            'un': qte, 'pu': pu, 'total': total,
            'promedio': prom, 'comsumido': cons,
        })

    # Totals from invoice footer
    totals = {}
    for label, key in [
        (r'N[°º]\s*DE\s*COLIS',      'n_colis'),
        (r'TOTAL\s*POIDS\s*BRUT',    'poids_brut'),
        (r'TOTAL\s*POIDS\s*NET',     'poids_net'),
        (r'TOTAL\s*FACTURE',         'total_facture'),
    ]:
        m = re.search(label + r'\s+([\d.,]+)', text, re.IGNORECASE)
        if m:
            totals[key] = _eu(m.group(1))

    return rows, totals

def parse_invoice_number(text: str) -> str:
    m = re.search(r'F[-\s]?(\d{2}/\d{4,6}|\d{5,8})', text)
    return m.group(0).replace(' ', '') if m else 'F-UNKNOWN'

def parse_packing_list(pages_text: list[str]) -> dict[str, int]:
    bultos: dict[str, int] = defaultdict(int)
    of_re = re.compile(r'(MP-\d{2}/\d{4,6})\s+\d+')
    for text in pages_text:
        for m in of_re.finditer(text):
            bultos[m.group(1)] += 1
    return dict(bultos)

# ── Template filler ───────────────────────────────────────────────────────────
def fill_template(
    template_path: str,
    rows: list[dict],
    bultos: dict[str, int],
    totals: dict,
    inv_number: str,
    output_path: str,
) -> None:
    wb  = openpyxl.load_workbook(template_path)
    fn  = _fn()
    fnb = _fn(bold=True)

    # ── Feuil1 ──────────────────────────────────────────────────────────────
    ws1 = wb['Feuil1']

    for i, row in enumerate(rows, 2):
        bul = bultos.get(row['of'])

        data = {
            'A': row['of'],    'B': row['orden'],
            'C': row['descripcion'], 'D': row['composicion'],
            'E': row['un'],    'F': row['pu'],
            'H': row['promedio'], 'L': bul,
        }
        for col, val in data.items():
            c = ws1[f'{col}{i}']
            c.value  = val
            c.border = _border()
            c.font   = fn
            c.fill   = INPUT_FILL
            c.alignment = _right() if col in 'BEFHL' else _wrap()

        # BRUTO cell (yellow, empty, user fills)
        c = ws1[f'J{i}']
        c.fill   = BRUTO_FILL
        c.border = _border()
        c.font   = fnb
        c.alignment = _right()
        c.number_format = '#,##0.00'

        # Formula cells (already set in template, just style them)
        for col in ('G', 'I', 'K'):
            c = ws1[f'{col}{i}']
            c.fill      = FORMULA_FILL
            c.border    = _border()
            c.font      = fn
            c.alignment = _right()
            c.number_format = '#,##0.00'

        ws1[f'E{i}'].number_format = '#,##0'
        ws1[f'F{i}'].number_format = '#,##0.00'
        ws1[f'H{i}'].number_format = '#,##0.00'

    # Info note
    note_row = len(rows) + 3
    ws1.cell(note_row, 1).value = (
        f'Factura: {inv_number}  |  '
        f'COLIS PDF: {int(totals.get("n_colis",0))}  |  '
        f'POIDS BRUT PDF: {int(totals.get("poids_brut",0))}  |  '
        f'POIDS NET PDF: {int(totals.get("poids_net",0))}'
    )
    ws1.cell(note_row, 1).font = _fn(italic=True, size=9, color='FF666666')

    # ── Resumen ──────────────────────────────────────────────────────────────
    ws2 = wb['Resumen']
    ws2['A2'].value = None

    # Group by description (keep first orden, sum UN)
    groups: dict[str, dict] = {}
    for row in rows:
        desc = row['descripcion']
        if desc not in groups:
            groups[desc] = {'orden': row['orden'], 'comp': row['composicion'], 'un': 0}
        groups[desc]['un'] += row['un']

    n = len(rows)
    f1_desc   = f"Feuil1!C2:C{n+1}"
    f1_bruto  = f"Feuil1!J2:J{n+1}"
    f1_neto   = f"Feuil1!K2:K{n+1}"
    f1_valor  = f"Feuil1!G2:G{n+1}"
    f1_bultos = f"Feuil1!L2:L{n+1}"

    for i, (desc, g) in enumerate(groups.items(), 2):
        ws2[f'A{i}'].value = desc
        ws2[f'B{i}'].value = g['comp']
        ws2[f'C{i}'].value = g['orden']
        ws2[f'D{i}'].value = f'=SUMIF({f1_desc},A{i},{f1_bultos})'
        ws2[f'E{i}'].value = f'=ROUND(SUMIF({f1_desc},A{i},{f1_bruto}),0)'
        ws2[f'F{i}'].value = f'=SUMIF({f1_desc},A{i},{f1_neto})'
        ws2[f'G{i}'].value = f'=SUMIF({f1_desc},A{i},{f1_valor})'
        ws2[f'H{i}'].value = g['un']
        for col in 'ABCDEFGH':
            c = ws2[f'{col}{i}']
            c.fill   = RESUMEN_FILL
            c.border = _border()
            c.font   = fn
            c.alignment = _right() if col in 'CDEFGH' else _wrap()

    # Totals
    tr = len(groups) + 2
    for col, fml in [
        ('D', f'=SUM(D2:D{tr-1})'), ('E', f'=SUM(E2:E{tr-1})'),
        ('F', f'=SUM(F2:F{tr-1})'), ('G', f'=SUM(G2:G{tr-1})'),
        ('H', f'=SUM(H2:H{tr-1})'),
    ]:
        c = ws2[f'{col}{tr}']
        c.value = fml; c.fill = TOTAL_FILL
        c.border = _border(); c.font = fnb; c.alignment = _right()

    # ── Hoja6 ────────────────────────────────────────────────────────────────
    ws6 = wb['Hoja6']
    ws6['A2'].value = None

    res_range = f"Resumen!A2:A{len(groups)+1}"
    res_D = f"Resumen!D2:D{len(groups)+1}"
    res_E = f"Resumen!E2:E{len(groups)+1}"
    res_F = f"Resumen!F2:F{len(groups)+1}"
    res_G = f"Resumen!G2:G{len(groups)+1}"
    res_H = f"Resumen!H2:H{len(groups)+1}"

    for i, (desc, g) in enumerate(groups.items(), 2):
        partida = PARTIDAS.get(desc, '')
        ws6[f'A{i}'].value = partida or '⚠ sin partida'
        ws6[f'B{i}'].value = desc
        ws6[f'C{i}'].value = g['comp']
        ws6[f'D{i}'].value = g['orden']
        ws6[f'E{i}'].value = f'=SUMIF({res_range},B{i},{res_D})'
        ws6[f'F{i}'].value = f'=SUMIF({res_range},B{i},{res_E})'
        ws6[f'G{i}'].value = f'=SUMIF({res_range},B{i},{res_F})'
        ws6[f'H{i}'].value = f'=SUMIF({res_range},B{i},{res_G})'
        ws6[f'I{i}'].value = f'=SUMIF({res_range},B{i},{res_H})'
        for col in 'ABCDEFGHI':
            c = ws6[f'{col}{i}']
            c.fill   = HOJA6_FILL
            c.border = _border()
            c.font   = fn if partida else _fn(color='FFCC0000')
            c.alignment = _right() if col in 'ADEFGHI' else _wrap()

    tr6 = len(groups) + 2
    for col, fml in [
        ('E', f'=SUM(E2:E{tr6-1})'), ('F', f'=SUM(F2:F{tr6-1})'),
        ('G', f'=SUM(G2:G{tr6-1})'), ('H', f'=SUM(H2:H{tr6-1})'),
    ]:
        c = ws6[f'{col}{tr6}']
        c.value = fml; c.fill = TOTAL_FILL
        c.border = _border(); c.font = fnb; c.alignment = _right()

    wb.save(output_path)
    print(f'  ✓ Excel guardado: {output_path}')

# ── CLI ───────────────────────────────────────────────────────────────────────

def _derive_output_path(pdf_path: Path, output: str | None) -> Path:
    """Accept either a .xlsx path or a directory — same as camion convention."""
    if not output:
        return pdf_path.with_suffix('.xlsx')
    out = Path(output)
    if out.suffix.lower() == '.xlsx':
        out.parent.mkdir(parents=True, exist_ok=True)
        return out
    # Directory (or path without extension) → put file inside
    out.mkdir(parents=True, exist_ok=True)
    return out / f'{pdf_path.stem}_resultado.xlsx'

def main() -> int:
    parser = argparse.ArgumentParser(description='IFORTEX Factura PDF → Excel')
    parser.add_argument('--pdf',      required=True)
    parser.add_argument('--template', default=None)
    parser.add_argument('--output',   default=None)
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    tpl_path = Path(args.template) if args.template else Path(__file__).parent / 'IFORTEX_TEMPLATE.xlsx'
    out_path = _derive_output_path(pdf_path, args.output)

    if not tpl_path.exists():
        print(f'❌ Template no encontrado: {tpl_path}')
        return 1

    print(f'\n=== IFORTEX Processor {SCRIPT_VERSION} ===\n')
    print(f'📄 PDF:      {pdf_path}')
    print(f'📋 Template: {tpl_path}')
    print(f'💾 Output:   {out_path}')

    print('\n🔍 OCR...')
    pages = _ocr_pages(str(pdf_path))
    print(f'  ✓ {len(pages)} páginas')

    print('\n📊 Parseando factura (página 1)...')
    rows, totals = parse_invoice(pages[0])
    inv_number = parse_invoice_number(pages[0])
    print(f'  ✓ {len(rows)} artículos  |  Factura: {inv_number}')
    print(f'  ✓ COLIS: {int(totals.get("n_colis",0))}  BRUT: {int(totals.get("poids_brut",0))}  NET: {int(totals.get("poids_net",0))}')
    for r in rows:
        print(f'    {r["of"]:<15} UN={r["un"]:>5}  PU={r["pu"]:>8.2f}  {r["descripcion"]}')

    if not rows:
        print('❌ No se extrajeron artículos. Revisar OCR.')
        return 1

    print('\n📦 Contando bultos...')
    bultos = parse_packing_list(pages[1:])
    for of, b in sorted(bultos.items()):
        print(f'  {of}: {b}')

    print('\n📝 Rellenando template...')
    fill_template(str(tpl_path), rows, bultos, totals, inv_number, str(out_path))

    print(f'\n✅ Listo!')
    print(f'   → Rellena columna BRUTO (🟡 amarillo) en Feuil1')
    print(f'   → Resumen y Hoja6 se actualizan solos\n')
    return 0

if __name__ == '__main__':
    raise SystemExit(main())