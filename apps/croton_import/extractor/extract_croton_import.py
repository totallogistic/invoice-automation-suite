#!/usr/bin/env python3
"""
extract_croton_import.py  –  IFORTEX Factura PDF + XLSX → Excel
================================================================
Toma el XLSX que llega de producción (Feuil1 con todos los datos)
y el PDF de la factura IFORTEX (imagen OCR).

- Feuil1        → copia del input con ORDEN secuencial y COMPOSICION simplificada
- PDF_Extraído  → tabla raw de lo que el OCR saca del PDF
- Validación    → diff campo a campo Feuil1 vs PDF (verde OK, rojo error,
                  naranja composición sin mapear en composicion_map.csv)
- Hoja6         → agrupado por (descripción canónica + composición), HS code correcto

Uso:
    python3 extract_croton_import.py \\
        --pdf  FACTURA_N_21.pdf \\
        --xlsx ARCHIVO_EXCEL_FRA_21.xlsx \\
        --output /data/out/
"""
from __future__ import annotations

SCRIPT_VERSION = "2026-03-27.v2"

SCRIPT_CHANGELOG = """
## 2026-03-27.v2
- HS codes corregidos y diferenciados por composición (sint. vs algodón)
- Tabla de aliases de descripción: agrupa CHAQUETA+CASACA, SOFTSHELL+AMERICANA,
  DELANTAL+BATA, PANTALON CBRO→SR, etc.
- Hoja6 agrupa por (descripción canónica, composición simplificada)
- composicion_map.csv ampliado con todas las composiciones conocidas

## 2026-03-27.v1
- Eliminada dependencia de template xlsx
- Entrada: PDF (OCR/validación) + XLSX (datos reales de producción)
- Composición simplificada vía composicion_map.csv mantenible
- 4 hojas de salida: Feuil1, PDF_Extraído, Validación, Hoja6
- Validación campo a campo con colores
"""

import argparse, csv, io, re, sys
from collections import defaultdict
from difflib import SequenceMatcher
from pathlib import Path

try:
    import pytesseract
    from PIL import Image
    from pypdf import PdfReader
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
except ImportError as e:
    print(f"❌ Dependencia no instalada: {e}", flush=True)
    print("   pip install pytesseract pypdf openpyxl pillow", flush=True)
    sys.exit(1)


# ── Partidas arancelarias (HS codes) ─────────────────────────────────────────
# Clave: (descripcion_canónica, composicion_simplificada) o solo str
# La función get_partida() gestiona la búsqueda con fallback.

PARTIDAS: dict = {
    # Diferenciadas por composición
    ('CAMISA SR',        'F SINTETICAS'): 6205300000,
    ('CAMISA SR',        'ALGODON'):      6205200090,
    ('PANTALON SR',      'F SINTETICAS'): 6203431100,
    ('PANTALON SR',      'ALGODON'):      6203421100,
    ('PANTALON SRA',     'ALGODON'):      6204621100,
    ('BATAS',            'ALGODON'):      6211421000,
    ('DELANTAL Y BATAS', 'F SINTETICAS'): 6211431000,
    # Sin distinción de composición
    'SOFTSHELL Y AMERICANAS SRA': 6204331000,
    'BLUSA SRA':                  6206400000,
    'POLO SR':                    6105201000,
    'CHAQUETA Y CASACA SR':       6203331000,
    'COFIA':                      6505009090,
}

def get_partida(canonical_desc: str, comp_sim: str) -> int | None:
    """Busca HS code por (desc, comp) con fallback a solo desc."""
    return (
        PARTIDAS.get((canonical_desc, comp_sim))
        or PARTIDAS.get(canonical_desc)
    )


# ── Aliases de descripción ────────────────────────────────────────────────────
# Fusionan artículos distintos en la misma línea de Hoja6.

DESC_ALIASES_SIMPLE: dict[str, str] = {
    'SOFTSHELL SRA':  'SOFTSHELL Y AMERICANAS SRA',
    'AMERICANA SRA':  'SOFTSHELL Y AMERICANAS SRA',
    'CHAQUETA SR':    'CHAQUETA Y CASACA SR',
    'CASACA SR':      'CHAQUETA Y CASACA SR',
    'PANTALON CBRO':  'PANTALON SR',
    'PANTALON':       'PANTALON SR',    # typo sin SR
    'DELANTAL':       'DELANTAL Y BATAS',
}

# Aliases dependientes de composición: {desc: {comp_sim: canonical}}
DESC_ALIASES_BY_COMP: dict[str, dict[str, str]] = {
    'BATA SR': {
        'F SINTETICAS': 'DELANTAL Y BATAS',
        'ALGODON':      'BATAS',
    },
}

def canonical_desc(raw_desc: str, comp_sim: str) -> str:
    d = raw_desc.strip()
    if d in DESC_ALIASES_BY_COMP:
        return DESC_ALIASES_BY_COMP[d].get(comp_sim, d)
    return DESC_ALIASES_SIMPLE.get(d, d)


# ── Composicion map por defecto ───────────────────────────────────────────────
DEFAULT_COMP_MAP: dict[str, str] = {
    # F SINTETICAS — mayoría sintética
    '100% POLIESTER':          'F SINTETICAS',
    '93% POL 7% ELASTAN':      'F SINTETICAS',
    '67% POL 33% ALG':         'F SINTETICAS',
    '65% POL 35% ALG':         'F SINTETICAS',
    '65% POL 35% VISC':        'F SINTETICAS',
    '70% POL 30% ALG':         'F SINTETICAS',
    '52% POL 48% ALG':         'F SINTETICAS',
    '50% POL 50% ALG':         'F SINTETICAS',
    # ALGODON — mayoría algodón
    '100% ALGODON':            'ALGODON',
    '100% ALGODÓN':            'ALGODON',
    '70% ALG 30% POL':         'ALGODON',
    '59%ALG 39%POL 2%ELAST':   'ALGODON',
    '60% ALG 40% POL':         'ALGODON',
}


# ── Styles ────────────────────────────────────────────────────────────────────
def _fill(rgb):   return PatternFill('solid', start_color=rgb, end_color=rgb)
def _fn(bold=False, size=10, italic=False, color='FF000000'):
    return Font(bold=bold, size=size, italic=italic, name='Calibri', color=color)
def _border():
    s = Side(style='thin', color='FF000000')
    return Border(left=s, right=s, top=s, bottom=s)
def _right():  return Alignment(horizontal='right',  vertical='center')
def _center(): return Alignment(horizontal='center', vertical='center')
def _wrap():   return Alignment(horizontal='left',   vertical='center', wrap_text=True)

FILL_INPUT  = _fill('FFE8F4E8')
FILL_OK     = _fill('FFC6EFCE')
FILL_ERROR  = _fill('FFFFC7CE')
FILL_WARN   = _fill('FFFFEB9C')
FILL_TOTAL  = _fill('FFD9E1F2')
FILL_HOJA6  = _fill('FFE2EFDA')
FILL_PDF    = _fill('FFEBF3FB')
FILL_HEADER = _fill('FF4472C4')
FONT_HEADER = Font(bold=True, size=10, color='FFFFFFFF', name='Calibri')
FONT_WARN   = _fn(bold=True, color='FFCC0000')


def _header_row(ws, cols, row=1):
    for c, label in enumerate(cols, 1):
        cell = ws.cell(row, c, label)
        cell.fill = FILL_HEADER; cell.font = FONT_HEADER
        cell.border = _border(); cell.alignment = _center()


# ── Composición map ───────────────────────────────────────────────────────────
def load_composicion_map(script_dir: Path) -> dict[str, str]:
    csv_path = script_dir / 'composicion_map.csv'

    if not csv_path.exists():
        lines = ['raw,simplificado\n']
        lines += [f'{raw},{sim}\n' for raw, sim in sorted(DEFAULT_COMP_MAP.items())]
        csv_path.write_text(''.join(lines), encoding='utf-8')
        print(f'  ✓ composicion_map.csv creado con {len(DEFAULT_COMP_MAP)} entradas',
              flush=True)

    # Partir de defaults y sobreescribir con CSV
    mapping = {k.upper(): v for k, v in DEFAULT_COMP_MAP.items()}
    with open(csv_path, encoding='utf-8') as f:
        for row in csv.DictReader(
            line for line in f if not line.strip().startswith('#')
        ):
            raw = (row.get('raw') or '').strip().upper()
            sim = (row.get('simplificado') or '').strip()
            if raw and sim:
                mapping[raw] = sim

    print(f'  ✓ composicion_map: {len(mapping)} entradas', flush=True)
    return mapping


def simplify_comp(raw: str, mapping: dict[str, str]) -> tuple[str, bool]:
    key = (raw or '').strip().upper()
    if key in mapping:
        return mapping[key], True
    return raw or '', False


# ── Leer XLSX de entrada ──────────────────────────────────────────────────────
def read_input_xlsx(xlsx_path: Path) -> list[dict]:
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = wb['Feuil1']
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        of = row[0]
        if not of or not isinstance(of, str) or not of.strip():
            continue
        rows.append({
            'of':         str(of).strip(),
            'ref':        row[1],
            'descripcion': str(row[2]).strip() if row[2] else '',
            'comp_raw':   str(row[3]).strip() if row[3] else '',
            'un':         row[4],
            'pu':         row[5],
            'valor':      row[6],
            'cons_medio': row[7],
            'cons_total': row[8],
            'bruto':      row[9],
            'neto':       row[10],
            'bultos':     row[11],
            'ubicacion':  row[12],
        })
    return rows


# ── OCR ───────────────────────────────────────────────────────────────────────
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
        results.append(pytesseract.image_to_string(img, config='--psm 6'))
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

KNOWN_DESC = list({
    'SOFTSHELL SRA', 'AMERICANA SRA', 'BLUSA SRA', 'CAMISA SR',
    'PANTALON SR', 'PANTALON SRA', 'PANTALON CBRO', 'POLO SR',
    'CHAQUETA SR', 'CASACA SR', 'BATA SR', 'DELANTAL', 'COFIA',
    'JERSEY POLAR CBRO', 'JERSEY POLAR SRA', 'SUDADERA SR',
    'CAMISETA SR', 'CHALECO POLAR CBRO', 'CHALECO  POLAR CBRO',
})

def _best_match(ocr_text: str) -> str:
    t = re.sub(r'\s+', ' ', ocr_text.upper().strip())
    if not t: return ''
    best, best_score = None, 0
    for desc in KNOWN_DESC:
        d = desc.upper()
        ow = set(t.split()); dw = set(d.split())
        score = (len(ow & dw) * 3
                 + sum(1 for w in ow for x in dw
                       if len(w) >= 3 and (x.endswith(w) or x[1:] == w or x == w)) * 2
                 + SequenceMatcher(None, t, d).ratio())
        if score > best_score:
            best, best_score = desc, score
    return best or t


def parse_invoice(text: str) -> tuple[list[dict], dict]:
    rows, orden = [], 0
    for line in text.splitlines():
        line = line.strip()
        m_of = _of_re.match(line)
        if not m_of: continue
        of = 'MP-' + re.sub(r'^[A-Z]*-?', '', m_of.group(1))
        rest = _ref_re.sub('', line[m_of.end():]).strip().lstrip('=,—- ')
        m_c  = _comp_re.search(rest)
        desc = _best_match(rest[:m_c.start()].strip() if m_c else '')
        nums = [_eu(m.group()) for m in _num_re.finditer(rest)]
        if len(nums) < 4: continue
        try:
            pu, total = nums[-4], nums[-3]
            qte = round(total / pu) if pu > 0 else 0
        except Exception: continue
        if qte <= 0 or pu <= 0: continue
        orden += 1
        rows.append({'of': of, 'orden': orden, 'descripcion': desc,
                     'un': qte, 'pu': pu, 'total': total})

    totals = {}
    for label, key in [
        (r'N[°º]\s*DE\s*COLIS',   'n_colis'),
        (r'TOTAL\s*POIDS\s*BRUT', 'poids_brut'),
        (r'TOTAL\s*POIDS\s*NET',  'poids_net'),
        (r'TOTAL\s*FACTURE',      'total_facture'),
    ]:
        m = re.search(label + r'\s+([\d.,]+)', text, re.IGNORECASE)
        if m: totals[key] = _eu(m.group(1))

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


# ── Generador de Excel ────────────────────────────────────────────────────────
def generate_output(
    xlsx_rows, pdf_rows, pdf_bultos, pdf_totals,
    inv_number, comp_map, output_path,
) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    fn  = _fn()
    fnb = _fn(bold=True)

    # ── 1. Feuil1 ─────────────────────────────────────────────────────────────
    ws1 = wb.create_sheet('Feuil1')
    _header_row(ws1, ['Nº OF', 'ORDEN', 'Descripción', 'COMPOSICION', 'UN',
                       'Precio venta', 'VALOR', 'Consumo medio', 'Consumo total',
                       'BRUTO', 'NETO', 'BULTOS', 'Cód. ubicación'])
    RIGHT1 = {2, 5, 6, 7, 8, 9, 10, 11, 12}

    for orden, xr in enumerate(xlsx_rows, 1):
        comp_sim, _ = simplify_comp(xr['comp_raw'], comp_map)
        data = [xr['of'], orden, xr['descripcion'], comp_sim,
                xr['un'], xr['pu'], xr['valor'], xr['cons_medio'],
                xr['cons_total'], xr['bruto'], xr['neto'],
                xr['bultos'], xr['ubicacion']]
        for c, val in enumerate(data, 1):
            cell = ws1.cell(orden + 1, c, val)
            cell.fill = FILL_INPUT; cell.font = fn; cell.border = _border()
            cell.alignment = _right() if c in RIGHT1 else _wrap()

    tr = len(xlsx_rows) + 2
    for col, src in [('E','un'),('G','valor'),('I','cons_total'),
                     ('J','bruto'),('K','neto'),('L','bultos')]:
        c = ws1[f'{col}{tr}']
        c.value = round(sum((r[src] or 0) for r in xlsx_rows), 4)
        c.fill = FILL_TOTAL; c.font = fnb; c.border = _border()
        c.alignment = _right()

    ws1.cell(tr + 2, 1).value = (
        f'Factura: {inv_number}  |  '
        f'PDF COLIS: {int(pdf_totals.get("n_colis",0))}  |  '
        f'PDF BRUT: {int(pdf_totals.get("poids_brut",0))}  |  '
        f'PDF NET: {int(pdf_totals.get("poids_net",0))}'
    )
    ws1.cell(tr + 2, 1).font = _fn(italic=True, size=9, color='FF666666')

    # ── 2. PDF_Extraído ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet('PDF_Extraído')
    _header_row(ws2, ['OF (PDF)', 'ORDEN (PDF)', 'Descripción (PDF)',
                       'UN (PDF)', 'PU (PDF)', 'BULTOS (PDF)'])
    for i, pr in enumerate(pdf_rows, 2):
        data = [pr['of'], pr['orden'], pr['descripcion'],
                pr['un'], pr['pu'], pdf_bultos.get(pr['of'], '')]
        for c, val in enumerate(data, 1):
            cell = ws2.cell(i, c, val)
            cell.fill = FILL_PDF; cell.font = fn; cell.border = _border()
            cell.alignment = _right() if c in {2,4,5,6} else _wrap()

    # ── 3. Validación ─────────────────────────────────────────────────────────
    ws3 = wb.create_sheet('Validación')
    _header_row(ws3, [
        'OF', 'ORDEN',
        'Desc XLSX', 'Desc PDF', 'Desc ✓',
        'UN XLSX', 'UN PDF', 'UN ✓',
        'PU XLSX', 'PU PDF', 'PU ✓',
        'BULTOS XLSX', 'BULTOS PDF', 'BULTOS ✓',
        'COMP RAW', 'COMP Simplif.', 'COMP ✓',
    ])

    pdf_by_of: dict[str, list[dict]] = defaultdict(list)
    for pr in pdf_rows:
        pdf_by_of[pr['of']].append(pr)

    def _ok(a, b, tol=0.01):
        if a is None and b is None: return True
        if a is None or b is None: return False
        try:    return abs(float(a) - float(b)) <= tol
        except: return str(a).strip().upper() == str(b).strip().upper()

    def _check(ws, row, col, ok, warn=False):
        cell = ws.cell(row, col)
        cell.value = '✓' if ok else ('⚠' if warn else '✗')
        cell.fill  = FILL_OK if ok else (FILL_WARN if warn else FILL_ERROR)
        cell.font = fn; cell.border = _border(); cell.alignment = _center()

    def _put(ws, row, col, val, fill=FILL_INPUT, align=None):
        c = ws.cell(row, col, val)
        c.fill = fill; c.font = fn; c.border = _border()
        c.alignment = align or _wrap()

    unmapped: set[str] = set()

    for orden, xr in enumerate(xlsx_rows, 1):
        r  = orden + 1
        of = xr['of']
        comp_sim, mapped = simplify_comp(xr['comp_raw'], comp_map)
        if not mapped and xr['comp_raw']:
            unmapped.add(xr['comp_raw'])

        pr = (pdf_by_of.get(of) or [None])[0]

        ws3.cell(r, 1, of).fill = FILL_INPUT
        ws3.cell(r, 1).font = fn; ws3.cell(r, 1).border = _border()
        ws3.cell(r, 1).alignment = _wrap()
        ws3.cell(r, 2, orden).fill = FILL_INPUT
        ws3.cell(r, 2).font = fn; ws3.cell(r, 2).border = _border()
        ws3.cell(r, 2).alignment = _right()

        _put(ws3, r, 3, xr['descripcion'])
        _put(ws3, r, 4, pr['descripcion'] if pr else '—', FILL_PDF)
        _check(ws3, r, 5,
               pr is not None and _ok(xr['descripcion'], pr['descripcion']),
               warn=(pr is None))

        _put(ws3, r, 6, xr['un'], align=_right())
        _put(ws3, r, 7, pr['un'] if pr else '—', FILL_PDF, _right())
        _check(ws3, r, 8, pr is not None and _ok(xr['un'], pr['un']),
               warn=(pr is None))

        _put(ws3, r, 9,  xr['pu'], align=_right())
        _put(ws3, r, 10, pr['pu'] if pr else '—', FILL_PDF, _right())
        _check(ws3, r, 11, pr is not None and _ok(xr['pu'], pr['pu']),
               warn=(pr is None))

        pdf_bul = pdf_bultos.get(of)
        _put(ws3, r, 12, xr['bultos'], align=_right())
        _put(ws3, r, 13, pdf_bul if pdf_bul is not None else '—', FILL_PDF, _right())
        bul_ok = (xr['bultos'] is not None and pdf_bul is not None
                  and _ok(xr['bultos'], pdf_bul))
        _check(ws3, r, 14, bul_ok, warn=(pdf_bul is None))

        _put(ws3, r, 15, xr['comp_raw'])
        _put(ws3, r, 16, comp_sim)
        if not xr['comp_raw'] or mapped:
            _check(ws3, r, 17, ok=True)
        else:
            c = ws3.cell(r, 17, '⚠ sin mapear')
            c.fill = FILL_WARN; c.font = fnb
            c.border = _border(); c.alignment = _center()

    if unmapped:
        wr = len(xlsx_rows) + 3
        ws3.cell(wr, 1).value = (
            '⚠ Añadir a composicion_map.csv: '
            + ', '.join(sorted(unmapped))
        )
        ws3.cell(wr, 1).font = _fn(bold=True, color='FFCC6600')

    # ── 4. Hoja6 ──────────────────────────────────────────────────────────────
    ws6 = wb.create_sheet('Hoja6')
    _header_row(ws6, ['HS CODE', 'Descripción', 'COMPOSICION', 'ORDEN',
                       'BULTOS', 'BRUTO', 'NETO', 'VALOR', 'UN'])

    # Agrupar por (descripción canónica, composición simplificada)
    groups: dict[tuple[str,str], dict] = {}
    for orden, xr in enumerate(xlsx_rows, 1):
        comp_sim, _ = simplify_comp(xr['comp_raw'], comp_map)
        cdesc = canonical_desc(xr['descripcion'], comp_sim)
        key   = (cdesc, comp_sim)
        if key not in groups:
            groups[key] = {'orden': orden, 'bultos': 0,
                           'bruto': 0.0, 'neto': 0.0,
                           'valor': 0.0, 'un': 0}
        g = groups[key]
        g['bultos'] += (xr['bultos'] or 0)
        g['bruto']  += (xr['bruto']  or 0.0)
        g['neto']   += (xr['neto']   or 0.0)
        g['valor']  += (xr['valor']  or 0.0)
        g['un']     += (xr['un']     or 0)

    RIGHT6 = {1, 4, 5, 6, 7, 8, 9}

    for i, ((cdesc, comp_sim), g) in enumerate(groups.items(), 2):
        partida = get_partida(cdesc, comp_sim)
        row_data = [
            partida or '⚠ sin partida', cdesc, comp_sim, g['orden'],
            g['bultos'] or None,
            round(g['bruto']),
            round(g['neto'], 4),
            round(g['valor'], 2),
            g['un'],
        ]
        for c, val in enumerate(row_data, 1):
            cell = ws6.cell(i, c, val)
            cell.fill   = FILL_HOJA6
            cell.border = _border()
            cell.font   = fn if partida else FONT_WARN
            cell.alignment = _right() if c in RIGHT6 else _wrap()

    tr6 = len(groups) + 2
    for c, vals in [
        (5, [g['bultos'] for g in groups.values()]),
        (6, [round(g['bruto']) for g in groups.values()]),
        (7, [g['neto']   for g in groups.values()]),
        (8, [g['valor']  for g in groups.values()]),
        (9, [g['un']     for g in groups.values()]),
    ]:
        cell = ws6.cell(tr6, c, round(sum(vals), 4))
        cell.fill = FILL_TOTAL; cell.font = fnb
        cell.border = _border(); cell.alignment = _right()

    wb.save(output_path)
    print(f'  ✓ Excel guardado: {output_path}', flush=True)


# ── CLI ───────────────────────────────────────────────────────────────────────
def _derive_output_path(xlsx_path: Path, output: str | None) -> Path:
    stem = xlsx_path.stem.replace('ARCHIVO_EXCEL_', 'FRA_') \
           if 'ARCHIVO_EXCEL_' in xlsx_path.stem else xlsx_path.stem
    if not output:
        return xlsx_path.with_name(f'{stem}_resultado.xlsx')
    out = Path(output)
    if out.suffix.lower() == '.xlsx':
        out.parent.mkdir(parents=True, exist_ok=True)
        return out
    out.mkdir(parents=True, exist_ok=True)
    return out / f'{stem}_resultado.xlsx'


def main() -> int:
    parser = argparse.ArgumentParser(description='IFORTEX PDF + XLSX → Excel')
    parser.add_argument('--pdf',    required=True)
    parser.add_argument('--xlsx',   required=True)
    parser.add_argument('--output', default=None)
    args = parser.parse_args()

    pdf_path   = Path(args.pdf)
    xlsx_path  = Path(args.xlsx)
    out_path   = _derive_output_path(xlsx_path, args.output)
    script_dir = Path(__file__).parent

    print(f'\n=== Croton Import {SCRIPT_VERSION} ===\n', flush=True)
    print(f'📄 PDF:    {pdf_path}',  flush=True)
    print(f'📊 XLSX:   {xlsx_path}', flush=True)
    print(f'💾 Output: {out_path}',  flush=True)

    print('\n📋 Cargando composicion_map.csv...', flush=True)
    comp_map = load_composicion_map(script_dir)

    print('\n📂 Leyendo XLSX de producción...', flush=True)
    xlsx_rows = read_input_xlsx(xlsx_path)
    print(f'  ✓ {len(xlsx_rows)} filas', flush=True)
    if not xlsx_rows:
        print('❌ No se encontraron filas en Feuil1.', flush=True)
        return 1

    print('\n🔍 OCR del PDF...', flush=True)
    pages = _ocr_pages(str(pdf_path))
    print(f'  ✓ {len(pages)} páginas', flush=True)

    print('\n📊 Parseando factura (página 1)...', flush=True)
    pdf_rows, pdf_totals = parse_invoice(pages[0] if pages else '')
    inv_number = parse_invoice_number(pages[0] if pages else '')
    print(f'  ✓ {len(pdf_rows)} artículos  |  Factura: {inv_number}', flush=True)
    for pr in pdf_rows:
        print(f'    {pr["of"]:<15}  UN={pr["un"]:>5}  PU={pr["pu"]:>8.2f}  {pr["descripcion"]}',
              flush=True)

    print('\n📦 Contando bultos...', flush=True)
    pdf_bultos = parse_packing_list(pages[1:])
    for of, b in sorted(pdf_bultos.items()):
        print(f'  {of}: {b}', flush=True)

    print('\n📝 Generando Excel...', flush=True)
    generate_output(
        xlsx_rows=xlsx_rows, pdf_rows=pdf_rows,
        pdf_bultos=pdf_bultos, pdf_totals=pdf_totals,
        inv_number=inv_number, comp_map=comp_map,
        output_path=str(out_path),
    )

    print('\n✅ Listo!', flush=True)
    print(f'   Feuil1       → datos producción con COMPOSICION simplificada', flush=True)
    print(f'   PDF_Extraído → extracción OCR del PDF', flush=True)
    print(f'   Validación   → diff XLSX vs PDF (🟢 OK  🔴 error  🟡 comp sin mapear)', flush=True)
    print(f'   Hoja6        → agrupado por artículo+composición con HS code', flush=True)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
