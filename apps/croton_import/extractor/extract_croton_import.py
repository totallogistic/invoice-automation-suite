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
- Hoja6         → agrupado por descripción, código arancelario, BRUTO entero

Composiciones:
  El fichero composicion_map.csv (mismo directorio que este script) mapea
  la composición raw del XLSX ("93% POL 7% ELASTAN") a la simplificada
  ("F SINTETICAS"). Las composiciones no mapeadas se dejan tal cual y se
  marcan en naranja en la hoja Validación para que el usuario añada la entrada.

Uso:
    python3 extract_croton_import.py \\
        --pdf  FACTURA_N_21.pdf \\
        --xlsx ARCHIVO_EXCEL_FRA_21.xlsx \\
        --output /data/out/
"""
from __future__ import annotations

SCRIPT_VERSION = "2026-03-27.v1"

SCRIPT_CHANGELOG = """
## 2026-03-27.v1
- Eliminada dependencia de template xlsx
- Entrada: PDF (OCR/validación) + XLSX (datos reales de producción)
- Composición simplificada vía composicion_map.csv mantenible
- 4 hojas de salida: Feuil1, PDF_Extraído, Validación, Hoja6
- Hoja6 con valores calculados en Python (sin fórmulas SUMIF)
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


# ── Partidas arancelarias ─────────────────────────────────────────────────────
PARTIDAS: dict[str, int] = {
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


# ── Styles ────────────────────────────────────────────────────────────────────
def _fill(rgb: str) -> PatternFill:
    return PatternFill('solid', start_color=rgb, end_color=rgb)

def _fn(bold=False, size=10, italic=False, color='FF000000') -> Font:
    return Font(bold=bold, size=size, italic=italic, name='Calibri', color=color)

def _border() -> Border:
    s = Side(style='thin', color='FF000000')
    return Border(left=s, right=s, top=s, bottom=s)

def _right()  -> Alignment: return Alignment(horizontal='right',  vertical='center')
def _center() -> Alignment: return Alignment(horizontal='center', vertical='center')
def _wrap()   -> Alignment: return Alignment(horizontal='left',   vertical='center', wrap_text=True)

# Fills
FILL_INPUT   = _fill('FFE8F4E8')   # verde claro  – datos del xlsx
FILL_OK      = _fill('FFC6EFCE')   # verde        – validación OK
FILL_ERROR   = _fill('FFFFC7CE')   # rojo         – discrepancia
FILL_WARN    = _fill('FFFFEB9C')   # naranja      – composición sin mapear
FILL_TOTAL   = _fill('FFD9E1F2')   # azul claro   – totales
FILL_HOJA6   = _fill('FFE2EFDA')   # verde pálido – Hoja6
FILL_PDF     = _fill('FFEBF3FB')   # azul muy claro – PDF_Extraído
FILL_HEADER  = _fill('FF4472C4')   # azul         – cabeceras
FONT_HEADER  = Font(bold=True, size=10, color='FFFFFFFF', name='Calibri')
FONT_WARN    = _fn(bold=True, color='FFCC0000')


def _style(cell, fill=None, font=None, align=None, fmt=None, border=True):
    if fill:   cell.fill = fill
    if font:   cell.font = font
    if align:  cell.alignment = align
    if fmt:    cell.number_format = fmt
    if border: cell.border = _border()


def _header_row(ws, cols: list[str], row=1):
    for c, label in enumerate(cols, 1):
        cell = ws.cell(row, c, label)
        cell.fill      = FILL_HEADER
        cell.font      = FONT_HEADER
        cell.border    = _border()
        cell.alignment = _center()


# ── Composición map ───────────────────────────────────────────────────────────
def load_composicion_map(script_dir: Path) -> dict[str, str]:
    """
    Carga composicion_map.csv (raw → simplificado).
    Si no existe, lo crea con ejemplos comentados.
    """
    csv_path = script_dir / 'composicion_map.csv'
    if not csv_path.exists():
        csv_path.write_text(
            'raw,simplificado\n'
            '# Añade aquí las composiciones que vayas encontrando\n'
            '# Ejemplo:\n'
            '# 93% POL 7% ELASTAN,F SINTETICAS\n'
            '# 100% POLIESTER,F SINTETICAS\n'
            '# 67% POL 33% ALG,F SINTETICAS\n'
            '# 100% ALGODON,F NATURALES\n',
            encoding='utf-8',
        )
        print(f'  ⚠ composicion_map.csv creado vacío en {csv_path}', flush=True)
        return {}

    mapping = {}
    with open(csv_path, encoding='utf-8') as f:
        for row in csv.DictReader(line for line in f if not line.strip().startswith('#')):
            raw = (row.get('raw') or '').strip().upper()
            sim = (row.get('simplificado') or '').strip()
            if raw and sim:
                mapping[raw] = sim
    print(f'  ✓ composicion_map.csv: {len(mapping)} entradas', flush=True)
    return mapping


def simplify_comp(raw: str, mapping: dict[str, str]) -> tuple[str, bool]:
    """
    Devuelve (simplificado, mapeado).
    mapeado=False → la composición no está en el CSV.
    """
    key = (raw or '').strip().upper()
    if key in mapping:
        return mapping[key], True
    return raw or '', False


# ── Leer XLSX de entrada ──────────────────────────────────────────────────────
# Columnas input (0-based): A=OF, B=ref, C=desc, D=comp_raw,
#   E=UN, F=PU, G=VALOR, H=cons_medio, I=cons_total,
#   J=BRUTO, K=NETO, L=BULTOS, M=ubicacion

def read_input_xlsx(xlsx_path: Path) -> list[dict]:
    """Lee Feuil1 del xlsx de entrada y devuelve lista de dicts."""
    wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    ws = wb['Feuil1']
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
        of = row[0]
        if not of or not isinstance(of, str) or not of.strip():
            continue
        # Fila de totales al final (OF es None o numérico)
        rows.append({
            'of':          str(of).strip(),
            'ref':         row[1],
            'descripcion': str(row[2]).strip() if row[2] else '',
            'comp_raw':    str(row[3]).strip() if row[3] else '',
            'un':          row[4],
            'pu':          row[5],
            'valor':       row[6],
            'cons_medio':  row[7],
            'cons_total':  row[8],
            'bruto':       row[9],
            'neto':        row[10],
            'bultos':      row[11],
            'ubicacion':   row[12],
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

_num_re = re.compile(r'\d+(?:[.,]\d+)*')
_of_re  = re.compile(r'^([A-Z]{0,2}-?\d{2}/\d{4,6})')
_ref_re = re.compile(r'\b[A-Z]{1,5}\d+/\S+\b')
_comp_re = re.compile(r'\d{1,3}%')


def _best_match(ocr_text: str) -> str:
    t = re.sub(r'\s+', ' ', ocr_text.upper().strip())
    if not t:
        return ''
    best, best_score = None, 0
    for desc in KNOWN_DESC:
        d = desc.upper()
        ocr_words   = set(t.split())
        desc_words  = set(d.split())
        word_score  = len(ocr_words & desc_words)
        substr_score = sum(
            1 for w in ocr_words for dw in desc_words
            if len(w) >= 3 and (dw.endswith(w) or dw[1:] == w or dw == w)
        )
        ratio = SequenceMatcher(None, t, d).ratio()
        score = word_score * 3 + substr_score * 2 + ratio
        if score > best_score:
            best, best_score = desc, score
    return best or t


def parse_invoice(text: str) -> tuple[list[dict], dict]:
    """Parsea la página de factura; devuelve (rows, totals)."""
    rows  = []
    orden = 0
    for line in text.splitlines():
        line = line.strip()
        m_of = _of_re.match(line)
        if not m_of:
            continue
        of_raw = m_of.group(1)
        of = 'MP-' + re.sub(r'^[A-Z]*-?', '', of_raw)

        rest = line[m_of.end():]
        rest = _ref_re.sub('', rest).strip().lstrip('=,—- ')

        m_c    = _comp_re.search(rest)
        desc_raw = rest[:m_c.start()].strip() if m_c else ''
        desc   = _best_match(desc_raw)

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
            'descripcion': desc,
            'un': qte, 'pu': pu, 'total': total,
        })

    totals = {}
    for label, key in [
        (r'N[°º]\s*DE\s*COLIS',    'n_colis'),
        (r'TOTAL\s*POIDS\s*BRUT',  'poids_brut'),
        (r'TOTAL\s*POIDS\s*NET',   'poids_net'),
        (r'TOTAL\s*FACTURE',       'total_facture'),
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


# ── Generador de Excel ────────────────────────────────────────────────────────

def generate_output(
    xlsx_rows:   list[dict],
    pdf_rows:    list[dict],
    pdf_bultos:  dict[str, int],
    pdf_totals:  dict,
    inv_number:  str,
    comp_map:    dict[str, str],
    output_path: str,
) -> None:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # quitar hoja por defecto

    fn  = _fn()
    fnb = _fn(bold=True)

    # ── 1. Feuil1 ─────────────────────────────────────────────────────────────
    ws1 = wb.create_sheet('Feuil1')
    cols1 = ['Nº OF', 'ORDEN', 'Descripción', 'COMPOSICION', 'UN',
             'Precio venta', 'VALOR', 'Consumo medio', 'Consumo total',
             'BRUTO', 'NETO', 'BULTOS', 'Cód. ubicación']
    _header_row(ws1, cols1)

    RIGHT_COLS1 = {2, 5, 6, 7, 8, 9, 10, 11, 12}   # 1-based

    for orden, xr in enumerate(xlsx_rows, 1):
        comp_sim, _ = simplify_comp(xr['comp_raw'], comp_map)
        data = [
            xr['of'], orden, xr['descripcion'], comp_sim,
            xr['un'], xr['pu'], xr['valor'],
            xr['cons_medio'], xr['cons_total'],
            xr['bruto'], xr['neto'], xr['bultos'], xr['ubicacion'],
        ]
        r = orden + 1
        for c, val in enumerate(data, 1):
            cell = ws1.cell(r, c, val)
            cell.fill      = FILL_INPUT
            cell.font      = fn
            cell.border    = _border()
            cell.alignment = _right() if c in RIGHT_COLS1 else _wrap()

    # Fila de totales
    tr = len(xlsx_rows) + 2
    for col_letter, src_col in [('E','un'),('G','valor'),('I','cons_total'),
                                 ('J','bruto'),('K','neto'),('L','bultos')]:
        total = sum((r[src_col] or 0) for r in xlsx_rows)
        c = ws1[f'{col_letter}{tr}']
        c.value = round(total, 4)
        c.fill  = FILL_TOTAL; c.font = fnb; c.border = _border()
        c.alignment = _right()

    # Nota de factura
    ws1.cell(tr + 2, 1).value = (
        f'Factura: {inv_number}  |  '
        f'PDF COLIS: {int(pdf_totals.get("n_colis", 0))}  |  '
        f'PDF BRUT: {int(pdf_totals.get("poids_brut", 0))}  |  '
        f'PDF NET: {int(pdf_totals.get("poids_net", 0))}'
    )
    ws1.cell(tr + 2, 1).font = _fn(italic=True, size=9, color='FF666666')

    # ── 2. PDF_Extraído ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet('PDF_Extraído')
    cols2 = ['OF (PDF)', 'ORDEN (PDF)', 'Descripción (PDF)', 'UN (PDF)', 'PU (PDF)', 'BULTOS (PDF)']
    _header_row(ws2, cols2)

    for i, pr in enumerate(pdf_rows, 2):
        bul = pdf_bultos.get(pr['of'], '')
        data = [pr['of'], pr['orden'], pr['descripcion'], pr['un'], pr['pu'], bul]
        for c, val in enumerate(data, 1):
            cell = ws2.cell(i, c, val)
            cell.fill      = FILL_PDF
            cell.font      = fn
            cell.border    = _border()
            cell.alignment = _right() if c in {2, 4, 5, 6} else _wrap()

    # Totales PDF
    tr2 = len(pdf_rows) + 2
    for c_idx, vals in [(4, [r['un'] for r in pdf_rows]), (5, [r['pu'] for r in pdf_rows])]:
        pass  # solo info, no totales en esta hoja

    # ── 3. Validación ─────────────────────────────────────────────────────────
    ws3 = wb.create_sheet('Validación')
    cols3 = [
        'OF', 'ORDEN',
        'Desc XLSX', 'Desc PDF', 'Desc ✓',
        'UN XLSX', 'UN PDF', 'UN ✓',
        'PU XLSX', 'PU PDF', 'PU ✓',
        'BULTOS XLSX', 'BULTOS PDF', 'BULTOS ✓',
        'COMP RAW', 'COMP Simplif.', 'COMP ✓',
    ]
    _header_row(ws3, cols3)

    # Indexar PDF por OF para comparación
    pdf_by_of: dict[str, list[dict]] = defaultdict(list)
    for pr in pdf_rows:
        pdf_by_of[pr['of']].append(pr)

    def _ok(a, b, tol=0.01) -> bool:
        if a is None and b is None:
            return True
        if a is None or b is None:
            return False
        try:
            return abs(float(a) - float(b)) <= tol
        except Exception:
            return str(a).strip().upper() == str(b).strip().upper()

    def _check_cell(ws, row, col, ok: bool, warn: bool = False):
        cell = ws.cell(row, col)
        cell.value = '✓' if ok else ('⚠' if warn else '✗')
        cell.fill  = FILL_OK if ok else (FILL_WARN if warn else FILL_ERROR)
        cell.font  = fn
        cell.border = _border()
        cell.alignment = _center()

    unmapped_comps: set[str] = set()

    for orden, xr in enumerate(xlsx_rows, 1):
        r = orden + 1
        of = xr['of']
        comp_sim, mapped = simplify_comp(xr['comp_raw'], comp_map)

        if not mapped and xr['comp_raw']:
            unmapped_comps.add(xr['comp_raw'])

        # Buscar contraparte en PDF (puede no existir)
        pdf_matches = pdf_by_of.get(of, [])
        pr = pdf_matches[0] if pdf_matches else None

        data = [of, orden]
        ws3.cell(r, 1, of).border   = _border()
        ws3.cell(r, 1).font         = fn
        ws3.cell(r, 1).fill         = FILL_INPUT
        ws3.cell(r, 2, orden).border = _border()
        ws3.cell(r, 2).font         = fn
        ws3.cell(r, 2).fill         = FILL_INPUT
        ws3.cell(r, 2).alignment    = _right()

        def _put(col, val, fill=FILL_INPUT, align=None):
            c = ws3.cell(r, col, val)
            c.fill = fill; c.font = fn; c.border = _border()
            c.alignment = align or _wrap()

        # Descripción
        _put(3, xr['descripcion'])
        _put(4, pr['descripcion'] if pr else '—', FILL_PDF)
        desc_ok = pr is not None and _ok(xr['descripcion'], pr['descripcion'])
        _check_cell(ws3, r, 5, desc_ok, warn=(pr is None))

        # UN
        _put(6, xr['un'], align=_right())
        _put(7, pr['un'] if pr else '—', FILL_PDF, _right())
        un_ok = pr is not None and _ok(xr['un'], pr['un'])
        _check_cell(ws3, r, 8, un_ok, warn=(pr is None))

        # PU
        _put(9, xr['pu'], align=_right())
        _put(10, pr['pu'] if pr else '—', FILL_PDF, _right())
        pu_ok = pr is not None and _ok(xr['pu'], pr['pu'])
        _check_cell(ws3, r, 11, pu_ok, warn=(pr is None))

        # Bultos
        xlsx_bul = xr['bultos']
        pdf_bul  = pdf_bultos.get(of, None)
        _put(12, xlsx_bul, align=_right())
        _put(13, pdf_bul if pdf_bul is not None else '—', FILL_PDF, _right())
        bul_ok = xlsx_bul is not None and pdf_bul is not None and _ok(xlsx_bul, pdf_bul)
        _check_cell(ws3, r, 14, bul_ok, warn=(pdf_bul is None))

        # Composición
        _put(15, xr['comp_raw'])
        _put(16, comp_sim)
        if not xr['comp_raw']:
            _check_cell(ws3, r, 17, ok=True)
        elif mapped:
            _check_cell(ws3, r, 17, ok=True)
        else:
            c = ws3.cell(r, 17, '⚠ sin mapear')
            c.fill = FILL_WARN; c.font = fnb; c.border = _border(); c.alignment = _center()

    # Advertencia de composiciones sin mapear
    if unmapped_comps:
        warn_row = len(xlsx_rows) + 3
        ws3.cell(warn_row, 1).value = (
            '⚠ Composiciones sin mapear — añadir a composicion_map.csv: '
            + ', '.join(sorted(unmapped_comps))
        )
        ws3.cell(warn_row, 1).font = _fn(bold=True, color='FFCC6600')

    # ── 4. Hoja6 ──────────────────────────────────────────────────────────────
    ws6 = wb.create_sheet('Hoja6')
    cols6 = ['Partida', 'Descripción', 'COMPOSICION', 'ORDEN',
             'BULTOS', 'BRUTO', 'NETO', 'VALOR', 'UN']
    _header_row(ws6, cols6)

    # Agrupar por descripción (mantiene primer orden, suma el resto)
    groups: dict[str, dict] = {}
    for orden, xr in enumerate(xlsx_rows, 1):
        desc     = xr['descripcion']
        comp_sim, _ = simplify_comp(xr['comp_raw'], comp_map)
        if desc not in groups:
            groups[desc] = {
                'orden':  orden,
                'comp':   comp_sim,
                'bultos': 0, 'bruto': 0.0,
                'neto':   0.0, 'valor': 0.0, 'un': 0,
            }
        groups[desc]['bultos'] += (xr['bultos'] or 0)
        groups[desc]['bruto']  += (xr['bruto']  or 0.0)
        groups[desc]['neto']   += (xr['neto']   or 0.0)
        groups[desc]['valor']  += (xr['valor']  or 0.0)
        groups[desc]['un']     += (xr['un']     or 0)

    RIGHT_COLS6 = {1, 4, 5, 6, 7, 8, 9}

    for i, (desc, g) in enumerate(groups.items(), 2):
        partida = PARTIDAS.get(desc, '')
        row_data = [
            partida or '⚠ sin partida',
            desc,
            g['comp'],
            g['orden'],
            g['bultos'] or None,
            round(g['bruto']),       # BRUTO → entero
            round(g['neto'], 4),
            round(g['valor'], 2),
            g['un'],
        ]
        for c, val in enumerate(row_data, 1):
            cell = ws6.cell(i, c, val)
            cell.fill      = FILL_HOJA6
            cell.border    = _border()
            cell.font      = fn if partida else FONT_WARN
            cell.alignment = _right() if c in RIGHT_COLS6 else _wrap()

    # Totales Hoja6
    tr6 = len(groups) + 2
    tot_data = {
        5: sum(g['bultos'] for g in groups.values()),
        6: round(sum(g['bruto']  for g in groups.values())),
        7: round(sum(g['neto']   for g in groups.values()), 4),
        8: round(sum(g['valor']  for g in groups.values()), 2),
        9: sum(g['un']    for g in groups.values()),
    }
    for c, val in tot_data.items():
        cell = ws6.cell(tr6, c, val)
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
    parser.add_argument('--pdf',    required=True,  help='Factura IFORTEX (PDF imagen)')
    parser.add_argument('--xlsx',   required=True,  help='Archivo Excel de producción (Feuil1)')
    parser.add_argument('--output', default=None,   help='Directorio o fichero .xlsx de salida')
    args = parser.parse_args()

    pdf_path  = Path(args.pdf)
    xlsx_path = Path(args.xlsx)
    out_path  = _derive_output_path(xlsx_path, args.output)
    script_dir = Path(__file__).parent

    print(f'\n=== Croton Import {SCRIPT_VERSION} ===\n', flush=True)
    print(f'📄 PDF:    {pdf_path}',  flush=True)
    print(f'📊 XLSX:   {xlsx_path}', flush=True)
    print(f'💾 Output: {out_path}',  flush=True)

    # Composición map
    print('\n📋 Cargando composicion_map.csv...', flush=True)
    comp_map = load_composicion_map(script_dir)

    # Leer XLSX input
    print('\n📂 Leyendo XLSX de producción...', flush=True)
    xlsx_rows = read_input_xlsx(xlsx_path)
    print(f'  ✓ {len(xlsx_rows)} filas', flush=True)
    if not xlsx_rows:
        print('❌ No se encontraron filas en Feuil1 del XLSX.', flush=True)
        return 1

    # OCR del PDF
    print('\n🔍 OCR del PDF...', flush=True)
    pages = _ocr_pages(str(pdf_path))
    print(f'  ✓ {len(pages)} páginas', flush=True)

    # Parsear factura (página 1)
    print('\n📊 Parseando factura (página 1)...', flush=True)
    pdf_rows, pdf_totals = parse_invoice(pages[0] if pages else '')
    inv_number = parse_invoice_number(pages[0] if pages else '')
    print(f'  ✓ {len(pdf_rows)} artículos  |  Factura: {inv_number}', flush=True)
    for pr in pdf_rows:
        print(f'    {pr["of"]:<15}  UN={pr["un"]:>5}  PU={pr["pu"]:>8.2f}  {pr["descripcion"]}',
              flush=True)

    # Packing list (páginas 2+)
    print('\n📦 Contando bultos (packing list)...', flush=True)
    pdf_bultos = parse_packing_list(pages[1:])
    for of, b in sorted(pdf_bultos.items()):
        print(f'  {of}: {b}', flush=True)

    # Generar Excel
    print('\n📝 Generando Excel...', flush=True)
    generate_output(
        xlsx_rows  = xlsx_rows,
        pdf_rows   = pdf_rows,
        pdf_bultos = pdf_bultos,
        pdf_totals = pdf_totals,
        inv_number = inv_number,
        comp_map   = comp_map,
        output_path = str(out_path),
    )

    print('\n✅ Listo!', flush=True)
    print(f'   Feuil1       → datos producción con COMPOSICION simplificada', flush=True)
    print(f'   PDF_Extraído → extracción OCR del PDF', flush=True)
    print(f'   Validación   → diff XLSX vs PDF (🟢 OK  🔴 error  🟡 comp sin mapear)', flush=True)
    print(f'   Hoja6        → agrupado por artículo con partida arancelaria', flush=True)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())