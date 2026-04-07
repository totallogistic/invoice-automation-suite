#!/usr/bin/env python3
"""
processor_dae.py  –  DAE (Declaración de Exportación) Packing List Processor
==============================================================================
Procesa el mismo XLSX fuente que camion_export_processor pero aplicando
las reglas del proceso DAE:
  - Todas las filas son declaraciones de exportación (prefijo EX: o sin prefijo)
  - Peso Bruto: ROUND(x, 0) equivalente Excel — round half-up en todas las filas
  - Sin T1 ni PDFs necesarios
  - Columnas de salida: Tour, Trailer, Shipper, MRN/Invoice,
    Peso Bruto, Peso Neto, PK, BX, PX, CL

Uso standalone:
    python processor_dae.py --xlsx PL.xlsx --output PL-DAE.xlsx

Uso integrado (desde run_processors.py):
    from processor_dae import add_dae_sheet
    add_dae_sheet(wb, result_rows, summary)
"""

from __future__ import annotations

import math
import argparse
from pathlib import Path
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ─────────────────────────────────────────────────────────────────────────────
# Shared styling (misma paleta que camion_export_processor)
# ─────────────────────────────────────────────────────────────────────────────

_GREY_HEADER = 'FFF0F0F0'
_GREEN_ROW   = 'FF66FF66'
_CYAN_TOTALS = 'FFCCFFFF'
_DARK_SEP    = 'FF6A6A6A'
_YELLOW_DATA = 'FFFFFFBB'


def _fill(hex_rgb: str) -> PatternFill:
    return PatternFill('solid', start_color=hex_rgb, end_color=hex_rgb)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _round_half_up(x: float) -> int:
    """ROUND(x, 0) equivalente a Excel: 0.5 siempre redondea hacia arriba."""
    return math.floor(x + 0.5)


def _detect_columns(header_row: tuple) -> dict:
    """
    Detecta índices de columnas a partir de la fila de cabeceras.
    Compatible con formato Kenitra (22 cols, sin BX/PX)
    y formato DAE/Tangier (23 cols, con BX/PX).
    """
    h = [str(v).strip().upper() if v else '' for v in header_row]

    def find(name: str, start: int = 0) -> Optional[int]:
        for i in range(start, len(h)):
            if h[i] == name.upper():
                return i
        return None

    pb_idx = find('PESO BRUTO')
    pk_idx = find('PK', (pb_idx or 17) + 1)

    return {
        'tour':         0,
        'trailer':      2,
        'shipper_name': 4,
        'mrn_invoice':  find('MRN / INVOICE') or 13,
        'peso_bruto':   pb_idx if pb_idx is not None else 17,
        'peso_neto':    (pb_idx or 17) + 1,
        'pk':           pk_idx if pk_idx is not None else 19,
        'bx':           find('BX', (pk_idx or 19) + 1),   # None en formato Kenitra
        'px':           find('PX', (pk_idx or 19) + 1),   # None en formato Kenitra
        'cl':           find('CL', (pk_idx or 19) + 1),
    }


def _get(row: tuple, col: dict, key: str):
    idx = col.get(key)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


# ─────────────────────────────────────────────────────────────────────────────
# Reader
# ─────────────────────────────────────────────────────────────────────────────

def read_sheet1_dae(xlsx_path: str) -> tuple[list[dict], dict]:
    """
    Lee Sheet1 del XLSX fuente para el pipeline DAE.
    Auto-detecta el layout de columnas.

    Returns:
        rows    – filas de datos (R7+), cada una con '_src_yellow'
        summary – totales de la fila de resumen (R5)
    """
    wb       = openpyxl.load_workbook(xlsx_path)
    ws       = wb['Sheet1']
    raw_rows = list(ws.iter_rows(values_only=False))
    raw      = [tuple(c.value for c in r) for r in raw_rows]

    col = _detect_columns(raw[0])

    # Fila de totales: R5 (índice 4)
    s = raw[4]
    summary = {
        'total_peso_bruto': _get(s, col, 'peso_bruto'),
        'total_peso_neto':  _get(s, col, 'peso_neto'),
        'total_pk':         _get(s, col, 'pk'),
        'total_bx':         _get(s, col, 'bx'),
        'total_px':         _get(s, col, 'px'),
        'total_cl':         _get(s, col, 'cl'),
    }

    rows = []
    for src_row_obj, raw_row in zip(raw_rows[6:], raw[6:]):   # R7+
        if all(v is None for v in raw_row):
            continue
        has_yellow = any(
            c.fill.fgColor.rgb == 'FFFFFFBB'
            for c in src_row_obj if c.fill
        )
        rows.append({
            'tour':         _get(raw_row, col, 'tour'),
            'trailer':      _get(raw_row, col, 'trailer'),
            'shipper_name': _get(raw_row, col, 'shipper_name'),
            'mrn_invoice':  _get(raw_row, col, 'mrn_invoice'),
            'peso_bruto':   _get(raw_row, col, 'peso_bruto'),
            'peso_neto':    _get(raw_row, col, 'peso_neto'),
            'pk':           _get(raw_row, col, 'pk'),
            'bx':           _get(raw_row, col, 'bx'),
            'px':           _get(raw_row, col, 'px'),
            'cl':           _get(raw_row, col, 'cl'),
            '_src_yellow':  has_yellow,
        })

    return rows, summary


# ─────────────────────────────────────────────────────────────────────────────
# Transformation
# ─────────────────────────────────────────────────────────────────────────────

def process_dae(rows: list[dict]) -> list[dict]:
    """Aplica ROUND(x,0) half-up a Peso Bruto de todas las filas."""
    result = []
    for row in rows:
        pb = row['peso_bruto']
        result.append({
            **row,
            'peso_bruto': _round_half_up(pb) if pb is not None else None,
        })
    return result


# ─────────────────────────────────────────────────────────────────────────────
# Sheet writer  (misma firma que add_processed_sheet en camion_export_processor)
# ─────────────────────────────────────────────────────────────────────────────

def add_dae_sheet(
    wb:          openpyxl.Workbook,
    result_rows: list[dict],
    summary:     dict,
    sheet_name:  str = 'DAE_RESULT',
) -> None:
    """
    Añade la hoja DAE al workbook existente.

    Layout de filas (igual al manual Sheet1_2_2):
        R1: cabeceras          → gris claro
        R2: "Name"             → gris claro
        R3: verde vacía        → verde
        R4: gris vacía         → gris claro
        R5: totales            → cyan
        R6: separador oscuro   → gris oscuro
        R7+: datos             → blanco / amarillo alternado
        última-1: separador    → amarillo
        última:   COINCIDE     → =SUM(PB) + =PK+BX+PX+CL
    """
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    ws = wb.create_sheet(sheet_name)

    arial10      = Font(name='Arial', size=10)
    arial10_bold = Font(name='Arial', size=10, bold=True)

    headers = ['Tour', 'Trailer', 'Shipper', 'MRN / Invoice',
               'Peso Bruto', 'Peso Neto', 'PK', 'BX', 'PX', 'CL']

    # R1 – cabeceras
    ws.append(headers)
    for cell in ws[1]:
        cell.font      = arial10_bold
        cell.fill      = _fill(_GREY_HEADER)
        cell.alignment = Alignment(horizontal='center')

    # R2 – "Name"
    ws.append([None, None, 'Name'] + [None] * 7)
    for cell in ws[2]:
        cell.font = arial10
        cell.fill = _fill(_GREY_HEADER)

    # R3 – verde
    ws.append([None] * 10)
    for cell in ws[3]:
        cell.fill = _fill(_GREEN_ROW)

    # R4 – gris
    ws.append([None] * 10)
    for cell in ws[4]:
        cell.fill = _fill(_GREY_HEADER)

    # R5 – totales (cyan)
    pk = int(summary['total_pk'] or 0)
    bx = int(summary['total_bx'] or 0)
    px = int(summary['total_px'] or 0)
    cl = int(summary['total_cl'] or 0)
    ws.append([None, None, None, None,
               summary['total_peso_bruto'],
               summary['total_peso_neto'],
               pk, bx, px, cl])
    for cell in ws[5]:
        cell.font = arial10_bold
        cell.fill = _fill(_CYAN_TOTALS)

    # R6 – separador oscuro
    ws.append([None] * 10)
    for cell in ws[6]:
        cell.fill = _fill(_DARK_SEP)

    # Filas de datos
    data_start_row = ws.max_row + 1

    for row in result_rows:
        excel_row_idx = ws.max_row + 1
        ws.append([
            row['tour'],
            row['trailer'],
            row['shipper_name'],
            row['mrn_invoice'],
            row['peso_bruto'],
            row['peso_neto'],
            row['pk'],
            row['bx'],
            row['px'],
            row['cl'],
        ])
        row_fill = _fill(_YELLOW_DATA) if row['_src_yellow'] else None
        for cell in ws[excel_row_idx]:
            cell.font = arial10
            if row_fill:
                cell.fill = row_fill

    data_end_row = ws.max_row

    # Separador final
    sep_idx = ws.max_row + 1
    ws.append([None] * 10)
    for cell in ws[sep_idx]:
        cell.fill = _fill(_YELLOW_DATA)

    # COINCIDE row
    coincide_row = ws.max_row + 1
    ws.append([
        None, None, None, None,
        f'=SUM(E{data_start_row}:E{data_end_row})',
        f'={pk}+{bx}+{px}+{cl}',
        None, None, None, None,
    ])
    for cell in ws[coincide_row]:
        cell.font = arial10

    # Anchos de columna
    for col_letter, width in zip('ABCDEFGHIJ', [14, 12, 35, 32, 12, 12, 8, 8, 8, 8]):
        ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = 'A7'


# ─────────────────────────────────────────────────────────────────────────────
# CLI standalone
# ─────────────────────────────────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(description='DAE Packing List Processor')
    parser.add_argument('--xlsx',          required=True, help='Input packing-list Excel')
    parser.add_argument('-o', '--output',  default=None,  help='Output Excel path')
    parser.add_argument('--verbose',       action='store_true')
    args = parser.parse_args()

    output_path = Path(args.output) if args.output else \
        Path(args.xlsx).with_name(Path(args.xlsx).stem + '_DAE.xlsx')

    print('\n=== DAE Processor ===\n')
    print(f'📊 Reading: {args.xlsx}')

    rows, summary = read_sheet1_dae(args.xlsx)
    print(f'  ✓ {len(rows)} filas de datos')
    print(f'  ✓ Totals: gross={summary["total_peso_bruto"]}  net={summary["total_peso_neto"]}'
          f'  PK={summary["total_pk"]}  BX={summary["total_bx"]}'
          f'  PX={summary["total_px"]}  CL={summary["total_cl"]}')

    result = process_dae(rows)

    if args.verbose:
        print(f'\n  {"Y":<3} {"Shipper":<28} {"MRN":<32} {"PB_src":>10} {"PB_out":>8}')
        print('  ' + '-' * 85)
        for src, out in zip(rows, result):
            ylw = '●' if src['_src_yellow'] else ''
            print(f'  {ylw:<3} {str(src["shipper_name"] or ""):<28}'
                  f' {str(src["mrn_invoice"] or ""):<32}'
                  f' {str(src["peso_bruto"]):>10} {str(out["peso_bruto"]):>8}')

    print(f'\n💾 Writing: {output_path}')
    wb = openpyxl.load_workbook(args.xlsx)
    add_dae_sheet(wb, result, summary)
    wb.save(output_path)

    print('\n✅ Done!\n')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())