#!/usr/bin/env python3
"""
run_processors.py  –  Orquestador: T1 + DAE
============================================
Ejecuta ambos pipelines sobre el mismo XLSX fuente y genera
un único fichero de salida con todas las hojas:

  Sheet1        – hoja original conservada
  PDF_VALIDADO  – validación XLSX vs DOC (si se pasa --doc)
  T1_RESULT     – proceso T1: tránsito aduanero (T1/EX/huérfanos, pesos del PDF)
  T1_SUMMARY    – datos extraídos de los PDFs T1
  DAE_RESULT    – proceso DAE: declaraciones de exportación (round half-up)

Modo terminal:
    python run_processors.py --xlsx PL.xlsx --t1 T1_1.pdf T1_2.pdf --doc DOC.pdf -o PL-PROCESSED.xlsx

Modo web (llamado por el stack, -o recibe un directorio):
    python run_processors.py --xlsx PL.xlsx --t1 T1_1.pdf T1_2.pdf --doc DOC.pdf -o /data/camion/out/batch_id
"""

from __future__ import annotations

SCRIPT_VERSION = "2026-03-26.v1"

SCRIPT_CHANGELOG = """
## 2026-03-26.v1

### Lógica general
Orquestador que ejecuta ambos pipelines (T1 y DAE) sobre el mismo XLSX
y genera un XLSX enriquecido con ambas hojas de resultado.

### Entradas
- `--xlsx`: Packing list del cliente (Excel)
- `--t1`: Uno o varios PDFs de declaración T1 (opcional)
- `--doc`: PDF de documento de validación (opcional)

### Hojas generadas
- Hoja original del XLSX conservada sin cambios
- `PDF_VALIDADO`: resultado de la validación XLSX vs PDF (si se pasa --doc)
- `T1_RESULT`: packing list procesado con lógica T1/EX/huérfanos
- `T1_SUMMARY`: datos extraídos de los PDFs T1
- `DAE_RESULT`: packing list procesado con lógica DAE (round half-up)
"""

import argparse
import sys
from pathlib import Path
from typing import Optional

import openpyxl

# ── Importar los dos processors ──────────────────────────────────────────────
try:
    import camion_export_processor as t1_processor
except ImportError:
    print('\n❌ No se encuentra camion_export_processor.py en el mismo directorio.\n')
    sys.exit(1)

try:
    from processor_dae import read_sheet1_dae, process_dae, add_dae_sheet
except ImportError:
    print('\n❌ No se encuentra processor_dae.py en el mismo directorio.\n')
    sys.exit(1)


# ── Output path helper (misma lógica que camion_export_processor) ────────────

def derive_output_path(xlsx_path: str, output: Optional[str]) -> Path:
    src = Path(xlsx_path)

    if not output:
        return src.with_name(f'{src.stem}-PROCESSED.xlsx')

    out = Path(output)

    if out.suffix.lower() == '.xlsx':
        out.parent.mkdir(parents=True, exist_ok=True)
        return out

    if out.exists() and out.is_dir():
        return out / f'{src.stem}-PROCESSED.xlsx'

    if out.suffix == '':
        out.mkdir(parents=True, exist_ok=True)
        return out / f'{src.stem}-PROCESSED.xlsx'

    out.parent.mkdir(parents=True, exist_ok=True)
    return out


# ── Main logic ───────────────────────────────────────────────────────────────

def run(
    xlsx_path:   str,
    t1_paths:    list[str],
    doc_path:    Optional[str],
    output_path: Path,
    verbose:     bool = False,
    dpi:         int  = 150,
) -> int:

    print('\n=== Camión Export Processor (T1 + DAE) ===\n')
    print(f'📂 Input:  {xlsx_path}')
    print(f'💾 Output: {output_path}')

    # ── 1. T1 PDFs (opcionales) ──────────────────────────────────────────────
    t1_info_list: list[dict] = []
    t1_map: dict[str, float] = {}

    if t1_paths:
        print('\n📄 Extracting T1 data...')
        for pdf_path in t1_paths:
            info = t1_processor.extract_t1_info(pdf_path)
            t1_info_list.append(info)
            if info['mrn']:
                t1_map[info['mrn']] = info['gross_kg']
                print(f"  ✓ {info['source_file']}  →  MRN: {info['mrn']}"
                      f" | Gross: {info['gross_kg']:.0f} kg"
                      f" | Pkgs: {info['packages']}"
                      f" | Deadline: {info['deadline']}")
            else:
                print(f"  ⚠ Could not extract MRN from: {pdf_path}")
    else:
        print('\nℹ️  Sin T1 PDFs – pesos T1 tomados del XLSX donde aplique.')

    # ── 2. Pipeline T1 ───────────────────────────────────────────────────────
    print('\n── Pipeline T1 (T1_RESULT) ───────────────────────────────')
    rows_t1, summary_t1 = t1_processor.read_sheet1(xlsx_path)
    print(f'  ✓ {len(rows_t1)} filas leídas')
    result_t1 = t1_processor.process_packing_list(rows_t1, t1_map)

    if verbose:
        print(f'  {"Sep":<4} {"Y":<3} {"Type":<8} {"MRN":<32} {"PB":>8}')
        for r in result_t1:
            sep = '↑' if r['_separator_before'] else ''
            ylw = '●' if r['_src_yellow'] else ''
            pb  = str(r['peso_bruto']) if r['peso_bruto'] is not None else '–'
            print(f'  {sep:<4} {ylw:<3} {r["_type"]:<8} {str(r["mrn_invoice"] or ""):<32} {pb:>8}')

    # ── 3. Validación DOC (opcional) ─────────────────────────────────────────
    report = None
    if doc_path:
        print(f'\n🔎 Validating XLSX vs DOC: {doc_path}')
        try:
            import camion_pdf_validator as pdf_validator
            report = pdf_validator.validate_xlsx_against_doc(
                xlsx_path=Path(xlsx_path),
                doc_path=Path(doc_path),
                dpi=dpi,
                t1_paths=[Path(p) for p in t1_paths],
            )
        except ImportError:
            print('  ⚠ camion_pdf_validator.py no encontrado – validación DOC omitida.')

    # ── 4. Pipeline DAE ───────────────────────────────────────────────────────
    print('\n── Pipeline DAE (DAE_RESULT) ─────────────────────────────')
    rows_dae, summary_dae = read_sheet1_dae(xlsx_path)
    print(f'  ✓ {len(rows_dae)} filas leídas')
    result_dae = process_dae(rows_dae)

    if verbose:
        print(f'  {"Y":<3} {"Shipper":<28} {"MRN":<32} {"PB_src":>10} {"PB_out":>8}')
        for src, out in zip(rows_dae, result_dae):
            ylw = '●' if src['_src_yellow'] else ''
            print(f'  {ylw:<3} {str(src["shipper_name"] or ""):<28}'
                  f' {str(src["mrn_invoice"] or ""):<32}'
                  f' {str(src["peso_bruto"]):>10} {str(out["peso_bruto"]):>8}')

    # ── 5. Workbook final ────────────────────────────────────────────────────
    print('\n💾 Escribiendo workbook final...')

    wb = openpyxl.load_workbook(xlsx_path)   # conserva Sheet1 original

    if report is not None:
        import camion_pdf_validator as pdf_validator
        pdf_validator.add_validated_sheet(wb, report, validated_sheet_name='PDF_VALIDADO')
        print('  ✓ Hoja: PDF_VALIDADO')

    t1_processor.add_processed_sheet(wb, result_t1, summary_t1, processed_sheet_name='T1_RESULT')
    print('  ✓ Hoja: T1_RESULT')

    if t1_info_list:
        t1_processor.add_t1_summary_sheet(wb, t1_info_list, sheet_name='T1_SUMMARY')
        print('  ✓ Hoja: T1_SUMMARY')

    add_dae_sheet(wb, result_dae, summary_dae, sheet_name='DAE_RESULT')
    print('  ✓ Hoja: DAE_RESULT')

    wb.save(output_path)
    print(f'\n✅ Done!')
    print(f'  ✓ XLSX: {output_path}')
    print()
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description='Camión Export Processor — T1 + DAE',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument('--xlsx',          required=True,         help='Input packing-list Excel')
    parser.add_argument('--t1',            nargs='*', default=[], help='T1 PDFs (opcionales)')
    parser.add_argument('--doc',           default=None,          help='DOC PDF para validación (opcional)')
    parser.add_argument('-o', '--output',  default=None,          help='Output: fichero .xlsx o directorio')
    parser.add_argument('--verbose',       action='store_true',   help='Detalle fila a fila')
    parser.add_argument('--dpi',           type=int, default=150, help='DPI para OCR del DOC')
    args = parser.parse_args()

    output_path = derive_output_path(args.xlsx, args.output)

    return run(
        xlsx_path   = args.xlsx,
        t1_paths    = args.t1 or [],
        doc_path    = args.doc,
        output_path = output_path,
        verbose     = args.verbose,
        dpi         = args.dpi,
    )


if __name__ == '__main__':
    raise SystemExit(main())