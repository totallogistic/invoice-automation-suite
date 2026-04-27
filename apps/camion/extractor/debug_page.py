#!/usr/bin/env python3
"""
debug_page.py  –  Diagnóstico de extracción de HS en UNA página específica
==============================================================================
Uso:
    python3 debug_page.py DOC.pdf 16                # pg 16 (Mastrotto)
    python3 debug_page.py DOC.pdf 16 --lang eng+ita
    python3 debug_page.py DOC.pdf 59 --lang eng+fra --dpi 200

Imprime:
  - Qué idiomas de tesseract están instalados
  - El texto OCR completo extraído de esa página
  - Qué patrón regex (si alguno) matchea
  - Todos los números de 8+ dígitos presentes (posibles HS candidates)

Útil cuando el pipeline completo no extrae un HS y quieres ver
exactamente qué está viendo tesseract y por qué la regex no matchea.
"""

from __future__ import annotations
import argparse
import re
import subprocess
import sys
from pathlib import Path

import fitz
import pytesseract
from PIL import Image


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument('pdf', help='Path al DOC.pdf')
    parser.add_argument('page', type=int, help='Número de página (1-indexed)')
    parser.add_argument('--lang', default='spa+fra+eng+ita+deu',
                        help='Idiomas tesseract (default: los 5)')
    parser.add_argument('--dpi',  type=int, default=200)
    args = parser.parse_args()

    # 1. Listar idiomas instalados
    try:
        langs = subprocess.check_output(['tesseract', '--list-langs'],
                                        stderr=subprocess.STDOUT, text=True)
        print('=== tesseract --list-langs ===')
        print(langs)
    except Exception as e:
        print(f'⚠ No se pudo listar idiomas: {e}')

    # 2. Render + OCR
    pdf_path = Path(args.pdf)
    doc = fitz.open(str(pdf_path))
    if args.page < 1 or args.page > len(doc):
        print(f'❌ Página {args.page} fuera de rango (PDF tiene {len(doc)} páginas)')
        return 1

    print(f'=== OCR pg {args.page} de {pdf_path.name} @ {args.dpi} dpi, lang={args.lang} ===')
    pix = doc[args.page - 1].get_pixmap(dpi=args.dpi)
    img = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)

    try:
        text = pytesseract.image_to_string(img, lang=args.lang)
    except pytesseract.TesseractError as e:
        print(f'❌ Tesseract error con lang={args.lang!r}: {e}')
        print('  → Probablemente falta algún idioma. Reinstala: brew install tesseract-lang')
        return 2

    print('\n--- Texto OCR completo ---')
    print(text)
    print('--- /fin texto ---\n')

    # 3. Mostrar todos los 8+ dígitos presentes
    eightdigits = sorted(set(re.findall(r'\b\d{8,10}\b', text)))
    print(f'=== Números de 8-10 dígitos (posibles HS candidates): {eightdigits} ===\n')

    # 4. Probar cada patrón de hs_extractor
    sys.path.insert(0, str(Path(__file__).parent))
    try:
        from hs_extractor import HS_PATTERNS
    except ImportError:
        print('❌ No se encuentra hs_extractor.py en el mismo directorio')
        return 3

    print(f'=== Matching contra {len(HS_PATTERNS)} patrones de hs_extractor ===')
    any_match = False
    for i, pat in enumerate(HS_PATTERNS, 1):
        for m in pat.finditer(text):
            any_match = True
            print(f'  ✓ pat#{i}  {pat.pattern[:55]!r}')
            print(f'           → HS = {m.group(1)} (HS4 = {m.group(1)[:4]})')
    if not any_match:
        print('  ✗ NINGÚN patrón matchea. Causas típicas:')
        print('    1. Falta idioma en tesseract (ej: "Tariffa" sale mangled sin ita)')
        print('    2. El layout de la factura es nuevo (añadir patrón a HS_PATTERNS)')
        print('    3. OCR de baja calidad (subir --dpi a 250 o 300)')

    return 0


if __name__ == '__main__':
    raise SystemExit(main())
