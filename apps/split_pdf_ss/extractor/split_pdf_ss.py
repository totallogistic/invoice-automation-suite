"""
split_pdf_ss.py
---------------
Extractor para el stack IASuite (unified_processor).

Calling convention (igual que el resto de extractores del stack):
    python3 split_pdf_ss.py file1.pdf [file2.pdf ...] -o /output/dir

Dependencias (añadir a services/unified_processor/requirements.txt):
    pypdf
    pdfplumber
"""

# TODO: implementar extracción real de documentos de Seguridad Social.
# Este fichero es un placeholder que acepta la firma del extractor y
# devuelve error hasta que la lógica esté implementada.

import argparse
import sys

SCRIPT_VERSION = "0.1.0-dev"
SCRIPT_CHANGELOG = """
### 0.1.0-dev
- Placeholder inicial. Extractor en desarrollo.
"""


def main() -> None:
    parser = argparse.ArgumentParser(description="Split PDF Seguridad Social")
    parser.add_argument("files", nargs="+", help="PDFs de entrada")
    parser.add_argument("-o", "--output", required=True, help="Directorio de salida")
    args = parser.parse_args()

    print(f"[split_pdf_ss] v{SCRIPT_VERSION}", flush=True)
    print(f"[split_pdf_ss] Archivos de entrada: {args.files}", flush=True)
    print(f"[split_pdf_ss] Directorio de salida: {args.output}", flush=True)
    print("[split_pdf_ss] ERROR: extractor aún no implementado.", flush=True)
    sys.exit(1)


if __name__ == "__main__":
    main()
