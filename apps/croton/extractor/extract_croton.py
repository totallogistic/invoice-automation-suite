#!/usr/bin/env python3
"""
Croton ODS Processor
Receives one or more ODS files and generates a processed ODS output to be sent via email.

Usage (unified processor interface):
  extract_croton.py <file1.ods> <file2.ods> ... -o <output_dir>
"""

import argparse
import sys
from pathlib import Path

SCRIPT_VERSION = "2026-03-02.v1"

try:
    from odf import opendocument
    from odf.opendocument import load as load_ods
except ImportError:
    print("ERROR: missing dependency. Install with: pip install odfpy", file=sys.stderr)
    sys.exit(1)


def process_ods(input_path: Path, output_dir: Path) -> Path:
    """Process a single ODS input file and write output ODS.

    Replace the body of this function with the actual extraction/transformation
    logic once the processing rules are defined.
    """
    doc = load_ods(str(input_path))

    # TODO: implement extraction / transformation logic here.
    # The processed document should be saved as a new ODS file in output_dir.

    stem = input_path.stem
    out_path = output_dir / f"CROTON_{stem}.ods"
    doc.save(str(out_path))
    print(f"OK -> {out_path}")
    return out_path


def iter_ods(inputs: list) -> list:
    result = []
    for p in inputs:
        p = Path(p)
        if p.is_dir():
            result.extend(sorted(p.glob("*.ods")))
            result.extend(sorted(p.glob("*.ODS")))
        else:
            result.append(p)
    return result


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Procesa archivos ODS de Croton y genera ODS de salida."
    )
    ap.add_argument("inputs", nargs="+", type=Path, help="Archivo(s) ODS o carpeta(s)")
    ap.add_argument("-o", "--out", type=Path, default=Path("out"), help="Carpeta de salida")
    args = ap.parse_args()

    out_dir: Path = args.out
    out_dir.mkdir(parents=True, exist_ok=True)

    files = iter_ods(args.inputs)
    if not files:
        print("ERROR: No se encontraron archivos ODS.", file=sys.stderr)
        return 1

    processed = 0
    for f in files:
        if f.suffix.lower() != ".ods":
            continue
        try:
            process_ods(f, out_dir)
            processed += 1
        except Exception as e:
            print(f"[WARN] Fallo {f.name}: {e}", file=sys.stderr)

    print(f"Total procesados: {processed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
