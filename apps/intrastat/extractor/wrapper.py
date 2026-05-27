#!/usr/bin/env python3
"""
Wrapper to adapt intrastat_generator.py to the unified processor interface.

Convierte:
    wrapper.py <input.xlsx|ods|csv> -o <output_dir>

En:
    intrastat_generator.py <input> -o <output_dir>/intrastat_<stem>.csv

El nombre del CSV de salida incluye el stem del fichero de entrada para que
el email adjunto sea autodescriptivo y no haya colisiones entre lotes.
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def main() -> int:
    args = sys.argv[1:]
    if "-o" not in args:
        print("ERROR: falta -o <output_dir>", file=sys.stderr)
        return 2

    o_idx = args.index("-o")
    inputs = args[:o_idx]
    if o_idx + 1 >= len(args):
        print("ERROR: -o sin valor", file=sys.stderr)
        return 2
    output_dir = Path(args[o_idx + 1])

    if len(inputs) != 1:
        print(
            f"ERROR: intrastat espera exactamente 1 fichero de entrada "
            f"(.xlsx, .ods o .csv), recibió {len(inputs)}: "
            f"{[Path(p).name for p in inputs]}",
            file=sys.stderr,
        )
        return 2

    src = Path(inputs[0])
    output_dir.mkdir(parents=True, exist_ok=True)
    out_csv = output_dir / f"intrastat_{src.stem}.csv"

    extractor = Path(__file__).parent / "intrastat_generator.py"
    if not extractor.exists():
        print(f"ERROR: no se encuentra el extractor: {extractor}", file=sys.stderr)
        return 2

    cmd = ["python3", str(extractor), str(src), "-o", str(out_csv)]
    print(f"$ {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.stdout:
        print(result.stdout, end="")
    if result.stderr:
        print(result.stderr, end="", file=sys.stderr)

    if result.returncode != 0:
        return result.returncode

    if not out_csv.exists():
        print(f"ERROR: el extractor no generó {out_csv}", file=sys.stderr)
        return 3

    print(f"OK -> {out_csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
