#!/usr/bin/env python3
# __version__ = "1.0.0"
# __changelog__ =
#   1.0.0  Version inicial. Consolida PDFs en uno solo segun reglas de orden
#          (cadena -> posicion). Reglas por defecto en default_rules.json,
#          override puntual via rules.json enviado desde la web UI. Los ficheros
#          sin match van al final ordenados alfabeticamente. Match por substring
#          o regex, case-insensitive configurable. Entrega: merged.pdf (descarga
#          directa + email).
"""
merge_pdf — Consolida varios PDFs en uno solo segun reglas de orden.

Reglas:
  - Cada regla casa una cadena (match) contra el nombre del fichero y le asigna
    una posicion (order). Menor order = antes en el PDF consolidado.
  - Match por substring, case-insensitive por defecto (configurable en el JSON).
  - Los ficheros que NO casan ninguna regla van al FINAL, ordenados
    alfabeticamente entre ellos.
  - Si dos ficheros casan reglas con el mismo order, se desempata por nombre.
  - Un fichero que casa varias reglas usa la de menor order (gana la mas prioritaria).

Reglas por defecto: default_rules.json (junto a este script).
Override puntual: si el batch incluye un fichero `rules.json`, se usa ese en su lugar.

Uso (contrato del unified_processor, igual que caratula_dhl):
    python3 merge_pdf.py file1.pdf file2.pdf ... [rules.json] -o <output_dir>

  - Los PDFs llegan como argumentos posicionales (uno por fichero).
  - Si entre los paths viene un `rules.json`, se usa como override de reglas.
  - El directorio de salida llega con -o/--output.

Salida:
    <output_dir>/merged.pdf          PDF consolidado
    <output_dir>/orden.txt           orden aplicado (para trazabilidad / email)
"""

import sys
import os
import json
import re
import shutil
import argparse
import subprocess
from pathlib import Path

from pypdf import PdfReader, PdfWriter

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_RULES_PATH = SCRIPT_DIR / "default_rules.json"
OUTPUT_NAME = "merged.pdf"

# Si el PDF consolidado supera este umbral, se intenta comprimir con Ghostscript.
COMPRESS_THRESHOLD_MB = 50
# Nivel de Ghostscript: ebook = 150 dpi, buen equilibrio calidad/tamano para
# documentos mezcla texto+escaneo. (screen=72dpi mas agresivo; printer=300dpi).
GS_PDFSETTINGS = "/ebook"


# ─────────────────────────────────────────────────────────────────────────────
# Carga de reglas
# ─────────────────────────────────────────────────────────────────────────────
def load_rules(rules_override: Path = None):
    """Devuelve (rules_list, match_mode, case_sensitive, origin).

    Prioridad:
      1. rules_override  (rules.json enviado desde la UI, si existe)
      2. default_rules.json  (versionado en el repo)
    """
    src = rules_override if (rules_override and rules_override.exists()) else DEFAULT_RULES_PATH

    with open(src, "r", encoding="utf-8") as f:
        data = json.load(f)

    rules = data.get("rules", [])
    match_mode = data.get("match_mode", "substring")
    case_sensitive = bool(data.get("case_sensitive", False))

    # Normaliza y valida cada regla
    clean = []
    for r in rules:
        match = str(r.get("match", "")).strip()
        if not match:
            continue
        try:
            order = int(r.get("order"))
        except (TypeError, ValueError):
            continue
        clean.append({"match": match, "order": order})

    origin = "rules.json (UI)" if src != DEFAULT_RULES_PATH else "default_rules.json (repo)"
    return clean, match_mode, case_sensitive, origin


# ─────────────────────────────────────────────────────────────────────────────
# Asignacion de orden a cada fichero
# ─────────────────────────────────────────────────────────────────────────────
def matches(filename: str, pattern: str, match_mode: str, case_sensitive: bool) -> bool:
    name = filename if case_sensitive else filename.lower()
    pat = pattern if case_sensitive else pattern.lower()
    if match_mode == "regex":
        flags = 0 if case_sensitive else re.IGNORECASE
        try:
            return re.search(pattern, filename, flags) is not None
        except re.error:
            return False
    # default: substring
    return pat in name


def assign_order(filename: str, rules, match_mode, case_sensitive):
    """Devuelve el order de la regla mas prioritaria (menor order) que casa,
    o None si no casa ninguna."""
    best = None
    for r in rules:
        if matches(filename, r["match"], match_mode, case_sensitive):
            if best is None or r["order"] < best:
                best = r["order"]
    return best


def build_sequence(pdf_files, rules, match_mode, case_sensitive):
    """Ordena los PDFs segun las reglas. Los que NO casan ninguna regla se
    DESCARTAN del merge (comportamiento global).

    Returns: (sequence, discarded)
      sequence  : lista de (Path, order) de los ficheros que SI se consolidan,
                  ya ordenada (order asc, desempate alfabetico).
      discarded : lista de Path descartados (alfabetica) para el reporte.
    """
    matched, discarded = [], []
    for p in pdf_files:
        order = assign_order(p.name, rules, match_mode, case_sensitive)
        if order is None:
            discarded.append(p)
        else:
            matched.append((p, order))

    matched.sort(key=lambda t: (t[1], t[0].name.lower()))
    discarded.sort(key=lambda p: p.name.lower())

    sequence = [(p, o) for (p, o) in matched]
    return sequence, discarded


# ─────────────────────────────────────────────────────────────────────────────
# Merge
# ─────────────────────────────────────────────────────────────────────────────
def merge(sequence, output_path: Path):
    writer = PdfWriter()
    pages_per_file = []
    for p, _order in sequence:
        reader = PdfReader(str(p))
        n = len(reader.pages)
        for page in reader.pages:
            writer.add_page(page)
        pages_per_file.append((p.name, n))
    with open(output_path, "wb") as f:
        writer.write(f)
    return pages_per_file


def write_order_report(sequence, pages_per_file, discarded, origin, report_path: Path):
    lines = []
    lines.append(f"Reglas aplicadas desde: {origin}")
    lines.append("")
    lines.append("Orden del PDF consolidado:")
    lines.append("")
    pages_map = dict(pages_per_file)
    for i, (p, order) in enumerate(sequence, 1):
        npages = pages_map.get(p.name, "?")
        lines.append(f"  {i:>2}. {p.name}  ·  {npages} pag.  ·  regla #{order}")
    if discarded:
        lines.append("")
        lines.append(f"Descartados (sin coincidencia, NO incluidos): {len(discarded)}")
        lines.append("")
        for p in discarded:
            lines.append(f"   -  {p.name}")
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# ─────────────────────────────────────────────────────────────────────────────
# Compresion (solo si supera el umbral)
# ─────────────────────────────────────────────────────────────────────────────
def _size_mb(path: Path) -> float:
    return path.stat().st_size / (1024 * 1024)


def compress_if_needed(pdf_path: Path, threshold_mb: int = COMPRESS_THRESHOLD_MB):
    """Si pdf_path supera threshold_mb, intenta comprimir con Ghostscript.

    Devuelve (final_size_mb, accion) donde accion es uno de:
      'sin-compresion'  : no superaba el umbral
      'comprimido'      : Ghostscript redujo el tamano (se sustituye el fichero)
      'sin-mejora'      : Ghostscript no redujo -> se mantiene el original
      'gs-no-disponible': Ghostscript no instalado -> se mantiene el original
      'error-gs'        : Ghostscript fallo -> se mantiene el original

    Nunca deja el PDF en peor estado: ante cualquier problema conserva el original.
    """
    orig_mb = _size_mb(pdf_path)
    if orig_mb <= threshold_mb:
        return orig_mb, "sin-compresion"

    gs = shutil.which("gs") or shutil.which("ghostscript")
    if not gs:
        return orig_mb, "gs-no-disponible"

    tmp_out = pdf_path.with_suffix(".compressed.pdf")
    cmd = [
        gs, "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.5",
        f"-dPDFSETTINGS={GS_PDFSETTINGS}",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        "-dDetectDuplicateImages=true",
        f"-sOutputFile={tmp_out}",
        str(pdf_path),
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
    except Exception:
        if tmp_out.exists():
            tmp_out.unlink(missing_ok=True)
        return orig_mb, "error-gs"

    if result.returncode != 0 or not tmp_out.exists() or tmp_out.stat().st_size == 0:
        if tmp_out.exists():
            tmp_out.unlink(missing_ok=True)
        return orig_mb, "error-gs"

    new_mb = _size_mb(tmp_out)
    # Solo sustituimos si realmente reduce (con un margen minimo del 2%)
    if new_mb < orig_mb * 0.98:
        tmp_out.replace(pdf_path)
        return new_mb, "comprimido"
    else:
        tmp_out.unlink(missing_ok=True)
        return orig_mb, "sin-mejora"


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Consolida varios PDFs en uno solo segun reglas de orden.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  merge_pdf.py a.pdf b.pdf c.pdf -o /salida
  merge_pdf.py *.pdf rules.json -o /salida
  merge_pdf.py /ruta/al/lote --output /salida   (carpeta unica)
""",
    )
    parser.add_argument("paths", nargs="+",
                        help="PDFs a consolidar (y opcionalmente rules.json). "
                             "Tambien admite una unica carpeta.")
    parser.add_argument("-o", "--output", default=None,
                        help="Directorio de salida.")
    args = parser.parse_args()

    input_paths = [Path(p).resolve() for p in args.paths]

    missing = [p for p in input_paths if not p.exists()]
    if missing:
        for p in missing:
            print(f"ERROR: no existe: {p}", file=sys.stderr)
        sys.exit(1)

    # Si se paso una unica carpeta, expandimos su contenido
    if len(input_paths) == 1 and input_paths[0].is_dir():
        base_dir = input_paths[0]
        input_paths = sorted(base_dir.iterdir(), key=lambda p: p.name.lower())
    else:
        base_dir = input_paths[0].parent

    # Determinar salida (igual criterio que caratula_dhl)
    if args.output:
        output_dir = Path(args.output).resolve()
    elif len(args.paths) == 1 and Path(args.paths[0]).resolve().is_dir():
        output_dir = base_dir
    else:
        output_dir = base_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    # Separar rules.json de los PDFs
    rules_override = None
    pdf_files = []
    for p in input_paths:
        if not p.is_file():
            continue
        if p.name.lower() == "rules.json" or p.suffix.lower() == ".json":
            rules_override = p
        elif p.suffix.lower() == ".pdf":
            pdf_files.append(p)

    pdf_files.sort(key=lambda p: p.name.lower())

    if not pdf_files:
        print("ERROR: no se encontraron PDFs para consolidar", file=sys.stderr)
        sys.exit(1)

    rules, match_mode, case_sensitive, origin = load_rules(rules_override)

    sequence, discarded = build_sequence(pdf_files, rules, match_mode, case_sensitive)

    if not sequence:
        print("ERROR: ningun PDF coincide con las reglas; no hay nada que consolidar. "
              f"Descartados {len(discarded)} fichero(s).", file=sys.stderr)
        sys.exit(1)

    out_pdf = output_dir / OUTPUT_NAME
    pages_per_file = merge(sequence, out_pdf)

    # Comprimir si supera el umbral (solo entonces)
    final_mb, action = compress_if_needed(out_pdf)

    report = output_dir / "orden.txt"
    write_order_report(sequence, pages_per_file, discarded, origin, report)

    total_pages = sum(n for _, n in pages_per_file)
    print(f"OK -> {out_pdf}  ({len(sequence)} PDFs incluidos / {len(discarded)} descartados, "
          f"{total_pages} paginas, {final_mb:.1f} MB)")
    print(f"Reglas: {origin} · match_mode={match_mode} · case_sensitive={case_sensitive}")
    _compress_msgs = {
        "sin-compresion":   f"Tamano por debajo de {COMPRESS_THRESHOLD_MB} MB, sin comprimir.",
        "comprimido":       f"Superaba {COMPRESS_THRESHOLD_MB} MB: comprimido con Ghostscript.",
        "sin-mejora":       f"Superaba {COMPRESS_THRESHOLD_MB} MB pero Ghostscript no redujo; se mantiene original.",
        "gs-no-disponible": f"Superaba {COMPRESS_THRESHOLD_MB} MB pero Ghostscript no esta instalado; se mantiene original.",
        "error-gs":         f"Superaba {COMPRESS_THRESHOLD_MB} MB pero Ghostscript fallo; se mantiene original.",
    }
    print(f"Compresion: {_compress_msgs.get(action, action)}")
    for i, (p, order) in enumerate(sequence, 1):
        print(f"  {i:>2}. [#{order}] {p.name}")
    for p in discarded:
        print(f"   x  [descartado] {p.name}")


if __name__ == "__main__":
    main()