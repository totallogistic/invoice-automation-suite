#!/usr/bin/env python3
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

SCRIPT_VERSION = "1.0.0"
SCRIPT_CHANGELOG = """
## 1.0.0
Versión inicial. Consolida varios PDFs en uno solo (merged.pdf) según reglas de
orden (cadena → posición). Reglas por defecto en default_rules.json; override
puntual vía rules.json enviado desde la web UI. Los ficheros sin match van al
final, ordenados alfabéticamente. Match por substring o regex, case-insensitive
configurable. Entrega: descarga directa del PDF + email.
"""

import sys
import os
import json
import re
import shutil
import argparse
import subprocess
import unicodedata
from pathlib import Path

from pypdf import PdfReader, PdfWriter

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_RULES_PATH = SCRIPT_DIR / "default_rules.json"
OUTPUT_NAME = "merged.pdf"

# ── Compresion ────────────────────────────────────────────────────────────────
# Si el PDF consolidado supera este tamano, se intenta comprimir con Ghostscript.
COMPRESS_THRESHOLD_MB = 15
# Niveles de Ghostscript a probar, EN ORDEN (de menos a mas agresivo). Se prueba
# uno; si no baja del objetivo se prueba el siguiente. En cada paso solo se acepta
# el resultado si es mas pequeno que el original.
GS_LEVELS = ["/ebook", "/screen"]
# Objetivo: tamano maximo de PDF EN DISCO para que el correo (mensaje base64, ~+37%)
# quepa en el destino mas restrictivo. Gmail limita a 25 MB de MENSAJE => ~18 MB de PDF.
TARGET_MAX_MB = 18
# Marca de cliente "portal": si algun PDF incluido casa este patron, el flujo es
# portal (no se adjunta a email; el usuario lo sube manualmente al portal del cliente).
PORTAL_FILENAME_REGEX = r"Factura_LR"
# Para el nombre del PDF final de cara al cliente:
#   - identificador de factura: todo lo que va tras "Factura_" (LR26539, ALI25549...)
#   - sufijo opcional: fichero NNNN_NNN (4 digitos _ 3 digitos)
FACTURA_ID_REGEX = r"Factura_([A-Za-z0-9]+)"
CUATRO_TRES_REGEX = r"(\d{4}_\d{3})"


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
def _nfc(s: str) -> str:
    """Normaliza a NFC para que tildes combinantes (NFD, tipicas de macOS) y
    caracteres precompuestos sean equivalentes. Sin esto, 'o'+U+0301 (o + tilde)
    no casa con 'ó' (U+00F3) aunque se vean igual."""
    return unicodedata.normalize("NFC", s)


def matches(filename: str, pattern: str, match_mode: str, case_sensitive: bool) -> bool:
    fname = _nfc(filename)
    pat_raw = _nfc(pattern)
    name = fname if case_sensitive else fname.lower()
    pat = pat_raw if case_sensitive else pat_raw.lower()
    if match_mode == "regex":
        flags = 0 if case_sensitive else re.IGNORECASE
        try:
            return re.search(pat_raw, fname, flags) is not None
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


def write_order_report(sequence, pages_per_file, discarded, origin, report_path: Path,
                       final_mb=None, compress_note=None):
    lines = []
    lines.append(f"Reglas aplicadas desde: {origin}")
    if final_mb is not None:
        lines.append(f"Tamano final: {final_mb:.1f} MB")
    if compress_note:
        lines.append(f"Compresion: {compress_note}")
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


def _gs_compress(gs: str, src: Path, level: str, dst: Path, timeout=900) -> bool:
    """Ejecuta Ghostscript src->dst con el nivel dado. True si genero dst no vacio."""
    cmd = [
        gs, "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.5",
        f"-dPDFSETTINGS={level}",
        "-dNOPAUSE", "-dQUIET", "-dBATCH",
        "-dDetectDuplicateImages=true",
        f"-sOutputFile={dst}",
        str(src),
    ]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    except Exception:
        if dst.exists():
            dst.unlink(missing_ok=True)
        return False
    if r.returncode != 0 or not dst.exists() or dst.stat().st_size == 0:
        if dst.exists():
            dst.unlink(missing_ok=True)
        return False
    return True


def compress_if_needed(pdf_path: Path,
                       threshold_mb: int = COMPRESS_THRESHOLD_MB,
                       target_mb: int = TARGET_MAX_MB,
                       levels=GS_LEVELS):
    """Si pdf_path supera threshold_mb, intenta comprimir con Ghostscript probando
    los niveles en orden y quedandose con el resultado MAS PEQUENO (nunca uno mayor
    que el original). Para en cuanto baja de target_mb.

    Devuelve (final_mb, action, cabe_bool) donde action es uno de:
      'sin-compresion'  : no superaba el umbral de disparo
      'comprimido'      : se redujo (se sustituye el fichero); cabe_bool indica si <= target
      'sin-mejora'      : ningun nivel redujo -> se mantiene original (cabe_bool=False)
      'gs-no-disponible': Ghostscript no instalado -> original (cabe_bool=False)

    cabe_bool: True si el tamano final <= target_mb (cabe en el correo).
    Nunca deja el PDF en peor estado.
    """
    orig_mb = _size_mb(pdf_path)
    if orig_mb <= threshold_mb:
        return orig_mb, "sin-compresion", orig_mb <= target_mb

    gs = shutil.which("gs") or shutil.which("ghostscript")
    if not gs:
        return orig_mb, "gs-no-disponible", False

    best_path = None
    best_mb = orig_mb
    for i, level in enumerate(levels):
        tmp = pdf_path.with_suffix(f".gs{i}.pdf")
        if _gs_compress(gs, pdf_path, level, tmp):
            mb = _size_mb(tmp)
            if mb < best_mb * 0.98:          # mejora real (>2%)
                # descartar el best anterior si era un temporal
                if best_path is not None:
                    best_path.unlink(missing_ok=True)
                best_path, best_mb = tmp, mb
                if best_mb <= target_mb:     # ya cabe: no probar niveles peores
                    break
            else:
                tmp.unlink(missing_ok=True)
        # si fallo, seguimos al siguiente nivel

    if best_path is None:
        return orig_mb, "sin-mejora", False

    best_path.replace(pdf_path)
    return best_mb, "comprimido", best_mb <= target_mb


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────
def build_download_name(sequence, batch_id):
    """Nombre del PDF de cara al cliente, a partir de los ficheros incluidos:
      - {factura}_{NNNN_NNN}_full.pdf  si hay factura y fichero 4d_3d
      - {factura}_full.pdf             si solo hay factura
      - merged_{batch_id}.pdf          fallback si no hay Factura_ (no deberia pasar)
    """
    factura_id = None
    cuatro_tres = None
    for p, _ in sequence:
        name = _nfc(p.name)
        if factura_id is None:
            m = re.search(FACTURA_ID_REGEX, name, re.IGNORECASE)
            if m:
                factura_id = m.group(1)
        if cuatro_tres is None:
            m = re.search(CUATRO_TRES_REGEX, name)
            if m:
                cuatro_tres = m.group(1)

    if not factura_id:
        return f"merged_{batch_id}.pdf"
    if cuatro_tres:
        return f"{factura_id}_{cuatro_tres}_full.pdf"
    return f"{factura_id}_full.pdf"


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
    final_mb, action, cabe = compress_if_needed(out_pdf)

    _compress_msgs = {
        "sin-compresion":   f"Tamano por debajo de {COMPRESS_THRESHOLD_MB} MB, sin comprimir.",
        "comprimido":       f"Comprimido con Ghostscript a {final_mb:.1f} MB.",
        "sin-mejora":       f"Superaba {COMPRESS_THRESHOLD_MB} MB pero Ghostscript no pudo reducir; se mantiene original ({final_mb:.1f} MB).",
        "gs-no-disponible": f"Superaba {COMPRESS_THRESHOLD_MB} MB pero Ghostscript no esta instalado; se mantiene original ({final_mb:.1f} MB).",
    }
    compress_note = _compress_msgs.get(action, action)
    if not cabe and final_mb > TARGET_MAX_MB:
        compress_note += (f" AVISO: supera el objetivo de {TARGET_MAX_MB} MB para email; "
                          f"posible rechazo en el envio, descargar desde la UI.")

    report = output_dir / "orden.txt"
    write_order_report(sequence, pages_per_file, discarded, origin, report,
                       final_mb=final_mb, compress_note=compress_note)

    # ── Decision de entrega (la ejecuta el processor leyendo _DELIVERY.json) ──
    #   portal   : hay Factura_LR* -> no adjuntar; notificar para subir al portal
    #   email    : no portal y cabe (<=TARGET) -> adjuntar al interno
    #   oversize : no portal pero no cabe ni comprimido -> notificar sin adjunto
    is_portal = any(re.search(PORTAL_FILENAME_REGEX, _nfc(p.name), re.IGNORECASE)
                    for p, _ in sequence)
    if is_portal:
        delivery_mode = "portal"
    elif final_mb <= TARGET_MAX_MB:
        delivery_mode = "email"
    else:
        delivery_mode = "oversize"

    download_name = build_download_name(sequence, output_dir.name)

    delivery = {
        "mode": delivery_mode,
        "size_mb": round(final_mb, 1),
        "target_mb": TARGET_MAX_MB,
        "merged_name": OUTPUT_NAME,
        "download_name": download_name,
        "n_included": len(sequence),
        "n_discarded": len(discarded),
        "compress_note": compress_note,
    }
    (output_dir / "_DELIVERY.json").write_text(
        json.dumps(delivery, ensure_ascii=False, indent=2), encoding="utf-8")

    total_pages = sum(n for _, n in pages_per_file)
    print(f"OK -> {out_pdf}  ({len(sequence)} PDFs incluidos / {len(discarded)} descartados, "
          f"{total_pages} paginas, {final_mb:.1f} MB)")
    print(f"Reglas: {origin} · match_mode={match_mode} · case_sensitive={case_sensitive}")
    print(f"Compresion: {compress_note}")
    print(f"Entrega: modo={delivery_mode} (portal={is_portal}, cabe={final_mb <= TARGET_MAX_MB})")
    print(f"Nombre descarga/cliente: {download_name}")
    for i, (p, order) in enumerate(sequence, 1):
        print(f"  {i:>2}. [#{order}] {p.name}")
    for p in discarded:
        print(f"   x  [descartado] {p.name}")


if __name__ == "__main__":
    main()