#!/usr/bin/env python3
"""
merge_pdf_watcher.py
====================
Vigila el inbox de merge_pdf en busca de PDFs sueltos (llegados por Samba desde
escaner/PC). Cuando detecta ficheros y pasan QUIET_SECONDS sin actividad nueva,
los empaqueta en una subcarpeta <batch_id>/ con un _DONE, con el MISMO formato
que crea la web UI, para que el unified_processor los recoja y consolide.

No copia rules.json: el extractor cae a default_rules.json (reglas por defecto
del servidor), que es justo lo que se quiere para este flujo automatico.

Formato de batch (identico al de la web API):
    <inbox>/<YYYYmmdd_HHMMSS_merge_pdf>/
        fichero1.pdf
        fichero2.pdf
        ...
        _DONE

Uso:
    python merge_pdf_watcher.py
    python merge_pdf_watcher.py --inbox /data/ias_prod/data/merge_pdf/inbox --quiet 15

Instalacion como servicio: ver install/merge-pdf-watcher.service.template
"""

import argparse
import logging
import os
import shutil
import time
from datetime import datetime
from pathlib import Path

try:
    from pypdf import PdfReader
    _HAS_PYPDF = True
except ImportError:
    _HAS_PYPDF = False

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("merge-pdf-watcher")

TOOL_NAME = "merge_pdf"

# Extensiones/patrones de ficheros temporales de Windows/ERP que NO son PDFs
# finales aunque acaben en .pdf o esten en la carpeta durante la copia.
TEMP_SUFFIXES = (".tmp", ".crdownload", ".part", ".partial", ".filepart")


def batch_id(inbox: Path) -> str:
    """<YYYYmmdd_HHMMSS>_merge_pdf, unico (evita colision si ya existe)."""
    base = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{TOOL_NAME}"
    candidate = base
    n = 1
    while (inbox / candidate).exists():
        candidate = f"{base}_{n}"
        n += 1
    return candidate


def loose_pdfs(inbox: Path) -> list:
    """PDFs sueltos en inbox/ (NO dentro de subcarpetas de batch).
    Excluye ocultos de macOS (._*), temporales de Windows/ERP y no-PDF.
    """
    out = []
    for p in inbox.iterdir():
        if not p.is_file():
            continue
        n = p.name
        nl = n.lower()
        # ocultos / temporales
        if n.startswith("._") or n.startswith(".") or n.startswith("~") or n.startswith("~$"):
            continue
        if nl.endswith(TEMP_SUFFIXES):
            continue
        # solo .pdf finales
        if p.suffix.lower() != ".pdf":
            continue
        out.append(p)
    return out


def is_pdf_complete(path: Path) -> bool:
    """True si el PDF parece completo (abre sin error y tiene >=1 pagina).
    Evita empaquetar ficheros a medio copiar por SMB. Si pypdf no esta
    disponible, cae a una comprobacion minima (empieza por %PDF y no vacio)."""
    try:
        if path.stat().st_size == 0:
            return False
    except OSError:
        return False

    if _HAS_PYPDF:
        try:
            reader = PdfReader(str(path))
            return len(reader.pages) >= 1
        except Exception:
            return False
    # Fallback sin pypdf: cabecera %PDF y marca EOF razonable
    try:
        with open(path, "rb") as f:
            head = f.read(5)
            if head[:4] != b"%PDF":
                return False
            f.seek(max(0, path.stat().st_size - 1024))
            tail = f.read()
            return b"%%EOF" in tail
    except OSError:
        return False


def sizes_snapshot(files: list) -> dict:
    """Mapa {path: size} para detectar ficheros que aun estan creciendo."""
    snap = {}
    for p in files:
        try:
            snap[p] = p.stat().st_size
        except OSError:
            snap[p] = -1
    return snap


def last_modified(files: list) -> float:
    return max(p.stat().st_mtime for p in files)


def empaquetar(inbox: Path, pdfs: list) -> str:
    """Mueve los PDFs a inbox/<batch_id>/ y escribe _DONE al final."""
    bid = batch_id(inbox)
    batch = inbox / bid
    batch.mkdir(parents=True, exist_ok=True)

    moved = 0
    for pdf in pdfs:
        dest = batch / pdf.name
        if dest.exists():
            dest = batch / f"{pdf.stem}_{bid[-4:]}{pdf.suffix}"
        try:
            shutil.move(str(pdf), str(dest))
            moved += 1
            log.info("  -> %s", dest.name)
        except Exception as e:
            log.error("  no se pudo mover %s: %s", pdf.name, e)

    if moved == 0:
        # nada que procesar; limpiar carpeta vacia
        try:
            batch.rmdir()
        except OSError:
            pass
        log.warning("Batch %s vacio tras el movimiento, descartado", bid)
        return ""

    # _DONE al final: senala al processor que el batch esta completo
    (batch / "_DONE").touch()
    log.info("Batch creado: %s (%d PDFs)", bid, moved)
    return bid


def watch(inbox: Path, quiet: int, poll: float) -> None:
    log.info("Vigilando: %s  (quietud: %ds, poll: %.1fs, pypdf: %s)",
             inbox, quiet, poll, "si" if _HAS_PYPDF else "no")
    pending_since = 0.0
    prev_sizes = {}

    while True:
        try:
            pdfs = loose_pdfs(inbox)

            if not pdfs:
                pending_since = 0.0
                prev_sizes = {}
                time.sleep(poll)
                continue

            now = time.time()
            last_mod = last_modified(pdfs)

            if pending_since == 0.0:
                pending_since = now
                log.info("Detectados %d PDF(s) - esperando quietud de %ds...",
                         len(pdfs), quiet)

            # Estabilidad de tamano: ningun fichero debe haber crecido desde el
            # ciclo anterior (pilla copias SMB donde el mtime ya no cambia pero
            # el fichero aun crece).
            cur_sizes = sizes_snapshot(pdfs)
            sizes_stable = (cur_sizes == prev_sizes)
            prev_sizes = cur_sizes

            quiet_ok = (now - last_mod) >= quiet

            if quiet_ok and sizes_stable:
                # Ultima barrera: validar que TODOS los PDFs estan completos.
                incompletos = [p for p in pdfs if not is_pdf_complete(p)]
                if incompletos:
                    for p in incompletos:
                        log.warning("PDF aun incompleto/ilegible, espero: %s", p.name)
                    # no reseteamos pending: reintentara en el proximo ciclo
                    time.sleep(poll)
                    continue

                log.info("Quietud + estabilidad + validez OK - empaquetando %d PDF(s)...",
                         len(pdfs))
                empaquetar(inbox, pdfs)
                pending_since = 0.0
                prev_sizes = {}
            else:
                motivo = []
                if not quiet_ok:
                    motivo.append(f"quedan {quiet - (now - last_mod):.0f}s de quietud")
                if not sizes_stable:
                    motivo.append("tamanos aun cambiando")
                log.debug("%d PDF(s) pendientes - %s", len(pdfs), "; ".join(motivo))

        except Exception as e:
            log.error("Error en el ciclo de vigilancia: %s", e, exc_info=True)

        time.sleep(poll)


def main() -> None:
    parser = argparse.ArgumentParser(description="merge_pdf Inbox Watcher")
    parser.add_argument(
        "--inbox",
        default=os.getenv("MERGE_PDF_INBOX", "/data/ias_prod/data/merge_pdf/inbox"),
        help="Ruta del directorio inbox a vigilar",
    )
    parser.add_argument(
        "--quiet",
        type=int,
        default=int(os.getenv("MERGE_PDF_QUIET_SECONDS", "30")),
        help="Segundos de inactividad antes de crear el batch (default: 30)",
    )
    parser.add_argument(
        "--poll",
        type=float,
        default=float(os.getenv("MERGE_PDF_POLL_SECONDS", "5")),
        help="Intervalo de comprobacion en segundos (default: 5)",
    )
    args = parser.parse_args()

    inbox = Path(args.inbox)
    if not inbox.exists():
        log.error("El directorio inbox no existe: %s", inbox)
        raise SystemExit(1)

    watch(inbox, args.quiet, args.poll)


if __name__ == "__main__":
    main()
