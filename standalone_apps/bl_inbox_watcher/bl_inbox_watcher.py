#!/usr/bin/env python3
"""
bl_inbox_watcher.py
===================
Vigila /data/.../bl/inbox/ en busca de PDFs sueltos (llegados por Samba).
Cuando detecta ficheros y pasan QUIET_SECONDS sin actividad nueva,
los empaqueta en una subcarpeta batch_id/ y escribe _DONE para que
el unified_processor los recoja.

Uso:
    python bl_inbox_watcher.py
    python bl_inbox_watcher.py --inbox /data/ias_prod/data/bl/inbox --quiet 10

Instalación como servicio: ver install/bl-inbox-watcher.service.template
"""

import argparse
import logging
import os
import shutil
import time
from datetime import datetime, timezone
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("bl-watcher")


def batch_id(inbox: Path) -> str:
    ts   = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    base = f"{ts}_bl"
    candidate = base
    counter = 2
    while (inbox / candidate).exists():
        candidate = f"{base}_{counter}"
        counter += 1
    return candidate


def loose_pdfs(inbox: Path) -> list[Path]:
    """PDFs sueltos en inbox/ (no dentro de subcarpetas de batch).
    Excluye ficheros ocultos de macOS (._*), temporales y no-PDF.
    """
    return [
        p for p in inbox.iterdir()
        if p.is_file()
        and p.suffix.lower() == ".pdf"
        and not p.name.startswith("._")
        and not p.name.startswith(".")
        and not p.name.startswith("~")
    ]


def last_modified(files: list[Path]) -> float:
    """Timestamp de modificación más reciente entre los ficheros."""
    return max(p.stat().st_mtime for p in files)


def empaquetar(inbox: Path, pdfs: list[Path]) -> str:
    """Mueve los PDFs a una subcarpeta de batch y escribe _DONE."""
    bid     = batch_id(inbox)
    batch   = inbox / bid
    batch.mkdir(parents=True, exist_ok=True)

    for pdf in pdfs:
        dest = batch / pdf.name
        # Si ya existe un fichero con ese nombre en el batch, añadir sufijo numérico
        if dest.exists():
            counter = 2
            while dest.exists():
                dest = batch / f"{pdf.stem}_{counter}{pdf.suffix}"
                counter += 1
        shutil.move(str(pdf), str(dest))
        log.info("  → %s", dest.name)

    (batch / "_DONE").touch()
    log.info("Batch creado: %s (%d PDFs)", bid, len(pdfs))
    return bid


def watch(inbox: Path, quiet: int, poll: float) -> None:
    log.info("Vigilando: %s  (quietud: %ds, poll: %.1fs)", inbox, quiet, poll)
    pending_since: float = 0.0   # cuándo vimos PDFs por primera vez

    while True:
        try:
            pdfs = loose_pdfs(inbox)

            if not pdfs:
                pending_since = 0.0
                time.sleep(poll)
                continue

            now      = time.time()
            last_mod = last_modified(pdfs)

            if pending_since == 0.0:
                pending_since = now
                log.info("Detectados %d PDF(s) — esperando quietud de %ds…", len(pdfs), quiet)

            # Esperar a que no lleguen ficheros nuevos durante QUIET_SECONDS
            if (now - last_mod) >= quiet:
                log.info("Quietud alcanzada — empaquetando %d PDF(s)…", len(pdfs))
                empaquetar(inbox, pdfs)
                pending_since = 0.0
            else:
                remaining = quiet - (now - last_mod)
                log.debug("%d PDF(s) pendientes — quedan %.0fs de quietud", len(pdfs), remaining)

        except Exception as e:
            log.error("Error en el ciclo de vigilancia: %s", e, exc_info=True)

        time.sleep(poll)


def main() -> None:
    parser = argparse.ArgumentParser(description="BL Inbox Watcher")
    parser.add_argument(
        "--inbox",
        default=os.getenv("BL_INBOX", "/data/ias_prod/data/bl/inbox"),
        help="Ruta del directorio inbox a vigilar",
    )
    parser.add_argument(
        "--quiet",
        type=int,
        default=int(os.getenv("BL_QUIET_SECONDS", "15")),
        help="Segundos de inactividad antes de crear el batch (default: 15)",
    )
    parser.add_argument(
        "--poll",
        type=float,
        default=float(os.getenv("BL_POLL_SECONDS", "15")),
        help="Intervalo de comprobación en segundos (default: 3)",
    )
    args = parser.parse_args()

    inbox = Path(args.inbox)
    if not inbox.exists():
        log.error("El directorio inbox no existe: %s", inbox)
        raise SystemExit(1)

    watch(inbox, args.quiet, args.poll)


if __name__ == "__main__":
    main()