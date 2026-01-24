from __future__ import annotations

import os
import time
import shutil
import logging
import subprocess
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Sequence

import smtplib

logging.basicConfig(
    level=logging.INFO,
    format="[%(levelname)s] [%(name)s] %(message)s",
)
logger = logging.getLogger("watcher")

# ---- Layout (vía env; defaults coherentes con compose.yaml) ----
SERVICE_ROOT = Path(os.environ.get("SERVICE_ROOT", "/data"))

INBOX_DIR = Path(os.environ.get("INBOX_DIR", str(SERVICE_ROOT / "inbox")))
PROCESSING_DIR = Path(os.environ.get("PROCESSING_DIR", str(SERVICE_ROOT / "processing")))
OUT_DIR = Path(os.environ.get("OUT_DIR", str(SERVICE_ROOT / "out")))
PROCESSED_DIR = Path(os.environ.get("PROCESSED_DIR", str(SERVICE_ROOT / "processed")))
ERROR_DIR = Path(os.environ.get("ERROR_DIR", str(SERVICE_ROOT / "error")))

DONE_MARKER = os.environ.get("DONE_MARKER", "_DONE")
POLL_SECONDS = int(os.environ.get("POLL_SECONDS", "3"))

# EMAIL_MODE:
#   - BATCH_ONLY (por defecto): solo email al cerrar lote
#   - NONE: no enviar emails (aunque haya SMTP config)
EMAIL_MODE = os.environ.get("EMAIL_MODE", "BATCH_ONLY").upper()

# Extractor (tu script)
EXTRACTOR_PATH = os.environ.get("EXTRACTOR_PATH", "/app/extractor/extract_lear_fields.py")

# Config SMTP (opcional)
SMTP_HOST = os.environ.get("SMTP_HOST")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER")
SMTP_PASS = os.environ.get("SMTP_PASS")

_MAIL_TO_RAW = os.environ.get("MAIL_TO", "")
RECIPIENTS: list[str] = [a.strip() for a in _MAIL_TO_RAW.split(",") if a.strip()]


def ensure_dirs() -> None:
    """Crear estructura mínima de directorios."""
    for d in (INBOX_DIR, PROCESSING_DIR, OUT_DIR, PROCESSED_DIR, ERROR_DIR):
        d.mkdir(parents=True, exist_ok=True)
    logger.info("Estructura lista: %s", SERVICE_ROOT)


def smtp_config_ok() -> bool:
    if EMAIL_MODE == "NONE":
        return False
    if not (SMTP_HOST and SMTP_USER and SMTP_PASS and RECIPIENTS):
        logger.warning(
            "[processor] SMTP no configurado completamente; no se enviarán emails. "
            "(SMTP_HOST/USER/PASS/MAIL_TO)"
        )
        return False
    return True


def _run_extractor(batch_dir: Path, out_dir: Path) -> list[Path]:
    """
    Ejecuta extract_lear_fields.py sobre un lote (carpeta) y genera outputs en out_dir.
    Esperado: invoices_extracted.json/csv/xlsx
    """
    if not Path(EXTRACTOR_PATH).exists():
        raise FileNotFoundError(f"No encuentro extractor en {EXTRACTOR_PATH}")

    out_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        "python3",
        str(EXTRACTOR_PATH),
        str(batch_dir),
        "-o",
        str(out_dir),
    ]

    logger.info("[processor] Ejecutando extractor: %s", " ".join(cmd))
    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        logger.error("[processor] ERROR extractor:")
        if result.stdout:
            logger.error("STDOUT:\n%s", result.stdout)
        if result.stderr:
            logger.error("STDERR:\n%s", result.stderr)
        raise RuntimeError("Fallo en extract_lear_fields.py")

    # Log de salida (útil para ver SUM lines)
    if result.stdout:
        logger.info("[processor] STDOUT extractor:\n%s", result.stdout.strip())
    if result.stderr:
        logger.info("[processor] STDERR extractor:\n%s", result.stderr.strip())

    artefacts: list[Path] = []
    for name in ("invoices_extracted.json", "invoices_extracted.csv", "invoices_extracted.xlsx"):
        p = out_dir / name
        if p.exists():
            artefacts.append(p)

    logger.info(
        "[processor] Extractor OK, artefactos: %s",
        ", ".join(a.name for a in artefacts) if artefacts else "ninguno",
    )
    return artefacts


def _send_email(artefacts: Sequence[Path], subject: str, body: str) -> None:
    """Envia un email con adjuntos."""
    if not smtp_config_ok():
        return

    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = subject
    msg.set_content(body)

    for path in artefacts:
        try:
            data = path.read_bytes()
        except OSError as e:
            logger.warning("[processor] No se pudo adjuntar %s: %s", path.name, e)
            continue

        msg.add_attachment(
            data,
            maintype="application",
            subtype="octet-stream",
            filename=path.name,
        )

    logger.info("[processor] Enviando email a %s con %d adjuntos...", RECIPIENTS, len(artefacts))
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)
    logger.info("[processor] Email enviado correctamente.")


def _is_batch_ready(batch_in_inbox: Path) -> bool:
    """Un lote está listo si es carpeta y contiene DONE_MARKER."""
    if not batch_in_inbox.is_dir():
        return False
    marker = batch_in_inbox / DONE_MARKER
    return marker.exists()


def _safe_move(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    if dst.exists():
        # Evitar pisar; si existe, lo movemos con sufijo timestamp
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        dst = dst.with_name(f"{dst.name}__{ts}")
    shutil.move(str(src), str(dst))


def process_ready_batches() -> None:
    """
    Procesa todos los lotes listos en INBOX_DIR:
      inbox/<LOTE>/_DONE  --> processing/<LOTE> (lock)
      ejecuta extractor   --> out/<LOTE>/
      luego mueve         --> processed/<LOTE>
      si falla            --> error/<LOTE>
    """
    batches = sorted([p for p in INBOX_DIR.iterdir() if _is_batch_ready(p)])
    for batch in batches:
        batch_name = batch.name
        logger.info("[RUN ] Lote listo: %s", batch_name)

        processing_path = PROCESSING_DIR / batch_name
        out_path = OUT_DIR / batch_name
        processed_path = PROCESSED_DIR / batch_name
        error_path = ERROR_DIR / batch_name

        # 1) Lock: mover inbox -> processing
        try:
            _safe_move(batch, processing_path)
        except Exception as e:
            logger.error("[watcher] ERROR moviendo a processing (%s): %s", batch_name, e)
            continue

        # 2) Quitar marcador dentro de processing (ya “capturado”)
        try:
            (processing_path / DONE_MARKER).unlink(missing_ok=True)
        except Exception:
            pass

        # 3) Ejecutar extractor
        try:
            artefacts = _run_extractor(processing_path, out_path)
        except Exception as e:
            logger.error("[watcher] ERROR extractor en %s: %s", batch_name, e)
            try:
                _safe_move(processing_path, error_path)
            except Exception as e2:
                logger.error("[watcher] ERROR moviendo a error (%s): %s", batch_name, e2)
            continue

        # 4) Email (solo al cierre de lote)
        if EMAIL_MODE == "BATCH_ONLY":
            subject = f"LEAR batch listo: {batch_name}"
            body = (
                f"Lote procesado: {batch_name}\n"
                f"Outputs en: {out_path}\n\n"
                "Adjuntos: invoices_extracted.json/csv/xlsx (si existen)."
            )
            try:
                _send_email(artefacts, subject=subject, body=body)
            except Exception as e:
                logger.error("[watcher] Error enviando email (%s): %s", batch_name, e)

        # 5) Archivar PDFs / lote
        try:
            _safe_move(processing_path, processed_path)
        except Exception as e:
            logger.error("[watcher] ERROR moviendo a processed (%s): %s", batch_name, e)
            # si no podemos archivar, al menos no lo dejamos en processing para siempre
            try:
                _safe_move(processing_path, error_path)
            except Exception:
                pass
            continue

        logger.info("[ OK ] Lote procesado: %s", batch_name)


def main() -> None:
    ensure_dirs()
    logger.info("[watcher] INBOX=%s", INBOX_DIR)
    logger.info("[watcher] DONE_MARKER=%s | POLL_SECONDS=%s", DONE_MARKER, POLL_SECONDS)
    logger.info("[watcher] EMAIL_MODE=%s | RECIPIENTS=%s", EMAIL_MODE, RECIPIENTS or "Ninguno")

    while True:
        try:
            process_ready_batches()
        except Exception as e:
            logger.exception("[watcher] ERROR en loop principal: %s", e)

        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
