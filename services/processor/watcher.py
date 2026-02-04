#!/usr/bin/env python3
from __future__ import annotations


import ssl
import logging
import os
import shutil
import subprocess
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

# ----------------------------
# Config (env)
# ----------------------------
SERVICE_ROOT = Path(os.getenv("SERVICE_ROOT", "/data"))
INBOX_DIR = Path(os.getenv("INBOX_DIR", str(SERVICE_ROOT / "inbox")))
PROCESSING_DIR = Path(os.getenv("PROCESSING_DIR", str(SERVICE_ROOT / "processing")))
OUT_DIR = Path(os.getenv("OUT_DIR", str(SERVICE_ROOT / "out")))
PROCESSED_DIR = Path(os.getenv("PROCESSED_DIR", str(SERVICE_ROOT / "processed")))
ERROR_DIR = Path(os.getenv("ERROR_DIR", str(SERVICE_ROOT / "error")))
STATUS_DIR = Path(os.getenv("STATUS_DIR", str(SERVICE_ROOT / "status")))

DONE_MARKER = os.getenv("DONE_MARKER", "_DONE")
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "3"))
BATCH_QUIET_SECONDS = int(os.getenv("BATCH_QUIET_SECONDS", "180"))

EXTRACTOR_PATH = os.getenv("EXTRACTOR_PATH", "/app/extractor/extract_lear_fields.py")

EMAIL_MODE = os.getenv("EMAIL_MODE", "BATCH_ONLY").strip().upper()
SMTP_HOST = os.getenv("SMTP_HOST", "").strip()
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_TIMEOUT = int(os.getenv("SMTP_TIMEOUT", "30"))
SMTP_USER = os.getenv("SMTP_USER", "").strip()
SMTP_PASS = os.getenv("SMTP_PASS", "").strip()
MAIL_FROM = os.getenv("MAIL_FROM", "").strip()
MAIL_TO = os.getenv("MAIL_TO", "").strip()



def parse_recipients(mail_to: str) -> list[str]:
    return [x.strip() for x in (mail_to or "").split(",") if x.strip()]

# ----------------------------
# Logging
# ----------------------------
logging.basicConfig(level=logging.INFO, format="[%(levelname)s] [watcher] %(message)s")
logger = logging.getLogger("watcher")


def ensure_structure():
    for d in [SERVICE_ROOT, INBOX_DIR, PROCESSING_DIR, OUT_DIR, PROCESSED_DIR, ERROR_DIR, STATUS_DIR]:
        d.mkdir(parents=True, exist_ok=True)
    logger.info("Estructura lista: %s", SERVICE_ROOT)
    logger.info("INBOX=%s", INBOX_DIR)
    logger.info("DONE_MARKER=%s | POLL_SECONDS=%s | BATCH_QUIET_SECONDS=%s", DONE_MARKER, POLL_SECONDS, BATCH_QUIET_SECONDS)


def _now_iso() -> str:
    return datetime.utcnow().isoformat() + "Z"


def _status_path(batch_id: str) -> Path:
    return STATUS_DIR / batch_id / "status.json"


def write_status(
    batch_id: str,
    state: str,
    stage: str,
    total_files: int | None = None,
    processed_files: int | None = None,
    recipients: list[str] | None = None,
    message: str | None = None,
):
    p = _status_path(batch_id)
    p.parent.mkdir(parents=True, exist_ok=True)

    data = {
        "batch_id": batch_id,
        "state": state,
        "stage": stage,
        "total_files": total_files if total_files is not None else 0,
        "processed_files": processed_files if processed_files is not None else 0,
        "updated_at": _now_iso(),
        "recipients": recipients or [],
        "message": message or "",
    }

    # Preserve started_at if exists
    if p.exists():
        try:
            import json

            prev = json.loads(p.read_text(encoding="utf-8"))
            if "started_at" in prev:
                data["started_at"] = prev["started_at"]
            else:
                data["started_at"] = _now_iso()
        except Exception:
            data["started_at"] = _now_iso()
    else:
        data["started_at"] = _now_iso()

    import json

    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def list_batches_in_inbox() -> list[Path]:
    # A batch is a directory directly inside INBOX
    if not INBOX_DIR.exists():
        return []
    return sorted([p for p in INBOX_DIR.iterdir() if p.is_dir()])


def batch_is_ready(batch_dir: Path) -> bool:
    """
    Ready conditions:
      - explicit DONE marker file exists in batch root (e.g. <batch>/_DONE), OR
      - no file inside batch has been modified for BATCH_QUIET_SECONDS (quiet time)
    """
    done_file = batch_dir / DONE_MARKER
    if done_file.exists():
        return True

    # Quiet-time close: if everything has been stable for >= BATCH_QUIET_SECONDS
    latest_mtime = 0.0
    for f in batch_dir.rglob("*"):
        if f.is_file():
            try:
                latest_mtime = max(latest_mtime, f.stat().st_mtime)
            except FileNotFoundError:
                continue

    if latest_mtime == 0.0:
        return False  # empty batch

    idle = time.time() - latest_mtime
    return idle >= BATCH_QUIET_SECONDS


def move_to_processing(batch_dir: Path) -> Path:
    dest = PROCESSING_DIR / batch_dir.name
    if dest.exists():
        shutil.rmtree(dest)
    shutil.move(str(batch_dir), str(dest))
    return dest


def move_to_processed(batch_id: str, processing_path: Path):
    dest = PROCESSED_DIR / batch_id
    if dest.exists():
        shutil.rmtree(dest)
    shutil.move(str(processing_path), str(dest))


def move_to_error(batch_id: str, processing_path: Path):
    dest = ERROR_DIR / batch_id
    if dest.exists():
        shutil.rmtree(dest)
    shutil.move(str(processing_path), str(dest))


def run_extractor(batch_id: str, processing_path: Path) -> Path:
    out_batch = OUT_DIR / batch_id
    out_batch.mkdir(parents=True, exist_ok=True)

    # Collect all PDFs in processing_path
    pdfs = sorted([p for p in processing_path.rglob("*.pdf") if p.is_file()])
    args = ["python3", EXTRACTOR_PATH] + [str(p) for p in pdfs] + ["-o", str(out_batch)]

    logger.info("[processor] Ejecutando extractor: %s", " ".join(args))
    cp = subprocess.run(args, capture_output=True, text=True)

    if cp.stdout:
        logger.info("[processor] STDOUT extractor:\n%s", cp.stdout.strip())
    if cp.stderr:
        logger.info("[processor] STDERR extractor:\n%s", cp.stderr.strip())

    if cp.returncode != 0:
        raise RuntimeError(f"Extractor falló (rc={cp.returncode}).")

    return out_batch


def email_config_ok() -> bool:
    return bool(SMTP_HOST and MAIL_FROM and MAIL_TO)


def send_email(subject: str, body: str, attachments: list[Path], recipients: list[str]):
    if not email_config_ok():
        logger.info("[processor] Email no configurado (SMTP_HOST/MAIL_FROM/MAIL_TO), se omite.")
        return

    if not recipients:
        logger.info("[processor] Sin destinatarios, se omite envío.")
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

    for f in attachments:
        data = f.read_bytes()
        maintype = "application"
        subtype = "octet-stream"
        # small convenience for common types
        if f.suffix.lower() == ".json":
            maintype, subtype = "application", "json"
        elif f.suffix.lower() == ".csv":
            maintype, subtype = "text", "csv"
        elif f.suffix.lower() == ".xlsx":
            maintype, subtype = "application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=f.name)

    # debug dump
    try:
        Path("/tmp/last_email.eml").write_bytes(bytes(msg))
    except Exception:
        pass
    import smtplib

    ctx = ssl.create_default_context()
    if SMTP_PORT == 465:
        smtp = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT, context=ctx)
    else:
        smtp = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT)
    with smtp:
        smtp.ehlo()
        if SMTP_PORT != 465:
            smtp.starttls(context=ctx)
            smtp.ehlo()
        smtp.login(SMTP_USER, SMTP_PASS)
        smtp.send_message(msg, from_addr=MAIL_FROM, to_addrs=recipients)


def process_batch(batch_id: str, processing_path: Path):
    # Status: processing
    pdf_count = len([p for p in processing_path.rglob("*.pdf") if p.is_file()])
    write_status(
        batch_id=batch_id,
        state="PROCESSING",
        stage="RUNNING_EXTRACTOR",
        total_files=pdf_count,
        processed_files=0,
        recipients=parse_recipients(MAIL_TO),
        message="Procesando lote...",
    )

    out_batch = run_extractor(batch_id, processing_path)

    artefacts = ["invoices_extracted.xlsx"]
    attachments = [out_batch / a for a in artefacts if (out_batch / a).exists()]

    logger.info("[processor] Extractor OK, artefactos: %s", ", ".join([p.name for p in attachments]))

    # Status: emailing
    write_status(
        batch_id=batch_id,
        state="PROCESSING",
        stage="SENDING_EMAIL",
        total_files=pdf_count,
        processed_files=pdf_count,
        recipients=parse_recipients(MAIL_TO),
        message="Extractor OK. Enviando email...",
    )

    recipients = parse_recipients(MAIL_TO)
    if recipients:
        logger.info("[processor] Enviando email a %s con %d adjuntos...", recipients, len(attachments))
    send_email(
        subject=f"[Lear Cable] Lote procesado: {batch_id}",
        body=f"Lote {batch_id} procesado correctamente.\nAdjuntos: {', '.join([p.name for p in attachments])}",
        attachments=attachments,
        recipients=recipients,
    )
    logger.info("[processor] Email enviado correctamente.")

    # Status: done
    write_status(
        batch_id=batch_id,
        state="DONE",
        stage="DONE",
        total_files=pdf_count,
        processed_files=pdf_count,
        recipients=recipients,
        message="Lote procesado y email enviado.",
    )


def watcher_loop():
    ensure_structure()
    logger.info("EMAIL_MODE=%s | RECIPIENTS=%s", EMAIL_MODE, parse_recipients(MAIL_TO))

    poll_seconds = POLL_SECONDS

    while True:
        try:
            # discover inbox batches
            batches = list_batches_in_inbox()

            for batch_dir in batches:
                batch_id = batch_dir.name

                # initialize status if missing
                sp = _status_path(batch_id)
                if not sp.exists():
                    pdf_count = len([p for p in batch_dir.rglob("*.pdf") if p.is_file()])
                    write_status(
                        batch_id=batch_id,
                        state="UPLOADED",
                        stage="WAITING",
                        total_files=pdf_count,
                        processed_files=0,
                        recipients=parse_recipients(MAIL_TO),
                        message="Lote subido. En espera de procesamiento.",
                    )

                if not batch_is_ready(batch_dir):
                    continue

                logger.info("[RUN ] Lote listo: %s", batch_id)

                processing_path = move_to_processing(batch_dir)
                try:
                    process_batch(batch_id, processing_path)
                    move_to_processed(batch_id, processing_path)
                    logger.info("[ OK ] Lote procesado: %s", batch_id)
                except Exception as e:
                    logger.error("[ERR] Lote falló: %s (%s)", batch_id, e)
                    # status to error
                    pdf_count = len([p for p in processing_path.rglob("*.pdf") if p.is_file()])
                    write_status(
                        batch_id=batch_id,
                        state="ERROR",
                        stage="ERROR",
                        total_files=pdf_count,
                        processed_files=0,
                        recipients=parse_recipients(MAIL_TO),
                        message=f"Error procesando lote: {e}",
                    )
                    move_to_error(batch_id, processing_path)

        except Exception as e:
            logger.exception("[watcher] Error en loop principal: %s", e)

        time.sleep(poll_seconds)


if __name__ == "__main__":
    watcher_loop()
