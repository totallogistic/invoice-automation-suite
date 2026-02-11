#!/usr/bin/env python3
from __future__ import annotations

import json
import os
import smtplib
import subprocess
import sys
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Any, Dict, List, Optional


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def atomic_write_json(path: Path, data: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def read_json(path: Path) -> Dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def log(msg: str) -> None:
    print(msg, flush=True)


@dataclass
class SmtpCfg:
    host: str
    port: int
    user: str
    password: str
    mail_from: str
    recipients: List[str]
    use_ssl: bool
    starttls: bool


def load_smtp_cfg() -> SmtpCfg:
    """
    Reusa el mismo set de env vars que ya tienes para Lear.
    """
    host = os.getenv("SMTP_HOST", "").strip()
    port = int((os.getenv("SMTP_PORT", "") or "0").strip() or "0")
    user = os.getenv("SMTP_USER", "").strip()
    password = os.getenv("SMTP_PASS", "").strip()
    mail_from = (os.getenv("MAIL_FROM", "") or user).strip()

    mail_to_raw = os.getenv("MAIL_TO", "").strip()
    recipients = [x.strip() for x in mail_to_raw.replace(";", ",").split(",") if x.strip()]

    use_ssl = (os.getenv("SMTP_SSL", "true").strip().lower() in ("1", "true", "yes", "on"))
    starttls = (os.getenv("SMTP_STARTTLS", "false").strip().lower() in ("1", "true", "yes", "on"))

    if not host or not port or not recipients:
        raise RuntimeError(
            f"SMTP config incompleta. Requiere SMTP_HOST/SMTP_PORT/MAIL_TO. "
            f"Actualmente: host={host!r} port={port!r} recipients={recipients!r}"
        )

    return SmtpCfg(
        host=host,
        port=port,
        user=user,
        password=password,
        mail_from=mail_from,
        recipients=recipients,
        use_ssl=use_ssl,
        starttls=starttls,
    )


def send_email_with_attachment(
    cfg: SmtpCfg,
    subject: str,
    body: str,
    attachment_path: Path,
) -> None:
    if not attachment_path.exists():
        raise RuntimeError(f"Adjunto no existe: {attachment_path}")

    msg = EmailMessage()
    msg["From"] = cfg.mail_from
    msg["To"] = ", ".join(cfg.recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    data = attachment_path.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_path.name,
    )

    if cfg.use_ssl:
        with smtplib.SMTP_SSL(cfg.host, cfg.port, timeout=30) as s:
            if cfg.user:
                s.login(cfg.user, cfg.password)
            s.send_message(msg)
    else:
        with smtplib.SMTP(cfg.host, cfg.port, timeout=30) as s:
            s.ehlo()
            if cfg.starttls:
                s.starttls()
                s.ehlo()
            if cfg.user:
                s.login(cfg.user, cfg.password)
            s.send_message(msg)


def is_batch_quiet(batch_dir: Path, quiet_seconds: int) -> bool:
    """
    Considera el batch “estable” si no hay cambios en los ficheros en quiet_seconds.
    """
    if quiet_seconds <= 0:
        return True
    latest_mtime = 0.0
    for p in batch_dir.rglob("*"):
        if p.is_file():
            latest_mtime = max(latest_mtime, p.stat().st_mtime)
    if latest_mtime == 0.0:
        return True
    return (time.time() - latest_mtime) >= quiet_seconds


def pick_single_xlsx(out_dir: Path) -> Path:
    """
    Devuelve el XLSX generado en out_dir.
    - Si hay 0 -> error
    - Si hay >1 -> el más reciente (y log warning)
    """
    xlsx = sorted(out_dir.glob("*.xlsx"))
    if not xlsx:
        raise RuntimeError(f"Extractor terminó OK pero no generó XLSX en: {out_dir}")
    if len(xlsx) == 1:
        return xlsx[0]
    # Si hay varios, cogemos el más nuevo (por seguridad)
    xlsx.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    log(f"[WARN] [import_partida] múltiples XLSX en {out_dir}, usando el más reciente: {xlsx[0].name}")
    return xlsx[0]


def main() -> int:
    service_root = Path(os.getenv("SERVICE_ROOT", "/data")).resolve()
    inbox_dir = Path(os.getenv("INBOX_DIR", str(service_root / "import_partida" / "inbox"))).resolve()
    status_dir = Path(os.getenv("STATUS_DIR", str(service_root / "import_partida" / "status"))).resolve()
    out_root = Path(os.getenv("OUT_DIR", str(service_root / "import_partida" / "out"))).resolve()

    poll_seconds = int((os.getenv("POLL_SECONDS", "3") or "3").strip())
    quiet_seconds = int((os.getenv("BATCH_QUIET_SECONDS", "30") or "30").strip())

    extractor = Path("/apps/import_partida/extractor/extract_import_partida_fields.py")

    log(f"[INFO] [import_partida] SERVICE_ROOT={service_root}")
    log(f"[INFO] [import_partida] INBOX_DIR={inbox_dir}")
    log(f"[INFO] [import_partida] STATUS_DIR={status_dir}")
    log(f"[INFO] [import_partida] OUT_DIR={out_root}")
    log(f"[INFO] [import_partida] POLL_SECONDS={poll_seconds} | BATCH_QUIET_SECONDS={quiet_seconds}")

    smtp_cfg: Optional[SmtpCfg] = None
    try:
        smtp_cfg = load_smtp_cfg()
        log(f"[INFO] [import_partida] EMAIL enabled | RECIPIENTS={smtp_cfg.recipients}")
    except Exception as e:
        log(f"[WARN] [import_partida] EMAIL disabled (SMTP config incomplete): {e}")

    inbox_dir.mkdir(parents=True, exist_ok=True)
    status_dir.mkdir(parents=True, exist_ok=True)
    out_root.mkdir(parents=True, exist_ok=True)

    processed_done: set[str] = set()

    while True:
        try:
            for batch_dir in sorted([p for p in inbox_dir.iterdir() if p.is_dir()]):
                batch_id = batch_dir.name
                if batch_id in processed_done:
                    continue

                st_file = status_dir / batch_id / "status.json"
                if not st_file.exists():
                    continue

                st = read_json(st_file)

                if st.get("state") in ("DONE", "ERROR"):
                    processed_done.add(batch_id)
                    continue

                if not is_batch_quiet(batch_dir, quiet_seconds):
                    continue

                # Import Partida: exactamente 1 PDF
                pdfs = sorted([p for p in batch_dir.glob("*.pdf") if p.is_file()])
                if len(pdfs) != 1:
                    st.update(
                        {
                            "state": "ERROR",
                            "stage": "ERROR",
                            "updated_at": _now_iso(),
                            "message": f"ERROR: se esperaba exactamente 1 PDF en el batch. Encontrados: {len(pdfs)}",
                        }
                    )
                    atomic_write_json(st_file, st)
                    processed_done.add(batch_id)
                    continue

                pdf_path = pdfs[0]
                out_dir = out_root / batch_id
                out_dir.mkdir(parents=True, exist_ok=True)

                # RUNNING
                st.update(
                    {
                        "state": "RUNNING",
                        "stage": "EXTRACTING",
                        "updated_at": _now_iso(),
                        "processed_files": 0,
                        "message": "Extrayendo campos y generando XLSX...",
                    }
                )
                atomic_write_json(st_file, st)

                # Ejecutar extractor (con args correctos)
                try:
                    if not extractor.exists():
                        raise RuntimeError(f"Extractor no existe en {extractor} (¿mount /apps en el container?)")

                    cmd = [sys.executable, str(extractor), str(pdf_path), str(out_dir)]
                    log(f"[INFO] [import_partida] run: {' '.join(cmd)}")

                    r = subprocess.run(cmd, capture_output=True, text=True)
                    if r.stdout:
                        log(f"[INFO] [import_partida] extractor stdout:\n{r.stdout.strip()}")
                    if r.stderr:
                        log(f"[WARN] [import_partida] extractor stderr:\n{r.stderr.strip()}")

                    if r.returncode != 0:
                        raise RuntimeError(f"extractor devolvió rc={r.returncode}")

                    # Validar salida
                    xlsx_path = pick_single_xlsx(out_dir)

                    st.update(
                        {
                            "state": "RUNNING",
                            "stage": "EMAILING" if smtp_cfg else "DONE",
                            "updated_at": _now_iso(),
                            "processed_files": 1,
                            "message": "XLSX generado. Enviando email..." if smtp_cfg else "XLSX generado (email deshabilitado).",
                        }
                    )
                    atomic_write_json(st_file, st)

                    # Email (si hay SMTP)
                    if smtp_cfg:
                        subject = f"[DEV] Import Partida - {batch_id}"
                        body = (
                            f"Se ha generado el XLSX para el batch {batch_id}.\n\n"
                            f"PDF: {pdf_path.name}\n"
                            f"XLSX: {xlsx_path.name}\n"
                        )
                        send_email_with_attachment(smtp_cfg, subject, body, xlsx_path)

                        st.update(
                            {
                                "recipients": smtp_cfg.recipients,
                                "updated_at": _now_iso(),
                                "message": "Email enviado correctamente.",
                            }
                        )
                        atomic_write_json(st_file, st)

                    # DONE
                    st.update(
                        {
                            "state": "DONE",
                            "stage": "DONE",
                            "updated_at": _now_iso(),
                            "processed_files": 1,
                            "message": "Procesado OK.",
                        }
                    )
                    atomic_write_json(st_file, st)
                    processed_done.add(batch_id)

                except Exception as e:
                    msg = f"ERROR: {e}"
                    log(f"[ERROR] [import_partida] {msg}")
                    log(traceback.format_exc())

                    st.update(
                        {
                            "state": "ERROR",
                            "stage": "ERROR",
                            "updated_at": _now_iso(),
                            "message": msg,
                        }
                    )
                    atomic_write_json(st_file, st)
                    processed_done.add(batch_id)

        except Exception:
            log("[ERROR] [import_partida] loop exception")
            log(traceback.format_exc())

        time.sleep(poll_seconds)


if __name__ == "__main__":
    raise SystemExit(main())
