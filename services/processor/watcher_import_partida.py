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
            "SMTP config incompleta. Requiere SMTP_HOST/SMTP_PORT/MAIL_TO. "
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


def send_email_with_csv(
    cfg: SmtpCfg,
    subject: str,
    body: str,
    csv_path: Path,
) -> None:
    if not csv_path.exists():
        raise RuntimeError(f"Adjunto no existe: {csv_path}")

    msg = EmailMessage()
    msg["From"] = cfg.mail_from
    msg["To"] = ", ".join(cfg.recipients)
    msg["Subject"] = subject
    msg.set_content(body)

    data = csv_path.read_bytes()
    # CSV
    msg.add_attachment(
        data,
        maintype="text",
        subtype="csv",
        filename=csv_path.name,
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
    if quiet_seconds <= 0:
        return True
    latest_mtime = 0.0
    for p in batch_dir.rglob("*"):
        if p.is_file():
            latest_mtime = max(latest_mtime, p.stat().st_mtime)
    if latest_mtime == 0.0:
        return True
    return (time.time() - latest_mtime) >= quiet_seconds


def pick_single_csv(out_dir: Path) -> Path:
    """
    Para Import Partida esperamos un único CSV de salida llamado 'import_partida.csv'
    (si quieres permitir nombres variables, lo ampliamos, pero por ahora lo dejamos estricto).
    """
    expected = out_dir / "import_partida.csv"
    if expected.exists() and expected.is_file():
        return expected

    # Fallback suave: si el extractor lo generase con otro nombre pero único
    csvs = sorted([p for p in out_dir.glob("*.csv") if p.is_file()])
    if len(csvs) == 1:
        return csvs[0]

    raise RuntimeError(
        f"Extractor terminó OK pero no generó CSV esperado en: {out_dir} "
        f"(busqué import_partida.csv y/o único *.csv, encontrados={len(csvs)})"
    )


def main() -> int:
    service_root = Path(os.getenv("SERVICE_ROOT", "/data")).resolve()
    inbox_dir = Path(os.getenv("INBOX_DIR", str(service_root / "import_partida" / "inbox"))).resolve()
    status_dir = Path(os.getenv("STATUS_DIR", str(service_root / "import_partida" / "status"))).resolve()
    out_root = Path(os.getenv("OUT_DIR", str(service_root / "import_partida" / "out"))).resolve()

    poll_seconds = int((os.getenv("POLL_SECONDS", "3") or "3").strip())
    quiet_seconds = int((os.getenv("BATCH_QUIET_SECONDS", "30") or "30").strip())

    extractor_timeout = int((os.getenv("EXTRACTOR_TIMEOUT_SECONDS", "120") or "120").strip())

    extractor = Path("/apps/import_partida/extractor/extract_import_partida_fields.py")

    log(f"[INFO] [import_partida] SERVICE_ROOT={service_root}")
    log(f"[INFO] [import_partida] INBOX_DIR={inbox_dir}")
    log(f"[INFO] [import_partida] STATUS_DIR={status_dir}")
    log(f"[INFO] [import_partida] OUT_DIR={out_root}")
    log(f"[INFO] [import_partida] POLL_SECONDS={poll_seconds} | BATCH_QUIET_SECONDS={quiet_seconds} | EXTRACTOR_TIMEOUT_SECONDS={extractor_timeout}")

    smtp_cfg: Optional[SmtpCfg] = None
    try:
        smtp_cfg = load_smtp_cfg()
        log(f"[INFO] [import_partida] EMAIL enabled | RECIPIENTS={smtp_cfg.recipients}")
    except Exception as e:
        log(f"[WARN] [import_partida] EMAIL disabled (SMTP config incomplete): {e}")

    inbox_dir.mkdir(parents=True, exist_ok=True)
    status_dir.mkdir(parents=True, exist_ok=True)
    out_root.mkdir(parents=True, exist_ok=True)

    processed_done = set()

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
                        "message": "Extrayendo campos y generando CSV...",
                    }
                )
                atomic_write_json(st_file, st)

                try:
                    if not extractor.exists():
                        raise RuntimeError(f"Extractor no existe en {extractor} (¿mount /apps en el container?)")

                    cmd = [sys.executable, str(extractor), str(pdf_path), str(out_dir)]
                    log(f"[INFO] [import_partida] run: {' '.join(cmd)}")

                    try:
                        r = subprocess.run(cmd, capture_output=True, text=True, timeout=extractor_timeout)
                    except subprocess.TimeoutExpired:
                        raise RuntimeError(f"extractor timeout después de {extractor_timeout}s")

                    if r.stdout:
                        log(f"[INFO] [import_partida] extractor stdout:\n{r.stdout.strip()}")
                    if r.stderr:
                        log(f"[WARN] [import_partida] extractor stderr:\n{r.stderr.strip()}")

                    if r.returncode != 0:
                        raise RuntimeError(f"extractor devolvió rc={r.returncode}")

                    # Validar salida CSV
                    csv_path = pick_single_csv(out_dir)

                    st.update(
                        {
                            "state": "RUNNING",
                            "stage": "EMAILING" if smtp_cfg else "DONE",
                            "updated_at": _now_iso(),
                            "processed_files": 1,
                            "message": "CSV generado. Enviando email..." if smtp_cfg else "CSV generado (email deshabilitado).",
                        }
                    )
                    atomic_write_json(st_file, st)

                    if smtp_cfg:
                        subject = f"Import Partida - {batch_id}"
                        body = (
                            f"Se ha generado el CSV para el batch {batch_id}.\n\n"
                            f"PDF: {pdf_path.name}\n"
                            f"CSV: {csv_path.name}\n"
                        )
                        send_email_with_csv(smtp_cfg, subject, body, csv_path)

                        st.update(
                            {
                                "recipients": smtp_cfg.recipients,
                                "updated_at": _now_iso(),
                                "message": "Email enviado correctamente.",
                            }
                        )
                        atomic_write_json(st_file, st)

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
