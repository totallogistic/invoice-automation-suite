#!/usr/bin/env python3
"""
dua_watcher.py
==============
Vigila un inbox (alimentado por Samba desde el escaner) con PDFs de DUAs de
aduanas. Por cada PDF:

  1. Espera quietud + estabilidad de tamaño + PDF completo (como el merge_pdf
     watcher) para no procesar ficheros a medio copiar por SMB.
  2. OCR de la 1ª pagina (validado contra el resto) para extraer el nº DUA.
  3. Segun la CONFIANZA del OCR:
       - alta   -> renombra a <DUA>.pdf y lo deja en OUTBOX (Samba salida ->
                   OneDrive -> SharePoint). Registra en historial.
       - dudoso -> deja el PDF en REVIEW con un <DUA>.propuesto.txt y un recorte
                   PNG de la zona del DUA; envia email interno al operador para
                   que confirme/corrija y lo coloque en OUTBOX. Registra como
                   "dudoso" en historial.
  4. Conserva SIEMPRE el original en PROCESSED (traza, como el resto del stack).
  5. Historial JSON append: DUA final, nombre original, fecha/hora, confianza,
     nº paginas.

El email interno reutiliza SMTP del stack (igual que merge_pdf): variables
SMTP_HOST/PORT/USER/PASS, MAIL_FROM, y destinatario en MAIL_INTERNO_DUA.

Uso:
    python3 dua_watcher.py
    python3 dua_watcher.py --inbox /data/ias_prod/data/dua/inbox --quiet 30
"""

import argparse
import json
import logging
import os
import shutil
import smtplib
import time
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

from pdf2image import convert_from_path

try:
    from pypdf import PdfReader
    _HAS_PYPDF = True
except ImportError:
    _HAS_PYPDF = False

from dua_extract import extract_dua

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("dua-watcher")

TEMP_SUFFIXES = (".tmp", ".crdownload", ".part", ".partial", ".filepart")


# ── Deteccion de PDFs estables (igual que merge_pdf_watcher) ──────────────────
def loose_pdfs(inbox: Path) -> list:
    out = []
    for p in inbox.iterdir():
        if not p.is_file():
            continue
        n = p.name
        nl = n.lower()
        if n.startswith("._") or n.startswith(".") or n.startswith("~") or n.startswith("~$"):
            continue
        if nl.endswith(TEMP_SUFFIXES):
            continue
        if p.suffix.lower() != ".pdf":
            continue
        out.append(p)
    return out


def is_pdf_complete(path: Path) -> bool:
    try:
        if path.stat().st_size == 0:
            return False
    except OSError:
        return False
    if _HAS_PYPDF:
        try:
            return len(PdfReader(str(path)).pages) >= 1
        except Exception:
            return False
    try:
        with open(path, "rb") as f:
            if f.read(5)[:4] != b"%PDF":
                return False
            f.seek(max(0, path.stat().st_size - 1024))
            return b"%%EOF" in f.read()
    except OSError:
        return False


def sizes_snapshot(files):
    snap = {}
    for p in files:
        try:
            snap[p] = p.stat().st_size
        except OSError:
            snap[p] = -1
    return snap


# ── Historial ─────────────────────────────────────────────────────────────────
def append_history(history_path: Path, record: dict):
    """Append al historial JSON (lista). Robusto si el fichero no existe o esta
    corrupto: en ese caso empieza lista nueva sin perder el registro actual."""
    data = []
    if history_path.exists():
        try:
            data = json.loads(history_path.read_text(encoding="utf-8"))
            if not isinstance(data, list):
                data = []
        except Exception:
            data = []
    data.append(record)
    tmp = history_path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(history_path)


# ── Recorte de la zona del DUA (ayuda visual para revision) ───────────────────
def save_dua_crop(pdf_path: Path, dest_png: Path):
    """Guarda un PNG con la cabecera superior-derecha de la 1ª pagina (donde
    esta 'Num. Dua'), para que el operador confirme de un vistazo."""
    try:
        imgs = convert_from_path(str(pdf_path), dpi=200, first_page=1, last_page=1)
        if not imgs:
            return False
        img = imgs[0]
        w, h = img.size
        crop = img.crop((int(w * 0.45), 0, w, int(h * 0.15)))
        crop.save(str(dest_png))
        return True
    except Exception as e:
        log.warning("No se pudo generar recorte DUA: %s", e)
        return False


# ── Email interno (revision) ──────────────────────────────────────────────────
def send_review_email(pdf_path: Path, dua: str, motivo: str, crop_png: Path):
    host = os.getenv("SMTP_HOST", "").strip()
    to = os.getenv("MAIL_INTERNO_DUA", "").strip()
    if not host or not to:
        log.warning("Sin SMTP_HOST o MAIL_INTERNO_DUA; no se envia email de revision")
        return False
    port = int(os.getenv("SMTP_PORT", "25"))
    user = os.getenv("SMTP_USER", "").strip()
    pw = os.getenv("SMTP_PASS", "").strip()
    mail_from = os.getenv("MAIL_FROM", user or "dua-watcher@localhost")

    msg = EmailMessage()
    msg["Subject"] = f"[DUA] Revision manual — {pdf_path.name}"
    msg["From"] = mail_from
    msg["To"] = to
    msg.set_content(
        f"El DUA de este documento necesita CONFIRMACION manual antes de subir.\n\n"
        f"  Fichero original : {pdf_path.name}\n"
        f"  DUA propuesto    : {dua or '(no detectado)'}\n"
        f"  Motivo           : {motivo}\n\n"
        f"ACCION: confirma o corrige el numero DUA mirando el recorte adjunto, "
        f"renombra el PDF a <DUA>.pdf y colocalo en la carpeta de salida.\n"
        f"El original queda conservado en el servidor.\n"
    )
    # adjuntar PDF + recorte
    try:
        msg.add_attachment(pdf_path.read_bytes(), maintype="application",
                           subtype="pdf", filename=pdf_path.name)
        if crop_png and crop_png.exists():
            msg.add_attachment(crop_png.read_bytes(), maintype="image",
                               subtype="png", filename="dua_zona.png")
    except Exception as e:
        log.warning("No se pudo adjuntar al email: %s", e)

    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            if user and pw:
                try:
                    s.starttls()
                except Exception:
                    pass
                s.login(user, pw)
            s.send_message(msg)
        log.info("Email de revision enviado a %s", to)
        return True
    except Exception as e:
        log.error("Fallo enviando email de revision: %s", e)
        return False


# ── Procesar un PDF ───────────────────────────────────────────────────────────
def process_pdf(pdf_path: Path, dirs: dict):
    log.info("Procesando: %s", pdf_path.name)
    res = extract_dua(str(pdf_path))
    dua = res["dua"]
    confianza = res["confianza"]

    ts = datetime.now()
    record = {
        "fecha": ts.isoformat(timespec="seconds"),
        "fichero_original": pdf_path.name,
        "dua": dua,
        "confianza": confianza,
        "motivo": res["motivo"],
        "n_paginas": res["n_paginas"],
        "paginas_con_dua": res["paginas_con_dua"],
    }

    # Conservar SIEMPRE el original en processed/
    processed_dest = dirs["processed"] / pdf_path.name
    if processed_dest.exists():
        processed_dest = dirs["processed"] / f"{pdf_path.stem}_{ts.strftime('%H%M%S')}{pdf_path.suffix}"
    shutil.copy2(str(pdf_path), str(processed_dest))

    if confianza == "alta" and dua:
        # Renombrar y a la bandeja de salida (OneDrive la sube)
        out = dirs["outbox"] / f"{dua}.pdf"
        if out.exists():
            out = dirs["outbox"] / f"{dua}_{ts.strftime('%H%M%S')}.pdf"
        shutil.move(str(pdf_path), str(out))
        record["destino"] = "outbox"
        record["nombre_final"] = out.name
        log.info("  OK [alta] -> %s", out.name)
    else:
        # Dudoso: a revision + email
        review_dir = dirs["review"]
        base = dua or pdf_path.stem
        review_pdf = review_dir / f"{base}.pdf"
        if review_pdf.exists():
            review_pdf = review_dir / f"{base}_{ts.strftime('%H%M%S')}.pdf"
        shutil.move(str(pdf_path), str(review_pdf))
        # txt con la propuesta + recorte
        (review_pdf.with_suffix(".propuesto.txt")).write_text(
            f"DUA propuesto: {dua or '(no detectado)'}\n"
            f"Confianza    : {confianza}\n"
            f"Motivo       : {res['motivo']}\n"
            f"Lecturas por pagina:\n" +
            "\n".join(f"  pag {pg}: {cs}" for pg, cs in res["lecturas"].items()),
            encoding="utf-8",
        )
        crop_png = review_pdf.with_suffix(".dua_zona.png")
        save_dua_crop(review_pdf, crop_png)
        send_review_email(review_pdf, dua, res["motivo"], crop_png)
        record["destino"] = "review"
        record["nombre_final"] = review_pdf.name
        log.info("  DUDOSO -> revision: %s (motivo: %s)", review_pdf.name, res["motivo"])

    append_history(dirs["history"], record)
    return record


# ── Loop de vigilancia ────────────────────────────────────────────────────────
def watch(inbox: Path, dirs: dict, quiet: int, poll: float):
    log.info("Vigilando DUA: %s  (quietud: %ds, poll: %.1fs, pypdf: %s)",
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
            last_mod = max(p.stat().st_mtime for p in pdfs)
            if pending_since == 0.0:
                pending_since = now
                log.info("Detectados %d PDF(s) - esperando quietud de %ds...",
                         len(pdfs), quiet)

            cur_sizes = sizes_snapshot(pdfs)
            sizes_stable = (cur_sizes == prev_sizes)
            prev_sizes = cur_sizes
            quiet_ok = (now - last_mod) >= quiet

            if quiet_ok and sizes_stable:
                incompletos = [p for p in pdfs if not is_pdf_complete(p)]
                if incompletos:
                    for p in incompletos:
                        log.warning("PDF aun incompleto, espero: %s", p.name)
                    time.sleep(poll)
                    continue
                # Procesar uno a uno (cada DUA es un documento independiente)
                for p in pdfs:
                    try:
                        process_pdf(p, dirs)
                    except Exception as e:
                        log.error("Error procesando %s: %s", p.name, e, exc_info=True)
                        # mover a error/ para no reprocesar en bucle
                        try:
                            shutil.move(str(p), str(dirs["error"] / p.name))
                        except Exception:
                            pass
                pending_since = 0.0
                prev_sizes = {}

        except Exception as e:
            log.error("Error en el ciclo de vigilancia: %s", e, exc_info=True)

        time.sleep(poll)


def main():
    parser = argparse.ArgumentParser(description="DUA Inbox Watcher (OCR -> SharePoint)")
    base = os.getenv("DUA_BASE", "/data/ias_prod/data/dua")
    parser.add_argument("--inbox", default=os.getenv("DUA_INBOX", f"{base}/inbox"))
    parser.add_argument("--outbox", default=os.getenv("DUA_OUTBOX", f"{base}/outbox"))
    parser.add_argument("--review", default=os.getenv("DUA_REVIEW", f"{base}/review"))
    parser.add_argument("--processed", default=os.getenv("DUA_PROCESSED", f"{base}/processed"))
    parser.add_argument("--error", default=os.getenv("DUA_ERROR", f"{base}/error"))
    parser.add_argument("--history", default=os.getenv("DUA_HISTORY", f"{base}/historial.json"))
    parser.add_argument("--quiet", type=int, default=int(os.getenv("DUA_QUIET_SECONDS", "30")))
    parser.add_argument("--poll", type=float, default=float(os.getenv("DUA_POLL_SECONDS", "5")))
    args = parser.parse_args()

    inbox = Path(args.inbox)
    dirs = {
        "inbox": inbox,
        "outbox": Path(args.outbox),
        "review": Path(args.review),
        "processed": Path(args.processed),
        "error": Path(args.error),
        "history": Path(args.history),
    }
    for key in ("inbox", "outbox", "review", "processed", "error"):
        dirs[key].mkdir(parents=True, exist_ok=True)
    if not inbox.exists():
        log.error("El inbox no existe: %s", inbox)
        raise SystemExit(1)

    watch(inbox, dirs, args.quiet, args.poll)


if __name__ == "__main__":
    main()
