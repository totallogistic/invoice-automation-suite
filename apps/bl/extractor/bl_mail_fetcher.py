#!/usr/bin/env python3
"""
bl_mail_fetcher.py
==================
Conecta a Gmail via IMAP, busca emails con label BL en la ventana de tiempo
configurada, descarga los adjuntos PDF y los deja en el inbox del watcher.

Uso:
    python3 bl_mail_fetcher.py
    python3 bl_mail_fetcher.py --dry-run        # sin descargar, solo muestra qué haría
    python3 bl_mail_fetcher.py --since-hours 12 # ventana custom de últimas 12 horas

Variables de entorno (todas en .env.prod):
    GMAIL_USER_LEAR         tls-lear@totallogistic.es
    GMAIL_APP_PASSWORD_LEAR contraseña de aplicación de 16 caracteres
    GMAIL_LABEL             Label de Gmail a buscar (default: BL)
    BL_INBOX                Carpeta donde dejar los PDFs (default: /data/bl/inbox)
    BL_MAIL_WINDOW_START_HOUR  Hora inicio ventana día anterior (default: 22)
    BL_MAIL_WINDOW_END_HOUR    Hora fin ventana hoy (default: 8)
    BL_MAIL_WINDOW_END_MIN     Minuto fin ventana hoy (default: 59)
"""

import argparse
import email
import imaplib
import logging
import os
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("bl-mail-fetcher")

# ── Config desde entorno ──────────────────────────────────────────────────────
GMAIL_USER         = os.getenv("GMAIL_USER_LEAR", "")
GMAIL_PASSWORD     = os.getenv("GMAIL_APP_PASSWORD_LEAR", "")
GMAIL_LABEL        = os.getenv("GMAIL_LABEL", "BL")
BL_INBOX           = Path(os.getenv("BL_INBOX", "/data/bl/inbox"))
WINDOW_START_HOUR  = int(os.getenv("BL_MAIL_WINDOW_START_HOUR", "22"))
WINDOW_END_HOUR    = int(os.getenv("BL_MAIL_WINDOW_END_HOUR",   "8"))
WINDOW_END_MIN     = int(os.getenv("BL_MAIL_WINDOW_END_MIN",    "59"))

# ── Filtros de nombre de fichero ──────────────────────────────────────────────
PALABRAS_SI = ["bl", "bill", "lading", "conocimiento", "embarque", "con_emb"]
PALABRAS_NO = ["peticion", "request", "booking", "nota", "instrucciones", "draft",
               "bl_exportaciones", "bl_importaciones"]


def ventana_tiempo(since_hours: int | None = None) -> tuple[datetime, datetime]:
    """Calcula la ventana de tiempo para buscar emails."""
    ahora = datetime.now(timezone.utc)

    if since_hours is not None:
        desde = ahora - timedelta(hours=since_hours)
        hasta = ahora
    else:
        desde = ahora.replace(hour=WINDOW_START_HOUR, minute=0, second=0, microsecond=0)
        if ahora.hour < WINDOW_START_HOUR:
            desde -= timedelta(days=1)
        else:
            desde -= timedelta(days=1)

        hasta = ahora.replace(hour=WINDOW_END_HOUR, minute=WINDOW_END_MIN, second=59, microsecond=0)

    return desde, hasta


def filtrar_adjunto(nombre: str) -> bool:
    """Devuelve True si el adjunto debe descargarse."""
    nombre_low = nombre.lower()

    # Solo PDF
    if not nombre_low.endswith(".pdf"):
        return False

    # Debe contener alguna palabra clave
    tiene_si = any(p in nombre_low for p in PALABRAS_SI)
    if not tiene_si:
        log.info("  IGNORADO (sin palabras clave): %s", nombre)
        return False

    # No debe contener palabras excluidas
    tiene_no = any(p in nombre_low for p in PALABRAS_NO)
    if tiene_no:
        log.info("  IGNORADO (palabra excluida): %s", nombre)
        return False

    return True


def descargar_adjuntos(dry_run: bool = False, since_hours: int | None = None) -> int:
    """Conecta a Gmail, busca emails y descarga adjuntos. Devuelve nº de ficheros descargados."""

    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.error("GMAIL_USER_LEAR o GMAIL_APP_PASSWORD_LEAR no configurados")
        sys.exit(1)

    desde, hasta = ventana_tiempo(since_hours)
    log.info("Ventana: %s → %s", desde.strftime("%Y-%m-%d %H:%M UTC"), hasta.strftime("%Y-%m-%d %H:%M UTC"))
    log.info("Label: %s | Inbox: %s", GMAIL_LABEL, BL_INBOX)

    BL_INBOX.mkdir(parents=True, exist_ok=True)

    # Conectar a Gmail
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(GMAIL_USER, GMAIL_PASSWORD)
        log.info("Conectado como %s", GMAIL_USER)
    except Exception as e:
        log.error("Error conectando a Gmail: %s", e)
        sys.exit(1)

    # Seleccionar label
    # Gmail labels en IMAP usan formato especial para labels con espacios
    label_imap = GMAIL_LABEL if " " not in GMAIL_LABEL else f'"{GMAIL_LABEL}"'
    status, _ = mail.select(label_imap)
    if status != "OK":
        log.error("Label '%s' no encontrado en Gmail", GMAIL_LABEL)
        mail.logout()
        sys.exit(1)

    # Buscar emails por fecha (IMAP usa fecha sin hora, filtramos por hora en código)
    fecha_desde_str = desde.strftime("%d-%b-%Y")
    fecha_hasta_str = (hasta + timedelta(days=1)).strftime("%d-%b-%Y")
    search_criteria = f'(SINCE "{fecha_desde_str}" BEFORE "{fecha_hasta_str}")'

    status, message_ids = mail.search(None, search_criteria)
    if status != "OK":
        log.warning("Sin resultados para la búsqueda")
        mail.logout()
        return 0

    ids = message_ids[0].split()
    log.info("Emails encontrados en label %s: %d", GMAIL_LABEL, len(ids))

    descargados = 0
    ignorados   = 0

    for msg_id in ids:
        status, msg_data = mail.fetch(msg_id, "(RFC822)")
        if status != "OK":
            continue

        msg = email.message_from_bytes(msg_data[0][1])

        # Fecha del email
        fecha_str = msg.get("Date", "")
        try:
            from email.utils import parsedate_to_datetime
            fecha_email = parsedate_to_datetime(fecha_str)
            if fecha_email.tzinfo is None:
                fecha_email = fecha_email.replace(tzinfo=timezone.utc)
        except Exception:
            log.warning("No se pudo parsear la fecha: %s", fecha_str)
            continue

        # Filtro temporal
        if fecha_email < desde or fecha_email > hasta:
            log.debug("FUERA DE VENTANA: %s | %s", fecha_email.strftime("%Y-%m-%d %H:%M"), msg.get("Subject", ""))
            continue

        log.info("Email: %s | %s", fecha_email.strftime("%Y-%m-%d %H:%M UTC"), msg.get("Subject", "")[:60])

        # Procesar adjuntos
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            nombre = part.get_filename()
            if not nombre:
                continue

            # Decodificar nombre si está encoded
            from email.header import decode_header
            decoded = decode_header(nombre)
            nombre  = "".join(
                chunk.decode(enc or "utf-8") if isinstance(chunk, bytes) else chunk
                for chunk, enc in decoded
            )

            if not filtrar_adjunto(nombre):
                ignorados += 1
                continue

            # Resolver nombre de destino — si ya existe añadir sufijo _1, _2...
            destino = BL_INBOX / nombre
            if destino.exists():
                stem    = Path(nombre).stem
                suffix  = Path(nombre).suffix
                contador = 1
                while destino.exists():
                    destino = BL_INBOX / f"{stem}_{contador}{suffix}"
                    contador += 1
                log.info("  Renombrado a: %s", destino.name)

            if dry_run:
                log.info("  [DRY-RUN] Descargaría: %s", destino.name)
                descargados += 1
                continue

            # Descargar
            payload = part.get_payload(decode=True)
            if not payload:
                continue

            destino.write_bytes(payload)
            log.info("  DESCARGADO: %s (%.1f KB)", nombre, len(payload) / 1024)
            descargados += 1

    mail.logout()

    log.info("─────────────────────────────────────")
    log.info("Descargados: %d | Ignorados: %d", descargados, ignorados)

    return descargados


def main():
    parser = argparse.ArgumentParser(description="BL Mail Fetcher — descarga adjuntos BL de Gmail")
    parser.add_argument("--dry-run",     action="store_true", help="Solo muestra qué descargaría, sin guardar")
    parser.add_argument("--since-hours", type=int, default=None, help="Ventana custom: últimas N horas")
    args = parser.parse_args()

    descargados = descargar_adjuntos(dry_run=args.dry_run, since_hours=args.since_hours)

    if descargados == 0:
        log.info("Nada nuevo que descargar")
    elif args.dry_run:
        log.info("Dry-run completado — %d ficheros habrían sido descargados", descargados)
    else:
        log.info("Completado — %d PDF(s) en inbox", descargados)


if __name__ == "__main__":
    main()
