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


def ventana_tiempo(
    since_hours: int | None = None,
    from_hour: int | None = None,
    from_min: int = 0,
    to_hour: int | None = None,
    to_min: int = 59,
) -> tuple[datetime, datetime]:
    """Calcula la ventana de tiempo para buscar emails.

    Prioridad:
      1. since_hours  → últimas N horas desde ahora
      2. from_hour / to_hour → ventana explícita (hoy si to_hour > from_hour, si no ayer→hoy)
      3. Variables de entorno / defaults del módulo
    """
    ahora = datetime.now(timezone.utc)

    if since_hours is not None:
        return ahora - timedelta(hours=since_hours), ahora

    fh = from_hour if from_hour is not None else WINDOW_START_HOUR
    fm = from_min
    th = to_hour   if to_hour  is not None else WINDOW_END_HOUR
    tm = to_min    if to_hour  is not None else WINDOW_END_MIN

    hasta = ahora.replace(hour=th, minute=tm, second=59, microsecond=0)

    # Si from_hour > to_hour el inicio es el día anterior (ventana nocturna: ej 22→09)
    if fh > th:
        desde = (ahora - timedelta(days=1)).replace(hour=fh, minute=fm, second=0, microsecond=0)
    else:
        desde = ahora.replace(hour=fh, minute=fm, second=0, microsecond=0)

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


def descargar_adjuntos(
    dry_run: bool = False,
    since_hours: int | None = None,
    from_hour: int | None = None,
    from_min: int = 0,
    to_hour: int | None = None,
    to_min: int = 59,
) -> int:
    """Conecta a Gmail, busca emails y descarga adjuntos. Devuelve nº de ficheros descargados."""

    if not GMAIL_USER or not GMAIL_PASSWORD:
        log.error("GMAIL_USER_LEAR o GMAIL_APP_PASSWORD_LEAR no configurados")
        sys.exit(1)

    desde, hasta = ventana_tiempo(since_hours, from_hour, from_min, to_hour, to_min)
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

    from email.utils import parsedate_to_datetime
    from email.header import decode_header as _decode_header

    def _decode_str(s: str) -> str:
        decoded = _decode_header(s)
        return "".join(
            chunk.decode(enc or "utf-8") if isinstance(chunk, bytes) else chunk
            for chunk, enc in decoded
        )

    for msg_id in ids:
        try:
            # ── Paso 1: solo cabeceras (bytes, no MB) ────────────────────────────
            status, hdr_data = mail.fetch(msg_id, "(BODY[HEADER])")
            if status != "OK":
                continue

            hdr = email.message_from_bytes(hdr_data[0][1])

            # Filtro temporal — sin descargar adjuntos
            fecha_str = hdr.get("Date", "")
            try:
                fecha_email = parsedate_to_datetime(fecha_str)
                if fecha_email.tzinfo is None:
                    fecha_email = fecha_email.replace(tzinfo=timezone.utc)
            except Exception:
                log.warning("No se pudo parsear la fecha: %s", fecha_str)
                continue

            if fecha_email < desde or fecha_email > hasta:
                log.debug("FUERA DE VENTANA: %s | %s",
                          fecha_email.strftime("%Y-%m-%d %H:%M"), hdr.get("Subject", ""))
                continue

            log.info("Email: %s | %s",
                     fecha_email.strftime("%Y-%m-%d %H:%M UTC"), hdr.get("Subject", "")[:60])

            # ── Paso 2: estructura MIME sin cuerpo (solo metadata de partes) ─────
            status, struct_data = mail.fetch(msg_id, "(BODYSTRUCTURE)")
            if status != "OK":
                continue

            # ── Paso 3: mensaje completo solo si tiene adjuntos PDF candidatos ───
            # Para simplificar el parsing de BODYSTRUCTURE usamos RFC822 solo
            # para los emails que pasaron el filtro temporal — ya son pocos.
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            msg = email.message_from_bytes(msg_data[0][1])

        except imaplib.IMAP4.abort as e:
            if "OVERQUOTA" in str(e):
                log.error("Límite IMAP de Gmail alcanzado (OVERQUOTA) — guardados %d hasta ahora. "
                          "El límite se resetea en 1-24h.", descargados)
            else:
                log.error("Conexión IMAP abortada: %s — guardados %d hasta ahora.", e, descargados)
            break
        except Exception as e:
            log.warning("Error procesando email %s: %s — continuando", msg_id, e)
            continue

        # Procesar adjuntos
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue

            nombre_raw = part.get_filename()
            if not nombre_raw:
                continue

            nombre = _decode_str(nombre_raw)

            if not filtrar_adjunto(nombre):
                ignorados += 1
                continue

            # Resolver nombre de destino — si ya existe añadir sufijo _1, _2...
            destino  = BL_INBOX / nombre
            stem     = Path(nombre).stem
            suffix   = Path(nombre).suffix
            contador = 1
            while destino.exists():
                destino = BL_INBOX / f"{stem}_{contador}{suffix}"
                contador += 1
            if destino != BL_INBOX / nombre:
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
    parser.add_argument("--since-hours", type=int,  default=None, help="Ventana custom: últimas N horas")
    parser.add_argument("--from-hour",   type=int,  default=None, help="Hora inicio ventana (0-23)")
    parser.add_argument("--from-min",    type=int,  default=0,    help="Minuto inicio ventana (default: 0)")
    parser.add_argument("--to-hour",     type=int,  default=None, help="Hora fin ventana (0-23)")
    parser.add_argument("--to-min",      type=int,  default=59,   help="Minuto fin ventana (default: 59)")
    args = parser.parse_args()

    descargados = descargar_adjuntos(
        dry_run=args.dry_run,
        since_hours=args.since_hours,
        from_hour=args.from_hour,
        from_min=args.from_min,
        to_hour=args.to_hour,
        to_min=args.to_min,
    )

    if descargados == 0:
        log.info("Nada nuevo que descargar")
    elif args.dry_run:
        log.info("Dry-run completado — %d ficheros habrían sido descargados", descargados)
    else:
        log.info("Completado — %d PDF(s) en inbox", descargados)


if __name__ == "__main__":
    main()
