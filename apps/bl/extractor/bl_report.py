#!/usr/bin/env python3
"""
bl_report.py
============
Lee el CSV del día, genera un PDF con los BLs procesados y lo envía por email.
Se llama desde el unified_processor al terminar un batch BL.

Uso directo:
    python3 bl_report.py
    python3 bl_report.py --date 2026-03-31
    python3 bl_report.py --csv /data/bl/csv/bl_20260331.csv
"""

import argparse
import csv
import glob
import logging
import os
import smtplib
import sys
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
import shutil


log = logging.getLogger("bl-report")
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

# ── Config desde entorno ──────────────────────────────────────────────────────
BL_CSV_DIR  = Path(os.getenv("BL_CSV_DIR",   "/data/bl/csv"))
SMTP_HOST   = os.getenv("SMTP_HOST",   "")
SMTP_PORT   = int(os.getenv("SMTP_PORT",   "587"))
SMTP_USER   = os.getenv("SMTP_USER",   "")
SMTP_PASS   = os.getenv("SMTP_PASS",   "")
MAIL_FROM   = os.getenv("MAIL_FROM",   SMTP_USER)
MAIL_TO_BL  = os.getenv("MAIL_TO_BL",  "")

# Colores de marca por naviera
NAV_STYLES = {
    "AML":      "background:#fde8ea;color:#C8102E;border:1px solid #f5c0c6;",
    "BALEARIA": "background:#e6eaf3;color:#0A2240;border:1px solid #b8c4d8;",
    "DFDS":     "background:#fff8e1;color:#002B5C;border:1px solid #F0AB00;",
    "RFS":      "background:#fde8e8;color:#CC0000;border:1px solid #f5b8b8;",
    "TRASME":   "background:#fef0e6;color:#C04000;border:1px solid #f5c99a;",
}


DESTINO_EXPORTACIONES = "ALGECIRAS"

def leer_csv(csv_path: Path, filtrar_fecha: str | None = None) -> list[dict]:
    """Lee el CSV y filtra opcionalmente por fecha de embarque (campo 'fecha')."""
    registros = []
    try:
        with open(csv_path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if filtrar_fecha is None or row.get("fecha", "") == filtrar_fecha:
                    registros.append(row)
    except Exception as e:
        log.error("Error leyendo CSV %s: %s", csv_path, e)
    registros.sort(key=lambda r: (r.get("fecha", ""), r.get("hora", "")), reverse=True)
    return registros


def csv_del_dia(fecha: str) -> Path | None:
    """Devuelve el CSV del día indicado (formato yyyy-mm-dd)."""
    tag  = fecha.replace("-", "")
    path = BL_CSV_DIR / f"bl_{tag}.csv"
    return path if path.exists() else None


def generar_html(registros: list[dict], fecha: str, titulo: str = "Conocimientos de Embarque — ALGECIRAS") -> str:
    total    = len(registros)
    navieras: dict[str, int] = {}
    for r in registros:
        nav = r.get("naviera", "—")
        navieras[nav] = navieras.get(nav, 0) + 1

    # Resumen por naviera
    resumen_pills = "".join(
        f'<span style="display:inline-block;margin:0 4px 4px 0;padding:3px 10px;'
        f'border-radius:10px;font-size:11px;font-weight:700;{NAV_STYLES.get(nav,"")}">'
        f'{nav}: {count}</span>'
        for nav, count in sorted(navieras.items())
    )

    # Filas de la tabla
    filas = ""
    for r in registros:
        nav   = r.get("naviera", "—")
        style = NAV_STYLES.get(nav, "")
        badge = f'<span style="display:inline-block;padding:2px 7px;border-radius:8px;font-size:11px;font-weight:700;{style}">{nav}</span>'
        filas += f"""<tr>
            <td>{r.get('fecha','—')}</td>
            <td>{r.get('hora','—')}</td>
            <td>{badge}</td>
            <td>{r.get('num_bl','—')}</td>
            <td>{r.get('nombre','—')}</td>
            <td>{r.get('buque','—')}</td>
            <td style="font-family:monospace;font-size:11px;">{r.get('matricula','—')}</td>
        </tr>"""

    ts = datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M UTC")

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<style>
  body {{ font-family: Arial, sans-serif; font-size: 12px; color: #1f2937; margin: 0; padding: 20px; }}
  .header {{ border-bottom: 3px solid #1a4d7e; padding-bottom: 12px; margin-bottom: 16px; }}
  .header h1 {{ color: #1a4d7e; font-size: 18px; margin: 0 0 4px; }}
  .header p  {{ color: #6b7280; font-size: 11px; margin: 0; }}
  .stats {{ margin-bottom: 16px; padding: 10px 14px; background: #f8fafc; border-radius: 6px; border-left: 4px solid #1a4d7e; }}
  .stats .total {{ font-size: 22px; font-weight: 700; color: #1a4d7e; }}
  .stats .label {{ font-size: 10px; color: #6b7280; text-transform: uppercase; letter-spacing: 0.4px; }}
  table {{ width: 100%; border-collapse: collapse; }}
  thead th {{ background: #1a4d7e; color: white; padding: 7px 10px; text-align: left; font-size: 10px; text-transform: uppercase; letter-spacing: 0.4px; }}
  tbody td {{ padding: 6px 10px; border-bottom: 1px solid #e5e7eb; font-size: 11px; }}
  tbody tr:nth-child(even) {{ background: #f9fafb; }}
  .footer {{ margin-top: 14px; font-size: 10px; color: #9ca3af; text-align: right; }}
</style>
</head>
<body>
  <div class="header">
    <h1>{titulo}</h1>
    <p>Procesados: {fecha} &nbsp;·&nbsp; Generado: {ts}</p>
  </div>
  <div class="stats">
    <div class="total">{total}</div>
    <div class="label">BLs procesados hoy</div>
    <div style="margin-top:8px;">{resumen_pills}</div>
  </div>
  <table>
    <thead>
      <tr>
        <th>Fecha</th><th>Hora</th><th>Naviera</th><th>Nº BL</th>
        <th>Embarcador</th><th>Buque</th><th>Matrícula</th>
      </tr>
    </thead>
    <tbody>{filas}</tbody>
  </table>
  <div class="footer">totallogistic · BL Report · {ts}</div>
</body>
</html>"""


def generar_pdf(html: str, output: Path) -> bool:
    try:
        from weasyprint import HTML  # type: ignore
        HTML(string=html).write_pdf(str(output))
        log.info("PDF generado: %s", output)
        return True
    except ImportError:
        log.error("weasyprint no instalado — ejecuta: pip install weasyprint --break-system-packages")
        return False
    except Exception as e:
        log.error("Error generando PDF: %s", e)
        return False


def enviar_email(pdfs: list[Path], registros_exp: list[dict], registros_imp: list[dict], fecha: str) -> bool:
    if not SMTP_HOST or not SMTP_USER or not SMTP_PASS:
        log.warning("SMTP no configurado — email no enviado")
        return False
    if not MAIL_TO_BL:
        log.warning("MAIL_TO_BL no configurado — email no enviado")
        return False

    total_exp = len(registros_exp)
    total_imp = len(registros_imp)

    def resumen_nav(registros):
        navieras = {}
        for r in registros:
            nav = r.get("naviera", "—")
            navieras[nav] = navieras.get(nav, 0) + 1
        return "  ·  ".join(f"{nav}: {n}" for nav, n in sorted(navieras.items()))

    msg = MIMEMultipart()
    msg["From"]    = MAIL_FROM
    msg["To"]      = MAIL_TO_BL
    msg["Subject"] = f"[BL] {total_exp} Exportaciones · {total_imp} Importaciones — {fecha}"

    body = f"""<html><body style="font-family:Arial,sans-serif;color:#1f2937;">
<p><strong>Exportaciones (destino ALGECIRAS): {total_exp} BLs</strong></p>
<p style="color:#6b7280;font-size:13px;">{resumen_nav(registros_exp) or '—'}</p>
<p style="margin-top:1rem;"><strong>Importaciones (otros destinos): {total_imp} BLs</strong></p>
<p style="color:#6b7280;font-size:13px;">{resumen_nav(registros_imp) or '—'}</p>
<p style="margin-top:1rem;">Adjuntos los informes completos en PDF.</p>
<hr style="border:none;border-top:1px solid #e5e7eb;margin:16px 0;">
<p style="font-size:11px;color:#9ca3af;">totallogistic · BL Report · {fecha}</p>
</body></html>"""

    msg.attach(MIMEText(body, "html"))

    for pdf_path in pdfs:
        if pdf_path.exists():
            with open(pdf_path, "rb") as f:
                part = MIMEBase("application", "pdf")
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f'attachment; filename="{pdf_path.name}"')
                msg.attach(part)

    try:
        if SMTP_PORT == 465:
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as s:
                s.login(SMTP_USER, SMTP_PASS)
                s.send_message(msg)
        else:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
                s.starttls()
                s.login(SMTP_USER, SMTP_PASS)
                s.send_message(msg)
        log.info("Email enviado a %s", MAIL_TO_BL)
        return True
    except Exception as e:
        log.error("Error enviando email: %s", e)
        return False


def main():
    parser = argparse.ArgumentParser(description="BL Report — genera PDF y envía email")
    parser.add_argument("--date", default=None, help="Fecha yyyy-mm-dd (default: hoy)")
    parser.add_argument("--csv",  default=None, help="Ruta explícita al CSV")
    args = parser.parse_args()

    fecha = args.date or datetime.now(timezone.utc).strftime("%Y-%m-%d")

    if args.csv:
        csv_path = Path(args.csv)
    else:
        csv_path = csv_del_dia(fecha)

    if not csv_path or not csv_path.exists():
        log.warning("No hay CSV para la fecha %s — nada que reportar", fecha)
        sys.exit(0)

    # Sin filtro de fecha de embarque: se reportan todos los BLs
    # que se procesaron hoy, independientemente de su fecha de embarque.
    registros = leer_csv(csv_path)
    if not registros:
        log.warning("CSV vacío — nada que reportar")
        sys.exit(0)

    # Dividir en exportaciones (ALGECIRAS) e importaciones (resto)
    registros_exp = [r for r in registros if DESTINO_EXPORTACIONES in (r.get("puerto_destino") or "").upper()]
    registros_imp = [r for r in registros if DESTINO_EXPORTACIONES not in (r.get("puerto_destino") or "").upper()]

    log.info("Exportaciones (ALGECIRAS): %d · Importaciones (otros): %d", len(registros_exp), len(registros_imp))

    tag      = fecha.replace("-", "")
    pdfs_generados: list[Path] = []

    # Reporte Exportaciones
    if registros_exp:
        html_exp  = generar_html(registros_exp, fecha, titulo="Exportaciones — ALGECIRAS")
        pdf_exp   = BL_CSV_DIR / f"bl_exportaciones_{tag}.pdf"
        if generar_pdf(html_exp, pdf_exp):
            pdfs_generados.append(pdf_exp)
    else:
        log.info("Sin exportaciones — PDF no generado")

    # Reporte Importaciones
    if registros_imp:
        html_imp  = generar_html(registros_imp, fecha, titulo="Importaciones — Otros destinos")
        pdf_imp   = BL_CSV_DIR / f"bl_importaciones_{tag}.pdf"
        if generar_pdf(html_imp, pdf_imp):
            pdfs_generados.append(pdf_imp)
    else:
        log.info("Sin importaciones — PDF no generado")

    if not pdfs_generados:
        sys.exit(1)

    # Copiar a carpeta de impresión si está configurada
    BL_PRINT_FOLDER = os.getenv("BL_PRINT_FOLDER", "")
    if BL_PRINT_FOLDER:
        dest = Path(BL_PRINT_FOLDER)
        if dest.exists():
            for pdf in pdfs_generados:
                try:
                    shutil.copyfile(pdf, dest / pdf.name)
                    log.info("PDF enviado a impresora: %s", dest / pdf.name)
                except Exception as e:
                    log.warning("Error copiando a impresora %s: %s — continuando", pdf.name, e)
        else:
            log.warning("Carpeta de impresión no accesible: %s", dest)

    enviar_email(pdfs_generados, registros_exp, registros_imp, fecha)


if __name__ == "__main__":
    main()