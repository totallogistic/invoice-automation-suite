"""Generación del PDF nativo del DeCA con QR incrustado.

- PDF nativo (generado desde datos, no escaneo), objetivo <= 5 MB.
- El QR codifica la URL pública del documento (que empieza por https://).
- El QR contiene SOLO la URL; los datos van en el cuerpo del PDF.
"""

from __future__ import annotations

import io

import segno
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from .models import DecaRecord

GREEN = colors.HexColor("#0b3a2e")
GREY = colors.HexColor("#52606d")
LINE = colors.HexColor("#c9d2da")


def _qr_image(url: str, box_mm: float = 34) -> Image:
    """Genera el QR de la URL como imagen embebible en el PDF."""
    qr = segno.make(url, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=10, border=2)
    buf.seek(0)
    return Image(buf, width=box_mm * mm, height=box_mm * mm)


def qr_png_bytes(url: str, scale: int = 8) -> bytes:
    """QR de la URL como PNG suelto (para la pantalla de resultado del formulario:
    lo que se enseña/envía al conductor)."""
    qr = segno.make(url, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=scale, border=2)
    return buf.getvalue()


def _styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("DecaTitle", parent=ss["Title"], fontSize=15,
                          textColor=GREEN, spaceAfter=2, leading=18))
    ss.add(ParagraphStyle("DecaSub", parent=ss["Normal"], fontSize=7.5,
                          textColor=GREY, leading=10))
    ss.add(ParagraphStyle("DecaLabel", parent=ss["Normal"], fontSize=7,
                          textColor=GREY, leading=9))
    ss.add(ParagraphStyle("DecaValue", parent=ss["Normal"], fontSize=9.5,
                          textColor=colors.black, leading=12))
    ss.add(ParagraphStyle("DecaSection", parent=ss["Normal"], fontSize=8.5,
                          textColor=GREEN, leading=11, spaceBefore=4, spaceAfter=2,
                          fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("DecaUrl", parent=ss["Normal"], fontSize=6,
                          textColor=GREY, leading=7, alignment=1))
    ss.add(ParagraphStyle("DecaFoot", parent=ss["Normal"], fontSize=6.5,
                          textColor=GREY, leading=9))
    return ss


def _field(ss, label: str, value) -> Table:
    """Celda etiqueta+valor."""
    value = "—" if value in (None, "") else str(value)
    t = Table([[Paragraph(label.upper(), ss["DecaLabel"])],
               [Paragraph(value, ss["DecaValue"])]], colWidths=[85 * mm])
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 1),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    return t


def render_pdf(record: DecaRecord) -> bytes:
    """Devuelve los bytes del PDF del DeCA."""
    d = record.datos
    ss = _styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=14 * mm,
                            bottomMargin=12 * mm, leftMargin=16 * mm,
                            rightMargin=16 * mm, title=f"DeCA {record.uuid}")
    story = []

    # Cabecera: título a la izquierda, QR a la derecha
    title_block = [
        Paragraph("Documento Electrónico de Control Administrativo", ss["DecaTitle"]),
        Paragraph("DeCA · Transporte público de mercancías por carretera<br/>"
                  "art. 6 Orden FOM/2861/2012 · Resolución 5-jun-2026 (BOE-A-2026-12784)",
                  ss["DecaSub"]),
    ]
    qr_block = [_qr_image(record.url_publica or "https://pendiente.example/d/" + record.uuid),
                Paragraph("Escanee para descargar", ss["DecaUrl"])]
    header = Table([[title_block, qr_block]], colWidths=[120 * mm, 38 * mm])
    header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story.append(header)
    story.append(Spacer(1, 6))
    story.append(Table([[""]], colWidths=[158 * mm],
                       style=TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, GREEN)])))
    story.append(Spacer(1, 6))

    # Partes: cargador contractual / transportista efectivo (identificados por separado)
    story.append(Paragraph("PARTES", ss["DecaSection"]))
    partes = Table([[
        [Paragraph("Cargador contractual", ss["DecaSection"]),
         _field(ss, "Nombre / razón social", d.cargador_contractual.nombre),
         _field(ss, "NIF", d.cargador_contractual.nif),
         _field(ss, "Domicilio", d.cargador_contractual.domicilio)],
        [Paragraph("Transportista efectivo", ss["DecaSection"]),
         _field(ss, "Nombre / razón social", d.transportista_efectivo.nombre),
         _field(ss, "NIF", d.transportista_efectivo.nif)],
    ]], colWidths=[79 * mm, 79 * mm])
    partes.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("BOX", (0, 0), (-1, -1), 0.5, LINE),
                                ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE),
                                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                                ("TOPPADDING", (0, 0), (-1, -1), 5),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    story.append(partes)
    story.append(Spacer(1, 8))

    # Datos del transporte
    story.append(Paragraph("DATOS DEL TRANSPORTE", ss["DecaSection"]))
    matriculas = d.matricula_tractora + (f"  /  {d.matricula_remolque}" if d.matricula_remolque else "")
    transporte = Table([
        [_field(ss, "Origen", d.origen), _field(ss, "Destino", d.destino)],
        [_field(ss, "Fecha del transporte", d.fecha_transporte.isoformat()),
         _field(ss, "Matrículas (tractora / remolque)", matriculas)],
        [_field(ss, "Naturaleza de la mercancía", d.mercancia.naturaleza),
         _field(ss, "Peso (kg)", f"{d.mercancia.peso_kg:,.0f}")],
        [_field(ss, "Autorizaciones especiales", d.autorizaciones_especiales), ""],
    ], colWidths=[79 * mm, 79 * mm])
    transporte.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                    ("BOX", (0, 0), (-1, -1), 0.5, LINE),
                                    ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE),
                                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    story.append(transporte)
    story.append(Spacer(1, 8))

    # Observaciones
    story.append(Paragraph("OBSERVACIONES", ss["DecaSection"]))
    obs = Table([[Paragraph(d.observaciones or "—", ss["DecaValue"])]], colWidths=[158 * mm])
    obs.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, LINE),
                             ("LEFTPADDING", (0, 0), (-1, -1), 6),
                             ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                             ("TOPPADDING", (0, 0), (-1, -1), 6),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
    story.append(obs)
    story.append(Spacer(1, 10))

    # Pie: identidad y sello temporal
    creado = record.creado_en.isoformat(timespec="seconds") if record.creado_en else "—"
    story.append(Table([[""]], colWidths=[158 * mm],
                       style=TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.5, LINE)])))
    story.append(Spacer(1, 3))
    story.append(Paragraph(
        f"Identificador del documento (UUID): <b>{record.uuid}</b><br/>"
        f"Generado: {creado} · URL de descarga: {record.url_publica or '(pendiente de bucket)'}<br/>"
        "La firma no es obligatoria para la validez del DeCA (Resolución 5-jun-2026, Apartado Cuarto). "
        "El enlace de descarga permanece activo durante el servicio y hasta 7 días naturales tras su finalización.",
        ss["DecaFoot"]))

    doc.build(story)
    return buf.getvalue()
