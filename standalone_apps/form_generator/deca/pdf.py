"""Generación del PDF nativo (DeCA / carta de porte / CMR) con QR incrustado.

- PDF nativo (desde datos, no escaneo), objetivo <= 5 MB.
- El QR codifica SOLO la URL pública del documento; los datos van en el cuerpo.
- Un mismo modelo de datos (models.DecaInput) genera los tres documentos según
  `tipo_documento`. Firma electrónica SIMPLE opcional (imagen del trazo + sello
  de tiempo propio + hash de integridad); NO es firma cualificada.
"""

from __future__ import annotations

import base64
import hashlib
import io
from xml.sax.saxutils import escape as _esc

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

_TITULOS = {
    "deca": ("Documento Electrónico de Control Administrativo",
             "DeCA · Transporte público de mercancías por carretera<br/>"
             "art. 6 Orden FOM/2861/2012 · Resolución 5-jun-2026 (BOE-A-2026-12784)"),
    "carta_porte": ("Carta de Porte",
                    "Contrato de transporte de mercancías por carretera (nacional)<br/>"
                    "Incluye los datos del documento de control (art. 6 Orden FOM/2861/2012)"),
    "cmr": ("Carta de Porte Internacional (e-CMR)",
            "Convenio relativo al contrato de transporte internacional de mercancías por carretera (CMR)<br/>"
            "Contiene los datos exigidos por el Convenio; formato electrónico"),
}


def _qr_image(url: str, box_mm: float = 34) -> Image:
    qr = segno.make(url, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=10, border=2)
    buf.seek(0)
    return Image(buf, width=box_mm * mm, height=box_mm * mm)


def qr_png_bytes(url: str, scale: int = 8) -> bytes:
    qr = segno.make(url, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=scale, border=2)
    return buf.getvalue()


def _styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle("DTitle", parent=ss["Title"], fontSize=15,
                          textColor=GREEN, spaceAfter=2, leading=18))
    ss.add(ParagraphStyle("DSub", parent=ss["Normal"], fontSize=7.5, textColor=GREY, leading=10))
    ss.add(ParagraphStyle("DLabel", parent=ss["Normal"], fontSize=7, textColor=GREY, leading=9))
    ss.add(ParagraphStyle("DValue", parent=ss["Normal"], fontSize=9.5, textColor=colors.black, leading=12))
    ss.add(ParagraphStyle("DSection", parent=ss["Normal"], fontSize=8.5, textColor=GREEN,
                          leading=11, spaceBefore=4, spaceAfter=2, fontName="Helvetica-Bold"))
    ss.add(ParagraphStyle("DUrl", parent=ss["Normal"], fontSize=6, textColor=GREY, leading=7, alignment=1))
    ss.add(ParagraphStyle("DFoot", parent=ss["Normal"], fontSize=6.5, textColor=GREY, leading=9))
    ss.add(ParagraphStyle("DSign", parent=ss["Normal"], fontSize=6.5, textColor=GREY, leading=8, alignment=1))
    return ss


def _field(ss, label: str, value) -> Table:
    value = "—" if value in (None, "") else _esc(str(value))
    t = Table([[Paragraph(_esc(label.upper()), ss["DLabel"])],
               [Paragraph(value, ss["DValue"])]])
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    return t


def _boxed(rows, colWidths):
    t = Table(rows, colWidths=colWidths)
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("BOX", (0, 0), (-1, -1), 0.5, LINE), ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE),
                           ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                           ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    return t


def _firma_image(b64: str, max_w_mm=42, max_h_mm=15) -> Image:
    from PIL import Image as PILImage
    if b64.startswith("data:"):
        b64 = b64.split(",", 1)[1]
    data = base64.b64decode(b64)
    iw, ih = PILImage.open(io.BytesIO(data)).size
    ar = (iw / ih) if ih else 3.0
    w, h = max_w_mm, max_w_mm / ar
    if h > max_h_mm:
        h, w = max_h_mm, max_h_mm * ar
    return Image(io.BytesIO(data), width=w * mm, height=h * mm)


def _firmas_block(ss, firmas):
    """Fila de cajas de firma (imagen del trazo o 'firmado' + rol/nombre/sello)."""
    cells = []
    for f in firmas:
        inner = []
        if f.firma_png:
            try:
                inner.append(_firma_image(f.firma_png))
            except Exception:
                inner.append(Paragraph("(firma)", ss["DSign"]))
        else:
            inner.append(Paragraph("✓ firmado", ss["DSign"]))
        sello = f.firmado_en.strftime("%d/%m/%Y %H:%M") if f.firmado_en else ""
        inner.append(Paragraph(f"<b>{_esc(f.rol.capitalize())}</b><br/>{_esc(f.nombre)}<br/>{_esc(sello)}", ss["DSign"]))
        cells.append(inner)
    while len(cells) < 3:
        cells.append([Paragraph("&nbsp;", ss["DSign"])])
    t = Table([cells[:3]], colWidths=[52.6 * mm, 52.6 * mm, 52.6 * mm], rowHeights=[24 * mm])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                           ("BOX", (0, 0), (-1, -1), 0.5, LINE), ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE)]))
    return t


def render_pdf(record: DecaRecord) -> bytes:
    d = record.datos
    tipo = d.tipo_documento if d.tipo_documento in _TITULOS else "deca"
    titulo, subtitulo = _TITULOS[tipo]
    ss = _styles()
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, topMargin=14 * mm, bottomMargin=12 * mm,
                            leftMargin=16 * mm, rightMargin=16 * mm, title=f"{titulo} {record.uuid}")
    story = []
    W = 158 * mm
    half = 79 * mm

    # ── Cabecera + QR ──
    title_block = [Paragraph(titulo, ss["DTitle"]), Paragraph(subtitulo, ss["DSub"])]
    qr_block = [_qr_image(record.url_publica or "https://pendiente.example/d/" + record.uuid),
                Paragraph("Escanee para descargar", ss["DUrl"])]
    header = Table([[title_block, qr_block]], colWidths=[120 * mm, 38 * mm])
    header.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story += [header, Spacer(1, 6),
              Table([[""]], colWidths=[W], style=TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1, GREEN)])),
              Spacer(1, 6)]

    # ── Partes ──
    remit_label = "Remitente" if tipo == "cmr" else "Cargador contractual"
    col_izq = [Paragraph(remit_label, ss["DSection"]),
               _field(ss, "Nombre / razón social", d.cargador_contractual.nombre),
               _field(ss, "NIF", d.cargador_contractual.nif),
               _field(ss, "Domicilio", d.cargador_contractual.domicilio)]
    col_der = [Paragraph("Transportista efectivo", ss["DSection"]),
               _field(ss, "Nombre / razón social", d.transportista_efectivo.nombre),
               _field(ss, "NIF", d.transportista_efectivo.nif)]
    story.append(Paragraph("PARTES", ss["DSection"]))
    story.append(_boxed([[col_izq, col_der]], [half, half]))
    story.append(Spacer(1, 6))

    # Destinatario (carta de porte / CMR)
    if tipo != "deca" and d.destinatario:
        story.append(_boxed([[
            [Paragraph("Destinatario", ss["DSection"]),
             _field(ss, "Nombre / razón social", d.destinatario.nombre),
             _field(ss, "NIF", d.destinatario.nif)],
            [_field(ss, "Domicilio", d.destinatario.domicilio), Spacer(1, 1)],
        ]], [half, half]))
        story.append(Spacer(1, 6))

    # ── Transporte + mercancía ──
    story.append(Paragraph("DATOS DEL TRANSPORTE", ss["DSection"]))
    mats = d.matricula_tractora + (f"  /  {d.matricula_remolque}" if d.matricula_remolque else "")
    filas = [
        [_field(ss, "Lugar de carga" if tipo == "cmr" else "Origen", d.lugar_carga or d.origen),
         _field(ss, "Lugar de entrega" if tipo == "cmr" else "Destino", d.lugar_entrega or d.destino)],
        [_field(ss, "Fecha del transporte", d.fecha_transporte.isoformat()),
         _field(ss, "Matrículas (tractora / remolque)", mats)],
        [_field(ss, "Naturaleza de la mercancía", d.mercancia.naturaleza),
         _field(ss, "Peso (kg)", f"{d.mercancia.peso_kg:,.0f}")],
        [_field(ss, "Bultos y marcas", d.mercancia.bultos),
         _field(ss, "Embalaje", d.mercancia.embalaje)],
        [_field(ss, "Autorizaciones especiales", d.autorizaciones_especiales), ""],
    ]
    story.append(_boxed(filas, [half, half]))
    story.append(Spacer(1, 6))

    # ── Condiciones (carta de porte / CMR) ──
    if tipo != "deca":
        story.append(Paragraph("CONDICIONES DEL TRANSPORTE", ss["DSection"]))
        cond = [
            [_field(ss, "Portes", d.portes), _field(ss, "Condiciones de pago", d.condiciones_pago)],
            [_field(ss, "Valor declarado", d.valor_declarado),
             _field(ss, "Instrucciones de aduana", d.instrucciones_aduana)],
        ]
        if tipo == "cmr":
            trasbordo = None if d.prohibicion_trasbordo is None else ("Sí" if d.prohibicion_trasbordo else "No")
            cond += [
                [_field(ss, "Reembolso", d.reembolso), _field(ss, "Instrucciones de seguro", d.instrucciones_seguro)],
                [_field(ss, "Prohibición de trasbordo", trasbordo), _field(ss, "Plazo de entrega", d.plazo_entrega)],
                [_field(ss, "Documentos anexos", d.documentos_anexos),
                 _field(ss, "Transportistas sucesivos", d.transportistas_sucesivos)],
            ]
        story.append(_boxed(cond, [half, half]))
        story.append(Spacer(1, 6))

    # ── Observaciones / reservas ──
    story.append(Paragraph("OBSERVACIONES Y RESERVAS", ss["DSection"]))
    story.append(_boxed([[Paragraph(_esc(d.observaciones or "—"), ss["DValue"])]], [W]))
    story.append(Spacer(1, 8))

    # ── Firmas (simple, opcional) ──
    if d.firmas:
        story.append(Paragraph("FIRMAS", ss["DSection"]))
        story.append(_firmas_block(ss, d.firmas))
        story.append(Spacer(1, 6))

    # ── Pie: identidad, sello e integridad ──
    creado = record.creado_en.isoformat(timespec="seconds") if record.creado_en else "—"
    integridad = hashlib.sha256(d.model_dump_json().encode("utf-8")).hexdigest()[:16]
    nota_firma = ("<b>Firma electrónica simple (no cualificada)</b> — imagen del trazo + sello de tiempo. "
                  f"Integridad SHA-256: {integridad}.<br/>" if d.firmas else "")
    if tipo == "deca":
        nota_legal = ("La firma no es obligatoria para la validez del DeCA (Resolución 5-jun-2026, Apartado Cuarto). "
                      "El enlace de descarga permanece activo durante el servicio y hasta 7 días naturales tras su finalización.")
    else:
        nota_legal = ("Documento de contrato de transporte. Si se requiere validez contractual electrónica plena, "
                      "la firma debe ser electrónica avanzada (eIDAS); la firma incluida aquí es simple, no cualificada.")
    story += [Table([[""]], colWidths=[W], style=TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.5, LINE)])),
              Spacer(1, 3),
              Paragraph(f"Identificador del documento (UUID): <b>{record.uuid}</b><br/>"
                        f"Generado: {creado} · URL de descarga: {_esc(record.url_publica or '(pendiente de bucket)')}<br/>"
                        f"{nota_firma}{nota_legal}", ss["DFoot"])]

    doc.build(story)
    return buf.getvalue()
