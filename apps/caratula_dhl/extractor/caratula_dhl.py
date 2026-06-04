#!/usr/bin/env python3
"""
caratula_dhl.py — Generador de Carátulas DHL · DDG51 Base Naval Rota
====================================================================
Total Logistic Services, S.L.

Uso (modo unified stack — file watcher / API):
    python3 caratula_dhl.py file1.pdf file2.pdf ... -o /output_dir

Uso (standalone — carpeta de lote):
    python3 caratula_dhl.py /ruta/al/lote
    python3 caratula_dhl.py /ruta/al/lote --output /salida
    python3 caratula_dhl.py /ruta/al/lote --dry-run

La plantilla debe estar en el mismo directorio que el script:
    caratula_template.pdf   ← modelo con los placeholders naranjas

Por cada Factura TLS (NNNNCADNN[NN]) detectada en el lote se genera un
Caratula_DHL_<claves>.pdf que une carátula + factura + cotización + albarán
+ seguro + certificado + invoice proveedor + packing list + labels + resto
en el orden de prioridad de CATEGORIES.
"""
from __future__ import annotations

SCRIPT_VERSION = "2026-05-27.v1"

SCRIPT_CHANGELOG = """
## 2026-05-27.v1

### Lógica general
Genera una carátula DHL DDG51 (programa Navantia · Base Naval Rota) por
cada Factura TLS detectada en un lote de PDFs, y la entrega fusionada con
toda la documentación asociada (factura, cotización, albarán, seguro,
certificado, invoice del proveedor, packing list, labels, doc. Navantia)
en el orden estándar de envío a DHL.

### Entradas
- Conjunto de PDFs en un lote (un lote = un embarque).
- Los PDFs se clasifican por nombre de fichero:
  · `NNNNCADNNNN.pdf`           → Factura TLS
  · `NNNN-NNNN.pdf`             → Cotización
  · `HELV*.pdf`                 → Seguro
  · `TLSNAV*.pdf`               → Albarán
  · `LABELS*.pdf` / `LABEL*.pdf`→ Labels DHL
  · `E-*.pdf`                   → Packing list proveedor
  · `INV-*.pdf`                 → Invoice proveedor
  · `Certificado_*.pdf`         → Certificado
  · `<n> NAVANTIA*.pdf`         → Doc. Navantia
  · `Caratula_DHL_*.pdf`        → Carátula previa (se ignora)

### Generación de carátula
- Plantilla PDF embebida (`caratula_template.pdf`): se eliminan las
  anotaciones FreeText y se hace overlay del nuevo valor (rect blanco que
  borra el placeholder + texto Helvetica 10.5 en negro) sobre los rows del
  formulario.
- Fallback ReportLab si la plantilla no está disponible.

### Extracción del PDF de factura
- Embarque nº, Tipo, Suministrador, Pedido suministrador, Pedido transporte,
  Nº factura → regex sobre el texto extraído con pdfplumber.
- Tipo: `MARITIMO` / `AEREO` por el texto del PDF; `PAQUETERIA` cuando hay
  labels DHL en el lote.

### Salida
Un fichero por cada Factura TLS:
  · `Caratula_DHL_<keyA>_<keyB>.pdf` si la factura tiene 2 claves
  · `Caratula_DHL_<key>_<5_últimos_pedidoSum>.pdf` en otro caso
"""

import os
import sys
import re
import logging
import argparse
from pathlib import Path
from io import BytesIO

import pdfplumber
from pypdf import PdfReader, PdfWriter
from pypdf.generic import ArrayObject, NameObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import white, HexColor

logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

A4_W, A4_H = A4   # 595.28 × 841.89 pt
SCRIPT_DIR  = Path(__file__).parent
TEMPLATE_PDF = SCRIPT_DIR / "caratula_template.pdf"

# ══════════════════════════════════════════════════════════════════════════════
# FILL TEMPLATE — coordenadas extraídas del modelo original
# ══════════════════════════════════════════════════════════════════════════════

VALUE_X   = 188.80   # inicio X de todos los valores
FONT_SIZE = 10.5
DESCENDER = 2.1      # offset baseline Arial 10.5pt

# (pdfplumber_top, pdfplumber_bottom, data_key)
FIELD_ROWS = [
    (388.72, 399.22, "embarque"),
    (408.42, 418.92, "tipo"),
    (428.22, 438.72, "suministrador"),
    (446.02, 456.52, "pedidoSuministrador"),
    (463.82, 474.32, "pedidoTransporte"),
    (481.62, 492.12, "factura"),
]


def fill_template(data: dict, template_path: Path) -> bytes:
    """
    Rellena la plantilla PDF con los datos extraídos de la factura:
    1. Elimina anotaciones (los placeholders naranjas FreeText del modelo)
    2. Overlay: rectángulo blanco borra valor antiguo, texto nuevo en su lugar
    3. Devuelve bytes del PDF resultante
    """
    # Paso 1 · Quitar anotaciones (FreeText naranjas + Square amarillo)
    reader = PdfReader(template_path)
    clean_writer = PdfWriter()
    clean_writer.add_page(reader.pages[0])
    clean_writer.pages[0][NameObject("/Annots")] = ArrayObject()
    base_buf = BytesIO()
    clean_writer.write(base_buf)

    # Paso 2 · Overlay con los nuevos valores
    overlay_buf = BytesIO()
    cv = canvas.Canvas(overlay_buf, pagesize=A4)

    for top, bottom, key in FIELD_ROWS:
        value = str(data.get(key, '') or '')
        if not value:
            continue
        pdf_top    = A4_H - top
        pdf_bottom = A4_H - bottom
        baseline_y = pdf_bottom + DESCENDER

        # Rect blanco → borra valor anterior
        cv.setFillColor(white)
        cv.setStrokeColor(white)
        cv.rect(VALUE_X - 2, pdf_bottom - 2, 340, (pdf_top - pdf_bottom) + 4,
                fill=1, stroke=1)

        # Nuevo valor en negro, misma fuente y tamaño
        cv.setFillColor(HexColor("#000000"))
        cv.setFont("Helvetica", FONT_SIZE)
        cv.drawString(VALUE_X, baseline_y, value)

    cv.save()

    # Paso 3 · Fusionar overlay sobre la base limpia
    base_page    = PdfReader(BytesIO(base_buf.getvalue())).pages[0]
    overlay_page = PdfReader(BytesIO(overlay_buf.getvalue())).pages[0]
    base_page.merge_page(overlay_page)

    out = PdfWriter()
    out.add_page(base_page)
    out_buf = BytesIO()
    out.write(out_buf)
    return out_buf.getvalue()


def create_cover_page(data: dict) -> bytes:
    """Genera la carátula. Usa la plantilla si está disponible, sino ReportLab."""
    if TEMPLATE_PDF.exists():
        log.info(f"    ↳ Usando plantilla: {TEMPLATE_PDF.name}")
        return fill_template(data, TEMPLATE_PDF)
    else:
        log.warning(f"    ⚠ Plantilla no encontrada ({TEMPLATE_PDF}). Generando con ReportLab.")
        return _create_cover_reportlab(data)


# ══════════════════════════════════════════════════════════════════════════════
# FALLBACK — ReportLab (si no hay plantilla)
# ══════════════════════════════════════════════════════════════════════════════

C_NAVY=HexColor("#0B1F45"); C_NAVY_MID=HexColor("#162B5B"); C_ACCENT=HexColor("#F0A500")
C_BLUE=HexColor("#2870C8"); C_GRAY=HexColor("#617A9A"); C_GRAY_LT=HexColor("#8C9DB8")
C_ROW1=HexColor("#F2F4F8"); C_ROW2=HexColor("#F8F9FC"); C_BLACK=HexColor("#0D0D0D")

def _create_cover_reportlab(data: dict) -> bytes:
    buf = BytesIO()
    cv = canvas.Canvas(buf, pagesize=A4)
    cv.setFillColor(C_NAVY); cv.rect(0, A4_H-175, A4_W, 175, fill=1, stroke=0)
    cv.setFillColor(C_ACCENT); cv.rect(0, A4_H-178, A4_W, 3, fill=1, stroke=0)
    cv.setFillColor(C_GRAY_LT); cv.setFont("Helvetica",20)
    tw=cv.stringWidth("total","Helvetica",20); cv.drawString(36,A4_H-42,"total")
    cv.setFillColor(white); cv.setFont("Helvetica-Bold",20); cv.drawString(36+tw,A4_H-42,"logistic")
    cv.setFont("Helvetica-Bold",17); cv.drawRightString(A4_W-36,A4_H-42,"Navantia")
    cv.setStrokeColor(HexColor("#2E4A7A")); cv.setLineWidth(0.5); cv.line(36,A4_H-52,A4_W-36,A4_H-52)
    cv.setFillColor(C_GRAY_LT); cv.setFont("Helvetica",9); cv.drawCentredString(A4_W/2,A4_H-88,"PROGRAMA")
    cv.setFillColor(white); cv.setFont("Helvetica-Bold",25); cv.drawCentredString(A4_W/2,A4_H-117,"DDG51 Base Naval Rota")
    cv.setFillColor(C_GRAY_LT); cv.setFont("Helvetica",10); cv.drawCentredString(A4_W/2,A4_H-137,"LOGÍSTICA DEDICADA INTEGRAL")
    cv.setFillColor(C_BLACK); cv.setFont("Helvetica-Bold",20); cv.drawString(36,A4_H-220,"DOCUMENTACIÓN DE EMBARQUE")
    cv.setFillColor(C_BLUE); cv.rect(36,A4_H-233,A4_W-72,2.5,fill=1,stroke=0)
    fields=[("Embarque nº:",data.get("embarque","")),("Tipo de embarque:",data.get("tipo","PAQUETERIA")),
            ("Suministrador:",data.get("suministrador","")),("Pedido suministrador nº:",data.get("pedidoSuministrador","")),
            ("Pedido transporte nº:",data.get("pedidoTransporte","")),("N/ Factura nº:",data.get("factura",""))]
    ry=A4_H-260; rh=26
    for i,(label,val) in enumerate(fields):
        cv.setFillColor(C_ROW1 if i%2==0 else C_ROW2); cv.rect(36,ry-6,A4_W-72,rh,fill=1,stroke=0)
        if i==0: cv.setFillColor(C_ACCENT); cv.rect(36,ry-6,4,rh,fill=1,stroke=0)
        cv.setFillColor(C_GRAY); cv.setFont("Helvetica",10.5); cv.drawString(36+(10 if i==0 else 6),ry+4,label)
        cv.setFillColor(C_BLUE if i==0 else C_BLACK); cv.setFont("Helvetica-Bold",10.5); cv.drawString(235,ry+4,val or "")
        ry-=rh
    by=ry-28; cv.setFillColor(C_BLACK); cv.setFont("Helvetica",11); cv.drawString(36,by,"Estimados Sres.:")
    by-=22
    for line in ["Por la presente les rogamos encuentren a continuación la documentación","de embarque correspondiente al pedido de referencia, según se indica:"]:
        cv.drawString(36,by,line); by-=16
    by-=10
    for item in ["–   N/ Factura","–   N/ Oferta de servicio","–   Albarán de entrega","–   Certificado de seguro",
                 "–   Conocimiento de Embarque ( HBL – AWB )","–   Factura comercial del suministrador","–   Packing list del suministrador"]:
        cv.drawString(54,by,item); by-=17
    by-=14; cv.drawString(36,by,"Atentamente,"); by-=20; cv.setFont("Helvetica-Bold",11); cv.drawString(36,by,"Total Logistic Services, S.L.")
    cv.setFillColor(C_NAVY); cv.rect(0,0,A4_W,52,fill=1,stroke=0); cv.setFillColor(C_ACCENT); cv.rect(0,50,A4_W,1.5,fill=1,stroke=0)
    cv.setFillColor(white); cv.setFont("Helvetica-Bold",6.2)
    cv.drawString(36,38,"TOTAL LOGISTIC SERVICES, S.L.   ·   OPERADOR LOGÍSTICO   ·   TRANSITARIA INTERNACIONAL   ·   ADUANAS   ·   TRANSPORTES")
    cv.setFont("Helvetica",7); cv.setFillColor(C_GRAY_LT)
    cv.drawString(36,26,"Polígono Industrial La Menacha – Avenida del Estrecho Parcelas 5 a 15 – 11204 Algeciras, España")
    cv.drawString(36,14,"Tfno.: (34) 956604001  ·  Fax: (34) 956602403  ·  info@totallogistic.es  ·  www.totallogistic.es")
    ix=A4_W-36
    for tag in reversed(["ISO 9001","ISO 14001","ISO 45001"]):
        tw2=cv.stringWidth(tag,"Helvetica-Bold",5.5); ix-=tw2+14
        cv.setFillColor(C_NAVY_MID); cv.rect(ix,18,tw2+10,14,fill=1,stroke=0)
        cv.setFillColor(white); cv.setFont("Helvetica-Bold",5.5); cv.drawString(ix+5,22,tag)
    cv.save()
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# CATEGORÍAS
# ══════════════════════════════════════════════════════════════════════════════

# Prioridades de ensamblado por TIPO DE ENVÍO
# PAQUETERIA = envío courier DHL con labels (orden original del programa)
# TERRESTRE  = camión/transporte terrestre (sin cotización ni labels)
# MARITIMO   = envío marítimo
# AEREO      = envío aéreo
SHIPMENT_PRIORITIES = {
    "PAQUETERIA": {
        "factura": 2, "cotizacion": 3, "navantia_doc": 35,
        "seguro": 4, "certificado": 45, "albaran": 5, "labels": 6,
        "a7": 46, "hbl": 47, "hawb": 48,
        "invoice_supplier": 7, "packing_list": 8, "unknown": 90,
    },
    "TERRESTRE": {
        # Factura → HELV → TLSNAV → Invoice → Packing (sin cotización, sin labels)
        "factura": 1, "seguro": 2, "albaran": 3,
        "invoice_supplier": 4, "packing_list": 5,
        "cotizacion": 91, "labels": 92, "certificado": 93,
        "navantia_doc": 94, "a7": 95, "hbl": 99, "hawb": 99, "unknown": 90,
    },
    "MARITIMO": {
        # Factura → HELV → TLSNAV → A7 → HBL → Invoice → Packing
        "factura": 1, "seguro": 2, "albaran": 3, "a7": 4,
        "hbl": 5, "invoice_supplier": 6, "packing_list": 7,
        "cotizacion": 8, "certificado": 9, "navantia_doc": 10, "labels": 11,
        "hawb": 99, "unknown": 90,
    },
    "AEREO": {
        # Factura → HELV → TLSNAV → A7 → HAWB → Invoice → Packing
        "factura": 1, "seguro": 2, "albaran": 3, "a7": 4,
        "hawb": 5, "invoice_supplier": 6, "packing_list": 7,
        "cotizacion": 8, "certificado": 9, "navantia_doc": 10, "labels": 11,
        "hbl": 99, "unknown": 90,
    },
}

# Texto a mostrar en el campo "Tipo de embarque" de la carátula
TIPO_DISPLAY = {
    "PAQUETERIA": "PAQUETERIA",
    "TERRESTRE":  "TERRESTRE",
    "MARITIMO":   "MARÍTIMO",
    "AEREO":      "AÉREO",
}


CATEGORIES = {
    "factura":          ("Factura TLS",               2,   False),
    "cotizacion":       ("Cotización",                3,   True),
    "navantia_doc":     ("Doc. Navantia",             35,  True),
    "seguro":           ("Seguro / HELV",             4,   True),
    "certificado":      ("Certificado",               45,  True),
    "albaran":          ("Albarán",                   5,   True),
    "labels":           ("Labels",                    6,   True),
    "a7":               ("Doc. A7",                   46,  True),
    "hbl":              ("HBL (Marítimo)",             47,  True),
    "hawb":             ("HAWB (Aéreo)",               48,  True),
    "invoice_supplier": ("Invoice Proveedor",         7,   False),
    "packing_list":     ("Packing List",              8,   False),
    "caratula_exist":   ("Carátula (existente)",      99,  False),
    "unknown":          ("Sin clasificar",            90,  True),
}
def cat_label(k):    return CATEGORIES.get(k,("?",99,True))[0]
def cat_priority(k): return CATEGORIES.get(k,("?",99,True))[1]

def categorize(path: Path) -> str:
    stem = path.stem.strip()
    if re.match(r"^Caratula[\s_]DHL",          stem, re.I): return "caratula_exist"
    # Factura TLS: acepta prefijos como "FACT " antes del número
    # Ej: "1741CAD24 - 2994.pdf"  /  "FACT 1188CAD26 - PO TOTAL ...pdf"
    if re.search(r"\b\d{4}CAD\d{2,4}\b",       stem, re.I): return "factura"
    if re.match(r"^HELV\b",                     stem, re.I): return "seguro"
    if re.match(r"^TLSNAV",                     stem, re.I): return "albaran"
    if re.match(r"^LABELS?[\s\d]",              stem, re.I): return "labels"
    if re.match(r"^E-[A-Z0-9]",                 stem, re.I): return "packing_list"
    if re.match(r"^INV-",                        stem, re.I): return "invoice_supplier"
    if re.match(r"^Certificado[_\s]",            stem, re.I): return "certificado"
    if re.match(r"^\d+\s+NAVANTIA",             stem, re.I): return "navantia_doc"
    # Documentos marítimos / aéreos
    if re.match(r"^A7[\s_\-]",                  stem, re.I): return "a7"
    if re.match(r"^HBL[\s_\-]",                 stem, re.I): return "hbl"
    if re.match(r"^HAWB[\s_\-]",                stem, re.I): return "hawb"
    if re.match(r"^\d+\s*[-\u2013]\s*\d+",   stem):       return "cotizacion"
    return "unknown"


def extract_factura_keys(path: Path) -> list[str]:
    """
    Extrae claves de embarque de 4 dígitos del nombre de la factura.
    Ejemplos:
      "1741CAD24 - 2994.pdf"                         -> ['2994']
      "1916CAD23 - 2022 2156.pdf"                    -> ['2022', '2156']
      "FACT 1188CAD26 - PO TOTAL 7000110945 ....pdf" -> []  (sin clave en nombre)
    """
    parts = re.split(r"\s*[-\u2013]\s*", path.stem, maxsplit=1)
    if len(parts) > 1:
        keys = re.findall(r"\b\d{4}\b", parts[1])
        if keys:
            return keys
    return []


def specific_to(path: Path, key: str, all_keys: set) -> bool:
    matched = [k for k in all_keys if re.search(r"\b"+re.escape(k)+r"\b", path.stem)]
    return matched == [key]


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACCIÓN DE DATOS — 100% LOCAL
# ══════════════════════════════════════════════════════════════════════════════

def extract_invoice_data(pdf_path: Path, folder_has_labels: bool = False) -> dict:
    log.info(f"    ↳ Leyendo PDF: {pdf_path.name}")
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    result = {}
    m = re.search(r'\b(\d{4}CAD\d{2,4})\b', text, re.I)
    result['factura'] = m.group(1) if m else ''
    m = re.search(r'\b[A-Z]{2,3}\s*[-–]\s*\d{3,}\s*/\s*(\d{4,})', text)
    result['embarque'] = m.group(1) if m else ''
    m = re.search(r'Pedido\b[^\n]*\n\s*(\d{7,})', text, re.I)
    result['pedidoTransporte'] = m.group(1) if m else ''
    m = re.search(r'Fabricante\b[^\n]*\n\s*(\d{7,})', text, re.I)
    result['pedidoSuministrador'] = m.group(1) if m else ''
    m = re.search(r'^([A-Z][A-Z &,\.]+?)\s+NAVANTIA\b', text, re.MULTILINE)
    result['suministrador'] = m.group(1).strip().split()[0] if m else ''
    if folder_has_labels:          result['tipo'] = 'PAQUETERIA'
    elif re.search(r'\bMARITIM[AO]\b', text, re.I): result['tipo'] = 'MARITIMO'
    elif re.search(r'\bAERE[AO]\b',    text, re.I): result['tipo'] = 'AEREO'
    else:                          result['tipo'] = 'PAQUETERIA'
    return result


# ══════════════════════════════════════════════════════════════════════════════
# ENSAMBLADO
# ══════════════════════════════════════════════════════════════════════════════

def merge_pdfs(parts: list[tuple[int, bytes, str]]) -> bytes:
    writer = PdfWriter()
    for _, pdf_bytes, name in parts:
        try:
            for page in PdfReader(BytesIO(pdf_bytes)).pages:
                writer.add_page(page)
        except Exception as e:
            log.warning(f"    ⚠  No se pudo incluir '{name}': {e}")
    out = BytesIO(); writer.write(out); return out.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# RECOPILACIÓN Y PROCESO
# ══════════════════════════════════════════════════════════════════════════════

def collect_pdfs(paths: list[Path]) -> list[Path]:
    """
    Acepta una mezcla de ficheros y/o carpetas y devuelve la lista de PDFs
    encontrados, sin duplicados (case-insensitive sobre el nombre).
    """
    seen: set[str] = set()
    pdf_files: list[Path] = []

    def _add(p: Path):
        if p.suffix.lower() == ".pdf" and p.name.lower() not in seen:
            seen.add(p.name.lower())
            pdf_files.append(p)

    for path in paths:
        if path.is_dir():
            for p in sorted(path.iterdir()):
                if p.is_file():
                    _add(p)
        elif path.is_file():
            _add(path)
        else:
            log.warning(f"  ⚠  Ignorado (no es fichero ni carpeta): {path}")

    return pdf_files


def process_pdfs(pdf_files: list[Path], output_dir: Path):
    """
    Procesa una lista de PDFs (un lote) y genera las carátulas correspondientes.
    """
    if not pdf_files:
        log.error("No se encontraron archivos PDF.")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    log.info("\n┌─ Clasificación " + "─"*50)
    categorized: dict[str, list[Path]] = {}
    for p in pdf_files:
        cat = categorize(p)
        categorized.setdefault(cat, []).append(p)
        log.info(f"│  [{cat_label(cat):25s}]  {p.name}")
    log.info("└" + "─"*66)

    facturas = categorized.get("factura", [])
    if not facturas:
        log.error("No se encontró ninguna Factura TLS.")
        sys.exit(1)

    has_labels = bool(categorized.get("labels"))

    # Tipo de envío → prioridades de ensamblado
    stype_file = folder / "_SHIPMENT_TYPE.txt"
    shipment_type = stype_file.read_text(encoding="utf-8").strip().upper() if stype_file.exists() else "TERRESTRE"
    priorities = SHIPMENT_PRIORITIES.get(shipment_type, SHIPMENT_PRIORITIES["TERRESTRE"])

    def get_priority(cat: str) -> int:
        return priorities.get(cat, CATEGORIES.get(cat, ("?", 90, True))[1])

    # Texto que aparecerá en la carátula (con tildes donde corresponde)
    tipo_display = TIPO_DISPLAY.get(shipment_type, shipment_type)
    log.info(f"  Tipo envío : {shipment_type} → carátula: {tipo_display}")

    all_fac_keys = {f: extract_factura_keys(f) for f in facturas}
    all_keys_set = {k for keys in all_fac_keys.values() for k in keys}

    log.info(f"\n  Facturas : {len(facturas)}  |  Labels: {has_labels}  |  Claves: {sorted(all_keys_set)}")
    log.info(f"  Plantilla: {'✓ '+TEMPLATE_PDF.name if TEMPLATE_PDF.exists() else '✗ no encontrada (fallback ReportLab)'}")

    # Metadatos de email acumulados de todas las facturas del lote
    email_meta: dict = {"pedidos": [], "facturas": [], "suministradores": [], "shipment_type": shipment_type}

    generated: list[Path] = []

    for fac_path in sorted(facturas):
        log.info(f"\n{'═'*66}")
        log.info(f"  Factura: {fac_path.name}")
        try:
            data = extract_invoice_data(fac_path, folder_has_labels=has_labels)
        except Exception as e:
            log.error(f"  ✗ Error leyendo PDF: {e}")
            continue

        for k, v in data.items():
            log.info(f"    {k:<25} = {v}")

        # Sobreescribir "tipo" con el valor del selector UI (con tilde si aplica)
        data["tipo"] = tipo_display

        # Acumular metadatos de email
        for field, meta_key in [("pedidoTransporte","pedidos"),("factura","facturas"),("suministrador","suministradores")]:
            val = data.get(field, "")
            if val and val not in email_meta[meta_key]:
                email_meta[meta_key].append(val)

        fac_keys = all_fac_keys[fac_path]
        fac_key  = fac_keys[-1] if fac_keys else ""

        # Nombre del PDF de salida: {factura}_{pedidoTransporte}_{suministrador}_{pedidoSuministrador}
        def _clean(v: str) -> str:
            """Normaliza valor para nombre de fichero."""
            import re as _re
            return _re.sub(r'[^A-Za-z0-9]', '', (v or "").strip().upper()) or "X"

        out_name = (
            f"{_clean(data.get('factura'))}"
            f"_{_clean(data.get('pedidoTransporte'))}"
            f"_{_clean(data.get('suministrador'))}"
            f"_{_clean(data.get('pedidoSuministrador'))}.pdf"
        )
        parts: list[tuple[int, bytes, str]] = []
        parts.append((1, create_cover_page(data), "CARÁTULA"))
        parts.append((2, fac_path.read_bytes(), fac_path.name))

        for cat in ("invoice_supplier", "packing_list"):
            for p in categorized.get(cat, []):
                if len(all_keys_set) <= 1 or (fac_key and specific_to(p, fac_key, all_keys_set)):
                    parts.append((get_priority(cat), p.read_bytes(), p.name))

        for cat in ("cotizacion", "navantia_doc", "seguro", "certificado",
                    "albaran", "labels", "a7", "hbl", "hawb", "unknown"):
            for p in categorized.get(cat, []):
                parts.append((get_priority(cat), p.read_bytes(), p.name))

        parts.sort(key=lambda x: x[0])
        log.info("    Orden:")
        for prio, _, name in parts:
            log.info(f"      [{prio:2d}] {name}")

        final = merge_pdfs(parts)
        out_path = output_dir / out_name
        out_path.write_bytes(final)
        log.info(f"  ✓ {out_path.name}  ({len(final)//1024} KB)")
        generated.append(out_path)

    # Escribir metadatos de email para que el processor los use en el asunto
    try:
        import json as _json
        meta_path = output_dir / "_email_meta.json"
        meta_path.write_text(_json.dumps(email_meta, ensure_ascii=False, indent=2), encoding="utf-8")
        log.info(f"  email_meta → {meta_path}")
    except Exception as e:
        log.warning(f"  No se pudo escribir _email_meta.json: {e}")

    log.info(f"\n{'═'*66}")
    log.info(f"  ✓ Completado → {output_dir}  ({len(generated)} fichero(s))")
    return generated


# ══════════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        prog="caratula_dhl",
        description="Generador de Carátulas DHL — DDG51 Base Naval Rota",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Ejemplos:

  # Unified stack (file watcher / API)
  caratula_dhl.py file1.pdf file2.pdf ... -o /salida

  # Standalone
  caratula_dhl.py ./lote_2994
  caratula_dhl.py ./lote_2994 --output ./salida
  caratula_dhl.py ./lote_2994 --dry-run

La plantilla PDF debe estar en: {TEMPLATE_PDF}
""",
    )
    parser.add_argument("paths", nargs="+",
                        help="PDFs y/o carpeta(s) con los PDFs del lote")
    parser.add_argument("-o", "--output", default=None,
                        help="Carpeta de salida (por defecto: misma carpeta del primer path)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Solo clasificar, sin generar")
    parser.add_argument("--version", action="version",
                        version=f"%(prog)s {SCRIPT_VERSION}")
    args = parser.parse_args()

    input_paths = [Path(p).resolve() for p in args.paths]

    # Validación: todos los paths existen
    missing = [p for p in input_paths if not p.exists()]
    if missing:
        for p in missing:
            log.error(f"No existe: {p}")
        sys.exit(1)

    # Determinar carpeta de salida por defecto:
    #   - Si --output → ese
    #   - Si solo se pasó una carpeta → esa carpeta
    #   - En modo unified stack (lista de ficheros) → padre del primer fichero
    if args.output:
        output_dir = Path(args.output).resolve()
    elif len(input_paths) == 1 and input_paths[0].is_dir():
        output_dir = input_paths[0]
    else:
        output_dir = input_paths[0].parent

    output_dir.mkdir(parents=True, exist_ok=True)
    log.info(f"  Entrada : {[str(p) for p in input_paths]}")
    log.info(f"  Salida  : {output_dir}")

    pdf_files = collect_pdfs(input_paths)

    if args.dry_run:
        log.info("\n[DRY RUN]\n")
        for p in pdf_files:
            cat = categorize(p)
            keys = extract_factura_keys(p) if cat == "factura" else []
            log.info(f"  [{cat_label(cat):25s}]  {p.name}  {keys or ''}")
        return

    process_pdfs(pdf_files, output_dir)


if __name__ == "__main__":
    main()
