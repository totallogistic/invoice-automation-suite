"""
Entrega de EPIs - backend actualizado.
- Usa Entrega_de_EPIs_clean.docx como template base (con logo y cabecera oficial)
- Incluye campos: codigo, realizacion, resp_prevencion
- Embebe firma PNG en el documento
- Convierte a PDF con LibreOffice si se solicita

Reemplaza el bloque save_entrega_epis en app.py.
El template debe estar en: templates/template-entrega-epis.docx
La imagen del logo ya está embebida en el template.
"""

import os, shutil, subprocess, base64, tempfile
from datetime import datetime
from pathlib import Path
from fastapi import HTTPException, Request
from fastapi.responses import JSONResponse
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from docx import Document as _DocxDocument
from docx.shared import Pt, Cm, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Constants ──────────────────────────────────────────────────────────────────
TEXTO_LEGAL_EPIS = (
    "De acuerdo con el art. 29.2 de la Ley 31/1995 de 8 de Noviembre, de Prevención de Riesgos "
    "Laborales, así como del sistema de gestión de la prevención de riesgos, el trabajador:\n\n"
    "a ) Se compromete a utilizar correctamente los medios y equipos de protección facilitados.\n"
    "b ) Se responsabiliza de su mantenimiento y conservación.\n"
    "c ) Entiende que el equipo se le asigna de manera personal.\n"
    "d ) Devuelve el EPI usado o deteriorado antes de recibir otro nuevo, quedando prohibida la "
    "utilización de equipos deteriorados o caducos.\n"
    "e ) Colabora con la gestión documentada para la entrega de los equipos.\n"
    "f ) Comunicará al responsable del grupo o trabajador-enlace la pérdida, merma, deterioro o "
    "caducidad que pudiera sufrir el equipo de protección individual."
)


def _set_run(run, bold=False, size=10, italic=False):
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic


def _cell_write(cell, text, bold=False, size=10, center=False, italic=False):
    for child in list(cell.paragraphs[0]._p):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('r', 'sdt', 'hyperlink'):
            cell.paragraphs[0]._p.remove(child)
    run = cell.paragraphs[0].add_run(str(text) if text else '')
    _set_run(run, bold=bold, size=size, italic=italic)
    if center:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def _shade_cell(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def _add_image_to_cell(cell, img_path, width_cm=4.0):
    """Add an image to a table cell."""
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    run.add_picture(img_path, width=Cm(width_cm))


def _generate_epis_docx_v2(data: dict, output_dir: Path) -> tuple:
    """
    Generate FR-58 EPI delivery DOCX using clean template style.
    Returns (docx_path, docx_filename).
    """
    empleado       = data.get('empleado', '')
    puesto         = data.get('puesto', '')
    fecha          = data.get('fecha', '')
    codigo         = data.get('codigo', 'FR.58')
    realizacion    = data.get('realizacion', 'COORDINADOR')
    resp_prev      = data.get('resp_prevencion', 'Juan Antonio Martínez Lázaro')
    epis           = data.get('epis', [])
    firma_png_b64  = data.get('firma_png')   # base64 data URL or None

    # Format date nicely
    fecha_fmt = fecha
    try:
        from datetime import datetime as _dt
        d = _dt.strptime(fecha, '%Y-%m-%d')
        fecha_fmt = d.strftime('%d/%m/%Y')
        mes_anio  = d.strftime('%m/%Y')
    except Exception:
        mes_anio = fecha

    doc = _DocxDocument()

    # Page margins (matching clean template)
    for section in doc.sections:
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2)
        section.right_margin  = Cm(2)

    # ── LOGO HEADER TABLE ─────────────────────────────────────────────────────
    # 3-column header: [Logo | SISTEMA DE CALIDAD... | CÓDIGO/empty]
    LOGO_PATH = Path(__file__).parent / 'static' / 'logo_totallogistic.jpeg'
    tbl_width = 9500  # DXA ≈ page width with these margins

    hdr_table = doc.add_table(rows=2, cols=3)
    hdr_table.style = 'Table Grid'

    # Row 0: Logo | Title | Code info
    r0 = hdr_table.rows[0]
    r0.height = Pt(50)

    # Col 0: Logo
    c00 = r0.cells[0]
    c00.width = Cm(5)
    if LOGO_PATH.exists():
        _add_image_to_cell(c00, str(LOGO_PATH), width_cm=4.5)
    else:
        _cell_write(c00, 'TOTALLOGISTIC', bold=True, size=12)

    # Col 1: Sistema de calidad
    c01 = r0.cells[1]
    c01.width = Cm(8)
    _cell_write(c01, 'SISTEMA DE CALIDAD Y PREVENCIÓN DE RIESGOS', bold=True, size=12, center=True)

    # Col 2: Code box
    c02 = r0.cells[2]
    c02.width = Cm(4)
    p02 = c02.paragraphs[0]
    p02.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_code = p02.add_run(f'{codigo}')
    _set_run(r_code, bold=True, size=14)

    # Row 1: "EPIS por puesto" spanning all cols
    r1 = hdr_table.rows[1]
    r1.cells[0].merge(r1.cells[2])
    _cell_write(r1.cells[0], 'EPIS por puesto', bold=True, size=11)

    doc.add_paragraph()  # spacing

    # ── DATA TABLE (Puesto / Trabajador / Fecha / Realización) ────────────────
    data_table = doc.add_table(rows=3, cols=4)
    data_table.style = 'Table Grid'

    # Row 0: Puesto de Trabajo | value | DATOS DE LA EVALUACIÓN | blank
    _cell_write(data_table.rows[0].cells[0], 'Puesto de Trabajo:', bold=True, size=10)
    data_table.rows[0].cells[0].width = Cm(4)
    _cell_write(data_table.rows[0].cells[1], puesto.upper(), size=10)
    data_table.rows[0].cells[1].width = Cm(6)
    _cell_write(data_table.rows[0].cells[2], 'DATOS DE LA EVALUACIÓN DE RIESGOS', bold=True, size=9, italic=True)
    data_table.rows[0].cells[2].width = Cm(4)
    data_table.rows[0].cells[3].width = Cm(3)

    # Row 1: Trabajador | value | FECHA | date
    _cell_write(data_table.rows[1].cells[0], 'Trabajador:', bold=True, size=10)
    _cell_write(data_table.rows[1].cells[1], empleado.upper(), size=10)
    _cell_write(data_table.rows[1].cells[2], 'FECHA:', bold=True, size=10)
    _cell_write(data_table.rows[1].cells[3], mes_anio, size=10)

    # Row 2: Fecha | value | REALIZACIÓN | value
    _cell_write(data_table.rows[2].cells[0], f'Fecha: {fecha_fmt}', bold=True, size=10)
    data_table.rows[2].cells[0].merge(data_table.rows[2].cells[1])
    _cell_write(data_table.rows[2].cells[2], 'REALIZACIÓN:', bold=True, size=10)
    _cell_write(data_table.rows[2].cells[3], realizacion, size=10)

    doc.add_paragraph()

    # ── EPI TABLE ─────────────────────────────────────────────────────────────
    title_p = doc.add_paragraph()
    r = title_p.add_run("Listado de EPI´s")
    _set_run(r, bold=True, size=11)

    # Group by category
    categories = {}
    for epi in epis:
        cat = epi.get('categoria', '')
        categories.setdefault(cat, []).append(epi)

    total_rows = 1 + sum(1 + len(v) for v in categories.values())
    epi_table = doc.add_table(rows=total_rows, cols=4)
    epi_table.style = 'Table Grid'

    # Set column widths
    col_widths_cm = [1.8, 9.0, 1.0, 5.2]
    for row in epi_table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = Cm(col_widths_cm[i])

    # Header row
    _cell_write(epi_table.rows[0].cells[0], 'Normativa', bold=True, center=True, size=9)
    _cell_write(epi_table.rows[0].cells[1], 'Equipo de Protección Individual', bold=True, size=9)
    _cell_write(epi_table.rows[0].cells[2], 'SÍ', bold=True, center=True, size=9)
    _cell_write(epi_table.rows[0].cells[3], 'Observaciones', bold=True, size=9)
    for j in range(4):
        _shade_cell(epi_table.rows[0].cells[j], 'd9e1f2')

    row_idx = 1
    for cat_name, items in categories.items():
        # Category header
        epi_table.cell(row_idx, 0).merge(epi_table.cell(row_idx, 3))
        _cell_write(epi_table.rows[row_idx].cells[0], cat_name.upper(), bold=True, size=9)
        _shade_cell(epi_table.rows[row_idx].cells[0], 'f2f2f2')
        row_idx += 1

        for epi in items:
            _cell_write(epi_table.rows[row_idx].cells[0], epi.get('normativa', ''), size=8, center=True)
            _cell_write(epi_table.rows[row_idx].cells[1], epi.get('nombre', ''), size=9)
            _cell_write(epi_table.rows[row_idx].cells[2],
                        'X' if epi.get('entregado') else '', center=True, size=10, bold=True)
            _cell_write(epi_table.rows[row_idx].cells[3], epi.get('observaciones', ''), size=9)
            row_idx += 1

    doc.add_paragraph()

    # ── LEGAL TEXT ────────────────────────────────────────────────────────────
    for line in TEXTO_LEGAL_EPIS.split('\n'):
        if not line.strip():
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(line)
        _set_run(run, size=9)

    doc.add_paragraph()

    # ── SIGNATURE TABLE ───────────────────────────────────────────────────────
    sig_table = doc.add_table(rows=3, cols=2)
    sig_table.style = 'Table Grid'

    _cell_write(sig_table.rows[0].cells[0], 'Fdo: Responsable Prevención', bold=True, size=10)
    _cell_write(sig_table.rows[0].cells[1], 'Fdo: Trabajador', bold=True, size=10)

    # Signature image row
    sig_row = sig_table.rows[1]
    sig_row.height = Cm(4)
    # Left: static "firma" placeholder
    _cell_write(sig_row.cells[0], '', size=10)
    # Right: embed signature PNG if provided
    if firma_png_b64:
        try:
            # Strip data URL prefix
            if ',' in firma_png_b64:
                _, b64data = firma_png_b64.split(',', 1)
            else:
                b64data = firma_png_b64
            img_bytes = base64.b64decode(b64data)
            tmp_sig = Path(tempfile.mktemp(suffix='.png'))
            tmp_sig.write_bytes(img_bytes)
            _add_image_to_cell(sig_row.cells[1], str(tmp_sig), width_cm=6.0)
            tmp_sig.unlink(missing_ok=True)
        except Exception as e:
            print(f"[epis] Error embebiendo firma: {e}")

    # Names row
    _cell_write(sig_table.rows[2].cells[0], f'Nombre: {resp_prev}', size=10)
    _cell_write(sig_table.rows[2].cells[1], f'Nombre: {empleado.upper()}', size=10)

    # ── SAVE ──────────────────────────────────────────────────────────────────
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug = empleado.replace(' ', '_').upper()
    docx_filename = f"FR58_EPIs_{slug}_{timestamp}.docx"
    docx_path = output_dir / docx_filename
    doc.save(str(docx_path))

    return docx_path, docx_filename


def _docx_to_pdf(docx_path: Path, output_dir: Path) -> Path:
    """Convert DOCX to PDF using LibreOffice. Returns PDF path."""
    result = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', '--outdir', str(output_dir), str(docx_path)],
        capture_output=True, text=True, timeout=60
    )
    pdf_path = output_dir / (docx_path.stem + '.pdf')
    if not pdf_path.exists():
        raise RuntimeError(f"PDF conversion failed: {result.stderr}")
    return pdf_path


# ── Endpoint (replaces existing save_entrega_epis in app.py) ──────────────────
# Add this import at top of app.py:
#   from app_epis_v2 import save_entrega_epis_v2
# Add this route (replaces or alongside existing):
#   app.post("/api/save-entrega-epis")(save_entrega_epis_v2)

async def save_entrega_epis_v2(request: Request, generate_docx: bool = False, convert_pdf: bool = False):
    """
    Guarda entrega EPIs con template oficial, firma y opción PDF.
    """
    data     = await request.json()
    empleado = data.get('empleado', '').strip()
    puesto   = data.get('puesto',   '').strip()
    fecha    = data.get('fecha',    '').strip()
    epis     = data.get('epis',     [])

    if not empleado:
        raise HTTPException(400, "El campo 'empleado' es obligatorio")
    if not fecha:
        raise HTTPException(400, "El campo 'fecha' es obligatorio")

    entregados = [e for e in epis if e.get('entregado')]
    if not entregados:
        raise HTTPException(400, "Debe marcarse al menos un EPI como entregado")

    # Dirs
    BASE    = Path(__file__).parent
    EXCEL_DIR  = BASE / 'excel_storage'
    OUTPUT_DIR = BASE / 'output'
    EXCEL_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    # ── 1. Excel ──────────────────────────────────────────────────────────────
    excel_file = EXCEL_DIR / 'entrega-epis.xlsx'
    if excel_file.exists():
        wb = load_workbook(excel_file)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Entregas EPIs"
        cols = ["Fecha Registro", "Empleado", "Puesto", "Fecha Entrega",
                "Categoría", "Normativa", "EPI", "Observaciones"]
        ws.append(cols)
        hdr_fill = PatternFill(start_color="1a4d7e", end_color="1a4d7e", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for epi in entregados:
        ws.append([
            now_str, empleado, puesto, fecha,
            epi.get('categoria', ''),
            epi.get('normativa', ''),
            epi.get('nombre',    ''),
            epi.get('observaciones', ''),
        ])

    for col in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)
    wb.save(excel_file)

    response = {
        "success": True,
        "message": f"Entrega registrada: {len(entregados)} EPI(s) para {empleado}",
        "excel_file": excel_file.name,
    }

    # ── 2. DOCX + PDF — todo en excel_storage ─────────────────────────────────
    if generate_docx:
        try:
            # DOCX en excel_storage
            docx_path, docx_filename = _generate_epis_docx_v2(data, EXCEL_DIR)
            response["docx_filename"] = docx_filename

            if convert_pdf:
                # PDF en excel_storage
                pdf_path = _docx_to_pdf(docx_path, EXCEL_DIR)
                pdf_filename = pdf_path.name
                response["pdf_filename"] = pdf_filename
                response["download_url"] = f"/download-storage/{pdf_filename}"
            else:
                response["download_url"] = f"/download-storage/{docx_filename}"

        except Exception as e:
            print(f"[epis] Error generando documento: {e}")
            response["doc_error"] = str(e)

    return JSONResponse(response)