#!/usr/bin/env python3
"""
Universal JSON Schema Form Generator - FastAPI Server
Generates HTML forms from JSON Schema and outputs JSON files.

Usage:
    python app.py
    
Then open: http://localhost:8200
"""

import json
import json as _json
from pathlib import Path
from typing import Any, Dict

import csv as _csv
import glob as _glob

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import FileResponse as _FileResponse
from jsonschema import Draft202012Validator, ValidationError, validate, Draft7Validator
from fastapi.responses import FileResponse as _FileResponse

import subprocess
import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime
import uvicorn

from docx import Document as _DocxDocument
from docx.shared import Pt as _Pt, Cm as _Cm, RGBColor as _RGBColor
from docx.oxml.ns import qn as _qn
from docx.oxml import OxmlElement as _OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH as _WD_ALIGN

# Configuration
SCHEMAS_DIR = Path(__file__).parent / "schemas"
OUTPUT_DIR = Path(__file__).parent / "output"
TEMPLATES_DIR = Path(__file__).parent / "templates"
STATIC_DIR = Path(__file__).parent / "static"

# Excel storage configuration
EXCEL_STORAGE_DIR = Path(__file__).parent / "excel_storage"
EXCEL_STORAGE_DIR.mkdir(exist_ok=True)

# Email configuration (read from env)
SMTP_HOST = os.getenv("SMTP_HOST", "")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASS = os.getenv("SMTP_PASS", "")
MAIL_FROM = os.getenv("MAIL_FROM", "")

# Ensure directories exist
OUTPUT_DIR.mkdir(exist_ok=True)

# Directorios donde el processor guarda los PDFs procesados
BL_PROCESSED_ROOTS = [
    Path(os.getenv("BL_DATA_ROOT", "/data/bl")) / "processed",
    Path(os.getenv("BL_DATA_ROOT", "/data/bl")) / "error",
]

# Initialize FastAPI
app = FastAPI(title="JSON Schema Form Generator", version="1.0")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)


def send_email_with_json(to_email: str, subject: str, schema_name: str, data: dict, json_path: str):
    """
    Send email with JSON attachment.
    Returns True if successful, False otherwise.
    """
    if not SMTP_HOST or not SMTP_USER or not SMTP_PASS:
        print("⚠️ Email not configured, skipping...")
        return False
    
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = MAIL_FROM or SMTP_USER
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Email body
        body = f"""
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 600px; margin: 0 auto; padding: 20px;">
        <h2 style="color: #2563a8; border-bottom: 2px solid #2563a8; padding-bottom: 10px;">
            {subject}
        </h2>
        
        <p>Se ha completado un nuevo formulario:</p>
        
        <div style="background: #f5f5f5; padding: 15px; border-radius: 8px; margin: 20px 0;">
            <h3 style="margin-top: 0; color: #1a4d7e;">📋 Datos del Formulario:</h3>
            <table style="width: 100%; border-collapse: collapse;">
"""
        
        # Add data rows
        for key, value in data.items():
            # Skip internal fields
            if key.startswith('_'):
                continue
            
            # Format key
            formatted_key = key.replace('_', ' ').replace('-', ' ').title()
            
            # Format value
            if isinstance(value, bool):
                formatted_value = "✓ Sí" if value else "✗ No"
            elif value is None or value == "":
                formatted_value = "-"
            else:
                formatted_value = str(value)
            
            body += f"""
                <tr>
                    <td style="padding: 8px; border-bottom: 1px solid #ddd; font-weight: 600;">
                        {formatted_key}:
                    </td>
                    <td style="padding: 8px; border-bottom: 1px solid #ddd;">
                        {formatted_value}
                    </td>
                </tr>
"""
        
        body += """
            </table>
        </div>
        
        <p style="color: #6b7280; font-size: 0.9rem; margin-top: 30px;">
            Este mensaje ha sido generado automáticamente por el sistema de formularios de Totallogistic.
        </p>
    </div>
</body>
</html>
"""
        
        msg.attach(MIMEText(body, 'html'))
        
        # Attach JSON file
        with open(json_path, 'rb') as f:
            attach = MIMEBase('application', 'json')
            attach.set_payload(f.read())
            encoders.encode_base64(attach)
            attach.add_header('Content-Disposition', f'attachment; filename="{Path(json_path).name}"')
            msg.attach(attach)
        
        # Send email - Handle SSL vs STARTTLS
        if SMTP_PORT == 465:
            # SSL directo para puerto 465
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as server:
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
        else:
            # STARTTLS para puerto 587
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
                server.starttls()
                server.login(SMTP_USER, SMTP_PASS)
                server.send_message(msg)
        
        print(f"✅ Email sent to {to_email}")
        return True
        
    except Exception as e:
        print(f"❌ Error sending email: {e}")
        return False


def load_schema(schema_name: str) -> Dict[str, Any]:
    """Load a JSON schema from the schemas directory."""
    schema_path = SCHEMAS_DIR / f"{schema_name}.json"
    if not schema_path.exists():
        raise HTTPException(404, f"Schema not found: {schema_name}")
    
    with open(schema_path, "r", encoding="utf-8") as f:
        return json.load(f)


def validate_data(schema: Dict[str, Any], data: Any) -> tuple[bool, list[str]]:
    """Validate data against schema. Returns (is_valid, errors)."""
    validator = Draft202012Validator(schema)
    errors = []
    
    for error in validator.iter_errors(data):
        # Build error path
        path = " → ".join(str(p) for p in error.path) if error.path else "root"
        errors.append(f"{path}: {error.message}")
    
    return len(errors) == 0, errors

def append_to_excel(schema_name: str, data: dict) -> str:
    """
    Append data to an Excel file, creating it if it doesn't exist.
    Returns the Excel file path.
    """
    excel_file = EXCEL_STORAGE_DIR / f"{schema_name}.xlsx"
    
    # Load or create workbook
    if excel_file.exists():
        wb = load_workbook(excel_file)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = schema_name
        
        # Create header row
        headers = list(data.keys())
        ws.append(headers)
        
        # Style header
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
    
    # Append data row
    row_data = [data.get(key, "") for key in [cell.value for cell in ws[1]]]
    ws.append(row_data)
    
    # Auto-adjust column widths
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    # Save
    wb.save(excel_file)
    
    return str(excel_file)

@app.get("/form-custom/estanterias", response_class=HTMLResponse)
async def estanterias_custom(request: Request):
    """Custom form for estanterías with multiple rows."""
    return templates.TemplateResponse("estanterias-custom.html", {"request": request})

@app.get("/politica-privacidad", response_class=HTMLResponse)
async def politica_privacidad(request: Request):
    """Display privacy policy page."""
    return templates.TemplateResponse("politica-privacidad.html", {"request": request})

@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Show list of available schemas."""
    schemas = []
    if SCHEMAS_DIR.exists():
        for schema_file in SCHEMAS_DIR.glob("*.json"):
            try:
                schema = load_schema(schema_file.stem)
                schemas.append({
                    "name": schema_file.stem,
                    "title": schema.get("title", schema_file.stem),
                    "description": schema.get("description", ""),
                    "category": schema.get("category", "📋 Sin Categoría")
                })
            except Exception as e:
                print(f"Error loading {schema_file}: {e}")
    
    return templates.TemplateResponse("index.html", {
        "request": request,
        "schemas": schemas
    })


@app.get("/form/{schema_name}", response_class=HTMLResponse)
async def show_form(request: Request, schema_name: str):
    """Display form for a specific schema."""
    schema = load_schema(schema_name)
    
    # Check if schema has custom template
    template_name = schema.get("custom_template", "form.html")
    
    return templates.TemplateResponse(template_name, {
        "request": request,
        "schema_name": schema_name,
        "schema": schema,
        "form": schema,  # Alias for template compatibility
        "schema_json": json.dumps(schema, ensure_ascii=False),
        "success": False  # Will be True after successful submission
    })


@app.post("/form/{schema_name}", response_class=HTMLResponse)
async def submit_form(request: Request, schema_name: str):
    """Handle form submission for custom templates."""
    schema = load_schema(schema_name)
    
    # Get form data
    form_data = await request.form()
    data = dict(form_data)
    
    # Convert checkbox values
    for key, value in data.items():
        if value == "on":  # HTML checkbox sends "on" when checked
            data[key] = True
    
    # Save to Excel
    try:
        excel_path = append_to_excel(schema_name, data)
        success = True
        error = None
    except Exception as e:
        success = False
        error = str(e)
        print(f"Error saving to Excel: {e}")
    
    # Send email if configured
    schema_key = schema_name.upper().replace("-", "_")
    email_to = os.getenv(f"MAIL_TO_{schema_key}", "").strip()
    if not email_to:
        email_to = os.getenv("MAIL_TO", "")
    if success and email_to and SMTP_HOST:
        # Generate JSON file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{schema_name}_{timestamp}.json"
        filepath = OUTPUT_DIR / filename
        
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        # Send email
        subject = f"Nuevo formulario: {schema.get('title', schema_name)}"
        send_email_with_json(
            to_email=email_to,
            subject=subject,
            schema_name=schema_name,
            data=data,
            json_path=str(filepath)
        )
    
    # Return form with success message
    template_name = schema.get("custom_template", "form.html")
    
    return templates.TemplateResponse(template_name, {
        "request": request,
        "schema_name": schema_name,
        "schema": schema,
        "form": schema,
        "schema_json": json.dumps(schema, ensure_ascii=False),
        "success": success,
        "error": error
    })


@app.get("/api/schema/{schema_name}")
async def get_schema(schema_name: str):
    """Get schema as JSON."""
    schema = load_schema(schema_name)
    return JSONResponse(schema)


@app.post("/api/validate/{schema_name}")
async def validate(schema_name: str, data: Dict[str, Any]):
    """Validate data against schema."""
    schema = load_schema(schema_name)
    is_valid, errors = validate_data(schema, data)
    
    return JSONResponse({
        "valid": is_valid,
        "errors": errors
    })


@app.post("/api/save/{schema_name}")
async def save_data(schema_name: str, data: Dict[str, Any]):
    """Save validated data to JSON file."""
    schema = load_schema(schema_name)
    
    # Validate
    is_valid, errors = validate_data(schema, data)
    if not is_valid:
        return JSONResponse({
            "success": False,
            "errors": errors
        }, status_code=400)
    
    # Generate filename
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{schema_name}_{timestamp}.json"
    filepath = OUTPUT_DIR / filename
    
    # Save
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    return JSONResponse({
        "success": True,
        "filename": filename,
        "download_url": f"/download/{filename}"
    })

@app.post("/api/save-to-excel/{schema_name}")
async def save_to_excel(
    schema_name: str,
    data: dict,
    send_email: bool = False,
    email_to: str = None
):
    """Save form data to Excel and optionally send email."""
    
    # Validate against schema
    schema_file = SCHEMAS_DIR / f"{schema_name}.json"
    if not schema_file.exists():
        raise HTTPException(status_code=404, detail=f"Schema '{schema_name}' not found")
    
    with open(schema_file, 'r', encoding='utf-8') as f:
        schema = json.load(f)
    
    validator = Draft7Validator(schema)
    errors = list(validator.iter_errors(data))

    if errors:
        error_messages = [f"{e.path}: {e.message}" for e in errors]
        raise HTTPException(status_code=400, detail=f"Validation errors: {', '.join(error_messages)}")

    # Save to Excel
    try:
        excel_path = append_to_excel(schema_name, data)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error saving to Excel: {str(e)}")
    
    response = {
        "success": True,
        "message": "Registro guardado correctamente",
        "excel_file": Path(excel_path).name,
        "row_number": None  # Could calculate this
    }
    
    # Send email if configured (per-form env var or explicit recipient)
    schema_key = schema_name.upper().replace("-", "_")
    recipient = email_to or os.getenv(f"MAIL_TO_{schema_key}", "").strip()
    if not recipient:
        recipient = os.getenv("MAIL_TO", "")
    if recipient and SMTP_HOST:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{schema_name}_{timestamp}.json"
        filepath = OUTPUT_DIR / filename
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        subject = f"Nuevo formulario: {schema.get('title', schema_name)}"
        email_sent = send_email_with_json(
            to_email=recipient,
            subject=subject,
            schema_name=schema_name,
            data=data,
            json_path=str(filepath)
        )
        response["email_sent"] = email_sent
    
    return response

@app.post("/api/save-and-email/{schema_name}")
async def save_and_email(
    schema_name: str,
    data: dict,
    email_to: str = None
):
    """Save form data as JSON and send via email."""
    
    # Validate against schema
    schema_file = SCHEMAS_DIR / f"{schema_name}.json"
    if not schema_file.exists():
        raise HTTPException(status_code=404, detail=f"Schema '{schema_name}' not found")
    
    with open(schema_file, 'r', encoding='utf-8') as f:
        schema = json.load(f)
    
    try:
        validator = Draft7Validator(schema)
        errors = list(validator.iter_errors(data))
        if errors:
            error_msg = errors[0].message
            raise HTTPException(status_code=400, detail=f"Validation error: {error_msg}")
    except Exception as e:
        if not isinstance(e, HTTPException):
            raise HTTPException(status_code=400, detail=f"Validation error: {str(e)}")
        raise
    # Generate filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{schema_name}_{timestamp}.json"
    filepath = OUTPUT_DIR / filename
    
    # Save JSON
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error saving file: {str(e)}")
    
    # Determine email recipient
    if not email_to:
        schema_key = schema_name.upper().replace("-", "_")
        email_to = os.getenv(f"MAIL_TO_{schema_key}", "").strip()
    if not email_to:
        email_to = os.getenv("MAIL_TO", "")
    
    # Send email
    subject = f"Nuevo formulario: {schema.get('title', schema_name)}"
    email_sent = send_email_with_json(
        to_email=email_to,
        subject=subject,
        schema_name=schema_name,
        data=data,
        json_path=str(filepath)
    )
    
    return {
        "success": True,
        "message": "Formulario enviado por email" if email_sent else "Formulario guardado (error en email)",
        "filename": filename,
        "email_sent": email_sent,
        "email_to": email_to if email_sent else None
    }

@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download a generated JSON file."""
    filepath = OUTPUT_DIR / filename
    if not filepath.exists():
        raise HTTPException(404, "File not found")
    
    return FileResponse(
        filepath,
        media_type="application/json",
        filename=filename
    )


@app.get("/health")
async def health():
    """Health check endpoint."""
    return {
        "status": "ok",
        "schemas_available": len(list(SCHEMAS_DIR.glob("*.json"))) if SCHEMAS_DIR.exists() else 0
    }

# ========================================
# MODIFICAR ENDPOINT EXISTENTE: /api/save-estanterias-completo
# Reemplazar el endpoint actual con este código
# ========================================

@app.post("/api/save-estanterias-completo")
async def save_estanterias_completo(data: dict, sede: str = "alg"):
    """
    Guarda una revisión completa de estanterías:
    1. Guarda cada fila en el Excel de log (registro general)
    2. Crea/actualiza una hoja con el formato bonito por nave+trimestre
    3. Envía UN SOLO EMAIL al final
    
    Parámetros:
    - data: datos del formulario
    - sede: 'alg' o 'mlg' (por defecto 'alg')
    """
    from datetime import datetime
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    # Extraer datos
    fecha_revision = data.get('fechaRevision')
    responsable = data.get('responsable')
    nave = data.get('nave')
    trimestre = data.get('trimestre')
    obs_generales = data.get('observacionesGenerales', '')
    filas = data.get('filas', [])
    
    if not filas:
        raise HTTPException(status_code=400, detail="No hay filas para guardar")
    
    # Archivo Excel específico por sede
    excel_file = EXCEL_STORAGE_DIR / f"revision-estanterias-{sede}.xlsx"
    
    # 1. Guardar en hoja de LOG (registro general de todas las revisiones)
    if excel_file.exists():
        wb = load_workbook(excel_file)
    else:
        wb = Workbook()
        wb.remove(wb.active)  # Eliminar hoja por defecto
    
    # Asegurar que existe la hoja LOG
    if "LOG" not in wb.sheetnames:
        ws_log = wb.create_sheet("LOG", 0)
        # Headers
        headers = ["Fecha", "Responsable", "Nave", "Fila", "Trimestre", 
                   "Colocación", "Accesibilidad", "Corrosión", "Anclajes", 
                   "Protecciones", "Observaciones"]
        ws_log.append(headers)
        
        # Estilo headers
        header_fill = PatternFill(start_color="1a4d7e", end_color="1a4d7e", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws_log[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
    else:
        ws_log = wb["LOG"]
    
    # Añadir filas al LOG
    for fila_data in filas:
        ws_log.append([
            fecha_revision,
            responsable,
            nave,
            fila_data['fila'],
            trimestre,
            fila_data['colocacion'],
            fila_data['accesibilidad'],
            fila_data['corrosion'],
            fila_data['anclajes'],
            fila_data['protecciones'],
            fila_data.get('observaciones', obs_generales)
        ])
    
    # 2. Crear/actualizar hoja con formato bonito
    sheet_name = f"{nave} {trimestre[:3]}".upper()  # Ej: "NAVE 1T25 1ER"
    
    # Nombre de delegación según sede
    delegacion = "ALGECIRAS" if sede == "alg" else "MÁLAGA"
    
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        
        # Título (con nombre de delegación dinámico)
        ws.merge_cells('B2:K2')
        ws['B2'] = f'REVISIÓN DE ESTANTERIAS EN DELEGACION DE {delegacion}'
        ws['B2'].font = Font(size=14, bold=True, color="1a4d7e")
        ws['B2'].alignment = Alignment(horizontal="center")
        
        # Subtítulo
        ws.merge_cells('B9:K9')
        ws['B9'] = f'ESTANTERIAS - {nave.upper()}, REVISIÓN TRIMESTRAL (mantenimiento mínimo por parte del usuario)'
        ws['B9'].font = Font(size=12, bold=True)
        
        # Responsable
        ws.merge_cells('B11:D11')
        ws['B11'] = 'Responsable de la comprobación:'
        ws.merge_cells('E11:K11')
        ws['E11'] = responsable
        ws['E11'].font = Font(bold=True)
        
        # Headers tabla
        headers_tabla = ['* ZONA ALMACEN', 'COLOCACION', 'Accesibilidad y señalización', 
                        'CORROSION PATAS', 'ANCLAJES AL SUELO', 'PROTECCIONES CONTRA GOLPES', 
                        'Observaciones', f'{trimestre} (fecha)']
        
        for idx, header in enumerate(headers_tabla, start=2):
            cell = ws.cell(row=13, column=idx)
            cell.value = header
            cell.font = Font(bold=True, size=9)
            cell.fill = PatternFill(start_color="d9e1f2", end_color="d9e1f2", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
    else:
        ws = wb[sheet_name]
    
    # Añadir datos de filas
    start_row = 14
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    for idx, fila_data in enumerate(filas):
        row_num = start_row + idx
        
        ws.cell(row=row_num, column=2, value=fila_data['fila'])  # FILA 1, etc
        ws.cell(row=row_num, column=3, value=fila_data['colocacion'])
        ws.cell(row=row_num, column=4, value=fila_data['accesibilidad'])
        ws.cell(row=row_num, column=5, value=fila_data['corrosion'])
        ws.cell(row=row_num, column=6, value=fila_data['anclajes'])
        ws.cell(row=row_num, column=7, value=fila_data['protecciones'])
        ws.cell(row=row_num, column=8, value=fila_data.get('observaciones', ''))
        ws.cell(row=row_num, column=9, value=datetime.strptime(fecha_revision, '%Y-%m-%d'))
        ws.cell(row=row_num, column=9).number_format = 'DD/MM/YYYY'
        
        # Aplicar bordes y centrado
        for col in range(2, 10):
            cell = ws.cell(row=row_num, column=col)
            cell.border = thin_border
            if col > 2:  # Centrar excepto la primera columna (nombre fila)
                cell.alignment = Alignment(horizontal="center")
    
    # Leyenda al final
    last_row = start_row + len(filas) + 2
    ws.merge_cells(f'B{last_row}:K{last_row}')
    ws.cell(row=last_row, column=2, value='* según lo indicado en plano de situación.')
    
    ws.merge_cells(f'B{last_row+1}:D{last_row+1}')
    ws.cell(row=last_row+1, column=2, value='√    Correcto')
    ws.merge_cells(f'E{last_row+1}:F{last_row+1}')
    ws.cell(row=last_row+1, column=5, value='x   Incorrecto')
    
    # Ajustar anchos de columna
    ws.column_dimensions['B'].width = 15
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 16
    ws.column_dimensions['G'].width = 20
    ws.column_dimensions['H'].width = 25
    ws.column_dimensions['I'].width = 18
    
    # Guardar Excel
    wb.save(excel_file)
    
    # 3. Enviar UN SOLO EMAIL
    # Primero intenta variable específica de sede, luego genérica
    email_var = f"MAIL_TO_REVISION_ESTANTERIAS_{sede.upper()}"
    email_to = os.getenv(email_var, "").strip()
    if not email_to:
        email_to = os.getenv("MAIL_TO_REVISION_ESTANTERIAS", "").strip()
    if not email_to:
        email_to = os.getenv("MAIL_TO", "")
    
    email_sent = False
    
    if email_to and SMTP_HOST:
        # Crear JSON temporal para adjuntar
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_filename = f"revision-estanterias-{sede}_{timestamp}.json"
        json_path = OUTPUT_DIR / json_filename
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        subject = f"[{sede.upper()}] Revisión Estanterías - {nave} - {trimestre} ({fecha_revision})"
        email_sent = send_email_with_json(
            to_email=email_to,
            subject=subject,
            schema_name=f"revision-estanterias-{sede}",
            data=data,
            json_path=str(json_path)
        )
    
    return {
        "success": True,
        "message": f"Revisión guardada: {len(filas)} filas",
        "excel_file": excel_file.name,
        "sheet_created": sheet_name,
        "email_sent": email_sent,
        "email_to": email_to if email_sent else None,
        "sede": sede.upper()
    }

# ========================================
# ENDPOINT ESPECÍFICO PARA MANTENIMIENTO MAQUINARIA MLG
# Añadir este código a app.py
# ========================================

@app.post("/api/save-maquinaria-mlg")
async def save_maquinaria_mlg(data: dict):
    """
    Guarda mantenimiento de maquinaria de Málaga:
    1. Guarda cada revisión en hoja LOG
    2. Actualiza hoja ANUAL con formato tabla por meses
    """
    from datetime import datetime
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    
    # Extraer datos
    mes = data.get('mes')
    año = data.get('año')
    fecha_revision = data.get('fecha_revision')
    responsable = data.get('responsable')
    
    # Mapeo de meses a columnas (columna 6 = ENERO, 7 = FEBRERO, etc.)
    mes_to_col = {
        'ENERO': 6, 'FEBRERO': 7, 'MARZO': 8, 'ABRIL': 9,
        'MAYO': 10, 'JUNIO': 11, 'JULIO': 12, 'AGOSTO': 13,
        'SEPTIEMBRE': 14, 'OCTUBRE': 15, 'NOVIEMBRE': 16, 'DICIEMBRE': 17
    }
    
    # Definir máquinas y sus filas en la hoja ANUAL
    maquinas_config = [
        {'nombre': 'TRASPALETA ELECTRICA', 'modelo': '7PML20/6-736980-2005', 'fila_baterias': 14, 'fila_engrase': 15, 'key': 'traspaleta1'},
        {'nombre': 'TRASPALETA ELECTRICA', 'modelo': '7PML20/6-723382-2004', 'fila_baterias': 16, 'fila_engrase': 17, 'key': 'traspaleta2'},
        {'nombre': 'CARRETILLA ELECTRICA', 'modelo': '7FBMF25', 'fila_baterias': 18, 'fila_engrase': 19, 'key': 'carretilla'},
        {'nombre': 'APILADORA STILL', 'modelo': 'FM14', 'fila_baterias': 20, 'fila_engrase': 21, 'key': 'apiladora'},
        {'nombre': 'FURGON CITROEN', 'modelo': '5516GDV', 'fila_engrase': 22, 'key': 'furgon'},
    ]
    
    excel_file = EXCEL_STORAGE_DIR / "mantenimiento-maquinas-mlg.xlsx"
    
    # 1. Cargar o crear workbook
    if excel_file.exists():
        wb = load_workbook(excel_file)
    else:
        wb = Workbook()
        wb.remove(wb.active)
    
    # 2. Asegurar hoja LOG
    if "LOG" not in wb.sheetnames:
        ws_log = wb.create_sheet("LOG", 0)
        headers = ["Fecha", "Mes", "Año", "Responsable", "Máquina", "Modelo", "Revisión", "Estado", "Observaciones"]
        ws_log.append(headers)
        
        # Estilo headers
        header_fill = PatternFill(start_color="1a4d7e", end_color="1a4d7e", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for cell in ws_log[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
    else:
        ws_log = wb["LOG"]
    
    # 3. Asegurar hoja ANUAL
    if "ANUAL" not in wb.sheetnames:
        ws_anual = wb.create_sheet("ANUAL")
        
        # Título
        ws_anual.merge_cells('B2:Q2')
        ws_anual['B2'] = f'HOJA DE MANTENIMIENTO DE MAQUINARIA DE ALMACEN DELEGACION DE MÁLAGA'
        ws_anual['B2'].font = Font(size=14, bold=True, color="1a4d7e")
        ws_anual['B2'].alignment = Alignment(horizontal="center")
        
        # Subtítulo
        ws_anual.merge_cells('B9:N9')
        ws_anual['B9'] = 'EQUIPO DE ALMACEN – MANTENIMIENTO PERIODICO'
        ws_anual['B9'].font = Font(size=12, bold=True)
        
        # Año
        ws_anual['P9'] = 'AÑO'
        ws_anual['P9'].font = Font(bold=True)
        ws_anual['Q9'] = año
        ws_anual['Q9'].font = Font(bold=True)
        
        # Responsable
        ws_anual['B11'] = 'Responsable del mantenimiento :'
        ws_anual['D11'] = responsable
        ws_anual['D11'].font = Font(bold=True)
        
        # Headers tabla
        headers = ['MAQUINA', 'MODELO', '', 'REVISIONES', 'ENERO', 'FEB', 'MARZO', 'ABRIL', 
                   'MAYO', 'JUN', 'JUL', 'AGO', 'SEPT', 'OCT', 'NOV', 'DIC']
        for idx, header in enumerate(headers, start=2):
            cell = ws_anual.cell(row=13, column=idx)
            cell.value = header
            cell.font = Font(bold=True, size=10)
            cell.fill = PatternFill(start_color="d9e1f2", end_color="d9e1f2", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = Border(
                left=Side(style='thin'), right=Side(style='thin'),
                top=Side(style='thin'), bottom=Side(style='thin')
            )
        
        # Crear filas de máquinas
        thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        
        for maquina in maquinas_config:
            if maquina['key'] == 'furgon':
                # Furgón solo tiene ENGRASE
                row = maquina['fila_engrase']
                ws_anual.cell(row=row, column=2, value=maquina['nombre']).border = thin_border
                ws_anual.cell(row=row, column=3, value=maquina['modelo']).border = thin_border
                ws_anual.cell(row=row, column=5, value='ENGRASE').border = thin_border
                for col in range(6, 18):
                    ws_anual.cell(row=row, column=col).border = thin_border
                    ws_anual.cell(row=row, column=col).alignment = Alignment(horizontal="center")
            else:
                # Otras máquinas tienen BATERIAS y ENGRASE
                # Fila BATERIAS
                row_bat = maquina['fila_baterias']
                ws_anual.cell(row=row_bat, column=2, value=maquina['nombre']).border = thin_border
                ws_anual.cell(row=row_bat, column=3, value=maquina['modelo']).border = thin_border
                ws_anual.cell(row=row_bat, column=5, value='NIVELES BATERIAS').border = thin_border
                for col in range(6, 18):
                    ws_anual.cell(row=row_bat, column=col).border = thin_border
                    ws_anual.cell(row=row_bat, column=col).alignment = Alignment(horizontal="center")
                
                # Fila ENGRASE
                row_eng = maquina['fila_engrase']
                ws_anual.cell(row=row_eng, column=5, value='ENGRASE').border = thin_border
                for col in range(6, 18):
                    ws_anual.cell(row=row_eng, column=col).border = thin_border
                    ws_anual.cell(row=row_eng, column=col).alignment = Alignment(horizontal="center")
        
        # Leyenda
        ws_anual['B24'] = '√    Correcto'
        ws_anual['B25'] = 'X       Necesita intervencion (rellenado de agua destilada y/o engrase)'
        
        # Ajustar anchos
        ws_anual.column_dimensions['B'].width = 25
        ws_anual.column_dimensions['C'].width = 25
        ws_anual.column_dimensions['D'].width = 3
        ws_anual.column_dimensions['E'].width = 18
        for col_letter in ['F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']:
            ws_anual.column_dimensions[col_letter].width = 10
    else:
        ws_anual = wb["ANUAL"]
    
    # 4. Procesar cada revisión
    revisiones_guardadas = 0
    
    for maquina in maquinas_config:
        key = maquina['key']
        
        # Revisar BATERIAS (si la máquina tiene)
        if 'fila_baterias' in maquina:
            campo_bat = f"{key}_baterias"
            campo_bat_obs = f"{key}_baterias_obs"
            
            if campo_bat in data:
                estado = data[campo_bat]
                obs = data.get(campo_bat_obs, '')
                
                # Añadir a LOG
                ws_log.append([
                    fecha_revision,
                    mes,
                    año,
                    responsable,
                    maquina['nombre'],
                    maquina['modelo'],
                    'NIVELES BATERIAS',
                    estado,
                    obs
                ])
                
                # Actualizar ANUAL
                if mes in mes_to_col:
                    fila = maquina['fila_baterias']
                    columna = mes_to_col[mes]
                    ws_anual.cell(row=fila, column=columna, value=estado)
                    ws_anual.cell(row=fila, column=columna).alignment = Alignment(horizontal="center")
                
                revisiones_guardadas += 1
        
        # Revisar ENGRASE
        campo_eng = f"{key}_engrase"
        campo_eng_obs = f"{key}_engrase_obs"
        
        if campo_eng in data:
            estado = data[campo_eng]
            obs = data.get(campo_eng_obs, '')
            
            # Añadir a LOG
            ws_log.append([
                fecha_revision,
                mes,
                año,
                responsable,
                maquina['nombre'],
                maquina['modelo'],
                'ENGRASE',
                estado,
                obs
            ])
            
            # Actualizar ANUAL
            if mes in mes_to_col:
                fila = maquina['fila_engrase']
                columna = mes_to_col[mes]
                ws_anual.cell(row=fila, column=columna, value=estado)
                ws_anual.cell(row=fila, column=columna).alignment = Alignment(horizontal="center")
            
            revisiones_guardadas += 1
    
    # 5. Guardar Excel
    wb.save(excel_file)
    
    # 6. Email (opcional)
    email_to = os.getenv("MAIL_TO_MANTENIMIENTO_MAQUINAS_MLG", "").strip()
    if not email_to:
        email_to = os.getenv("MAIL_TO", "")
    
    email_sent = False
    
    if email_to and SMTP_HOST:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        json_filename = f"mantenimiento-maquinas-mlg_{timestamp}.json"
        json_path = OUTPUT_DIR / json_filename
        
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        subject = f"[MLG] Mantenimiento Maquinaria - {mes} {año}"
        email_sent = send_email_with_json(
            to_email=email_to,
            subject=subject,
            schema_name="mantenimiento-maquinas-mlg",
            data=data,
            json_path=str(json_path)
        )
    
    return {
        "success": True,
        "message": f"Mantenimiento guardado: {revisiones_guardadas} revisiones",
        "excel_file": excel_file.name,
        "mes": mes,
        "año": año,
        "email_sent": email_sent
    }



# ── Rutas de datos ────────────────────────────────────────────────────────────
# BL_DATA_ROOT se define en .env.prod  ej: /data/ias_prod/data/bl
_BL_ROOT     = Path(os.getenv("BL_DATA_ROOT", "/data/bl"))
BL_CSV_DIR   = Path(os.getenv("BL_CSV_DIR",   str(_BL_ROOT / "csv")))
BL_ESTADO_FILE = BL_CSV_DIR / "bl_estado.json"
 
_BL_PDF_ROOTS = [_BL_ROOT / "processed", _BL_ROOT / "error"]
 
 
# ── Helpers ───────────────────────────────────────────────────────────────────
 
def _bl_leer_registros() -> list:
    registros = []
    vistos: set = set()
    for csv_path in sorted(_glob.glob(str(BL_CSV_DIR / "bl_*.csv")), reverse=True):
        try:
            with open(csv_path, newline="", encoding="utf-8") as f:
                for row in _csv.DictReader(f):
                    k = (row.get("archivo",""), row.get("naviera",""), row.get("num_bl",""))
                    if k in vistos:
                        continue
                    vistos.add(k)
                    row["_csv_file"] = Path(csv_path).name
                    registros.append(row)
        except Exception as e:
            print(f"[BL] Error leyendo {csv_path}: {e}")
    registros.sort(key=lambda r: (r.get("fecha",""), r.get("hora","")), reverse=True)
    return registros
 
 
def _bl_leer_estado() -> dict:
    try:
        if BL_ESTADO_FILE.exists():
            return _json.loads(BL_ESTADO_FILE.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[BL] Error leyendo estado: {e}")
    return {}
 
 
def _bl_guardar_estado(estado: dict) -> None:
    try:
        BL_CSV_DIR.mkdir(parents=True, exist_ok=True)
        BL_ESTADO_FILE.write_text(
            _json.dumps(estado, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        print(f"[BL] Error guardando estado: {e}")
        raise
 
 
# ── Endpoints ─────────────────────────────────────────────────────────────────
 
@app.get("/bl", response_class=HTMLResponse)
async def bl_viewer(request: Request):
    return templates.TemplateResponse("bl-viewer.html", {"request": request})
 
 
@app.post("/api/bl/sync")
async def bl_sync():
    """Fuerza la copia de BLs desde Google Drive al inbox local."""
    sync_script = Path(os.getenv("BL_SYNC_SCRIPT", "/opt/bl_sync.sh"))
    if not sync_script.exists():
        raise HTTPException(404, f"Script de sync no encontrado: {sync_script}")
    try:
        result = subprocess.run(
            ["bash", str(sync_script)],
            capture_output=True, text=True, timeout=120
        )
        if result.returncode == 0:
            return JSONResponse({"ok": True, "message": "Sync completado correctamente"})
        else:
            raise HTTPException(500, f"Error en sync: {result.stderr.strip() or result.stdout.strip()}")
    except subprocess.TimeoutExpired:
        raise HTTPException(504, "Sync tardó demasiado (>120s)")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, str(e))


@app.get("/api/bl/registros")
def bl_registros(
    naviera:        str = "",
    fecha_desde:    str = "",
    fecha_hasta:    str = "",
    q:              str = "",
    solo_hecho:     str = "",
    tipo:           str = "",   # "exp" = ALGECIRAS, "imp" = otros
    puerto_destino: str = "",
):
    registros = _bl_leer_registros()
    estado    = _bl_leer_estado()
 
    for r in registros:
        clave = f"{r.get('archivo','')}|{r.get('naviera','')}|{r.get('num_bl','')}|{r.get('fecha','')}"
        r["hecho"] = estado.get(clave, False)
 
    if naviera:
        registros = [r for r in registros if r.get("naviera","").upper() == naviera.upper()]
    if fecha_desde:
        registros = [r for r in registros if r.get("fecha","") >= fecha_desde]
    if fecha_hasta:
        registros = [r for r in registros if r.get("fecha","") <= fecha_hasta]
    if tipo == "exp":
        registros = [r for r in registros if "ALGECIRAS" in (r.get("puerto_destino") or "").upper()]
    elif tipo == "imp":
        registros = [r for r in registros if "ALGECIRAS" not in (r.get("puerto_destino") or "").upper()]
    if puerto_destino:
        pd_up = puerto_destino.strip().upper()
        registros = [r for r in registros if pd_up in (r.get("puerto_destino") or "").upper()]
    if solo_hecho == "1":
        registros = [r for r in registros if r.get("hecho")]
    elif solo_hecho == "0":
        registros = [r for r in registros if not r.get("hecho")]
    if q:
        q_low = q.lower()
        registros = [r for r in registros if any(q_low in str(v).lower() for v in r.values())]
 
    return JSONResponse({"total": len(registros), "registros": registros})
 
 
@app.get("/api/bl/stats")
def bl_stats():
    registros = _bl_leer_registros()
    estado    = _bl_leer_estado()
 
    navieras: dict = {}
    for r in registros:
        nav = r.get("naviera", "—")
        navieras[nav] = navieras.get(nav, 0) + 1
 
    fechas = [r.get("fecha","") for r in registros if r.get("fecha","")]
    csvs   = [Path(p).name for p in sorted(_glob.glob(str(BL_CSV_DIR / "bl_*.csv")), reverse=True)]
 
    hechos = sum(1 for r in registros
                 if estado.get(f"{r.get('archivo','')}|{r.get('naviera','')}|{r.get('num_bl','')}", False))
 
    return JSONResponse({
        "total_registros": len(registros),
        "hechos":          hechos,
        "pendientes":      len(registros) - hechos,
        "navieras":        navieras,
        "fecha_min":       min(fechas) if fechas else None,
        "fecha_max":       max(fechas) if fechas else None,
        "csvs":            csvs,
    })
 
 
@app.post("/api/bl/toggle")
async def bl_toggle(request: Request):
    try:
        body  = await request.json()
        clave = body.get("clave", "").strip()
        if not clave:
            raise HTTPException(400, "clave requerida")
        estado = _bl_leer_estado()
        nuevo  = not estado.get(clave, False)
        estado[clave] = nuevo
        _bl_guardar_estado(estado)
        return JSONResponse({"clave": clave, "hecho": nuevo})
    except HTTPException:
        raise
    except Exception as e:
        print(f"[BL] toggle error: {e}")
        raise HTTPException(500, f"Error al guardar estado: {e}")
 
 
@app.get("/api/bl/pdf/{filename}")
def bl_pdf_download(filename: str):
    safe = Path(filename).name
    if not safe.lower().endswith(".pdf"):
        raise HTTPException(400, "Solo ficheros PDF")
    for root in _BL_PDF_ROOTS:
        if not root.exists():
            continue
        matches = sorted(root.glob(f"*/{safe}"), reverse=True)
        if matches:
            return _FileResponse(str(matches[0]), media_type="application/pdf", filename=safe)
    raise HTTPException(404, f"PDF no encontrado: {safe}")


def _docx_set_cell(cell, text, bold=False, size=11, align=None):
    """Clear a table cell and write formatted text."""
    from docx.shared import Pt
    # Remove all existing paragraphs except the first
    while len(cell.paragraphs) > 1:
        p = cell.paragraphs[-1]._p
        p.getparent().remove(p)
    para = cell.paragraphs[0]
    # Clear existing runs / SDTs from the paragraph
    for child in list(para._p):
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if tag in ('r', 'sdt', 'hyperlink', 'ins', 'del', 'bookmarkStart', 'bookmarkEnd'):
            para._p.remove(child)
    run = para.add_run(str(text) if text is not None else '')
    run.font.name = 'Calibri'
    run.font.size = Pt(size)
    run.font.bold = bold
    if align:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        para.alignment = align
 
 
def _docx_shade_cell(cell, hex_color="1a4d7e"):
    """Apply background fill to a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = _OxmlElement('w:shd')
    shd.set(_qn('w:val'), 'clear')
    shd.set(_qn('w:color'), 'auto')
    shd.set(_qn('w:fill'), hex_color)
    tcPr.append(shd)
 
 
def _generate_liquidacion_docx(data: dict):
    """
    Generate a Liquidación de Gastos de Viaje DOCX from data dict.
    Returns (Path, filename) tuple.
    """
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt, Cm
 
    doc = _DocxDocument()
 
    # --- Page margins ---
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)
 
    def section_header(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.bold = True
        return p
 
    def two_col_table(rows_data):
        """rows_data: list of (label, value)"""
        t = doc.add_table(rows=len(rows_data), cols=2)
        t.style = 'Table Grid'
        for i, (label, value) in enumerate(rows_data):
            t.rows[i].cells[0].width = Cm(5)
            t.rows[i].cells[1].width = Cm(11)
            _docx_set_cell(t.rows[i].cells[0], label, bold=True)
            _docx_set_cell(t.rows[i].cells[1], value)
        return t
 
    # ── TITLE ──────────────────────────────────────────────
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run('LIQUIDACIÓN DE GASTOS DE VIAJE')
    title_run.font.name = 'Calibri'
    title_run.font.size = Pt(14)
    title_run.font.bold = True
 
    doc.add_paragraph()  # spacing
 
    # ── DATOS DEL EMPLEADO ─────────────────────────────────
    section_header('DATOS DEL EMPLEADO')
    two_col_table([
        ('Nombre y apellidos', data.get('empleado', '')),
        ('DNI',                data.get('dni', '')),
    ])
 
    # ── INFORMACIÓN DEL VIAJE ──────────────────────────────
    section_header('INFORMACIÓN DEL VIAJE')
    two_col_table([
        ('Día de salida',  data.get('dia_salida', '')),
        ('Día de regreso', data.get('dia_regreso', '')),
    ])
 
    # ── MOTIVO ─────────────────────────────────────────────
    section_header('MOTIVO DEL DESPLAZAMIENTO')
    t_motivo = doc.add_table(rows=1, cols=1)
    t_motivo.style = 'Table Grid'
    t_motivo.rows[0].cells[0].width = Cm(16)
    t_motivo.rows[0].height = Cm(1.2)
    _docx_set_cell(t_motivo.rows[0].cells[0], data.get('motivo', ''))
 
    # ── GASTOS ─────────────────────────────────────────────
    section_header('GASTOS DEL VIAJE')
 
    gastos = data.get('gastos', [])
 
    # Rows: 1 header + N data + 1 total
    n_rows = 1 + len(gastos) + 1
    t = doc.add_table(rows=n_rows, cols=5)
    t.style = 'Table Grid'
 
    # Column widths (total ~17 cm usable)
    col_widths = [Cm(2.2), Cm(4.0), Cm(5.5), Cm(3.0), Cm(2.3)]
    for row in t.rows:
        for i, cell in enumerate(row.cells):
            cell.width = col_widths[i]
 
    # Header row
    headers = ['FECHA', 'Nº FRA/TICKET', 'PROVEEDOR', 'CONCEPTO', 'IMPORTE']
    for j, h in enumerate(headers):
        _docx_set_cell(t.rows[0].cells[j], h, bold=True)
        _docx_shade_cell(t.rows[0].cells[j], 'd9e1f2')  # light blue
 
    # Data rows
    total = 0.0
    for i, gasto in enumerate(gastos):
        row_idx = i + 1
        try:
            importe_val = float(str(gasto.get('importe', 0)).replace(',', '.'))
        except (ValueError, TypeError):
            importe_val = 0.0
        total += importe_val
 
        _docx_set_cell(t.rows[row_idx].cells[0], gasto.get('fecha', ''))
        _docx_set_cell(t.rows[row_idx].cells[1], str(gasto.get('num_fra', '')))
        _docx_set_cell(t.rows[row_idx].cells[2], gasto.get('proveedor', ''))
        _docx_set_cell(t.rows[row_idx].cells[3], gasto.get('concepto', ''))
        _docx_set_cell(t.rows[row_idx].cells[4], f"{importe_val:.2f}")
 
    # Total row — merge first 3 cells (cols 0-2), label in col 3, value in col 4
    # After merge, cells[0]=cells[1]=cells[2]=merged master; cells[3]=col3; cells[4]=col4
    total_row_idx = n_rows - 1
    t.cell(total_row_idx, 0).merge(t.cell(total_row_idx, 2))
    _docx_set_cell(t.rows[total_row_idx].cells[0], '')                        # merged: empty
    _docx_set_cell(t.rows[total_row_idx].cells[3], 'TOTAL', bold=True)        # concepto col
    _docx_set_cell(t.rows[total_row_idx].cells[4], f"{total:.2f} €", bold=True)  # importe col
    _docx_shade_cell(t.rows[total_row_idx].cells[3], 'f2f2f2')
    _docx_shade_cell(t.rows[total_row_idx].cells[4], 'f2f2f2')
 
    # ── FIRMA ──────────────────────────────────────────────
    doc.add_paragraph()
    firma_p = doc.add_paragraph()
    firma_run = firma_p.add_run('Firma empleado')
    firma_run.font.name = 'Calibri'
    firma_run.font.size = Pt(10)
 
    # ── SAVE ───────────────────────────────────────────────
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    empleado_slug = data.get('empleado', 'empleado').replace(' ', '_').upper()
    filename = f"liquidacion_{empleado_slug}_{timestamp}.docx"
    filepath = OUTPUT_DIR / filename
    doc.save(str(filepath))
 
    return filepath, filename
 
 
@app.post("/api/save-liquidacion")
async def save_liquidacion(request: Request, generate_docx: bool = False):
    """
    Guarda una liquidación de gastos:
    1. Añade todas las líneas al Excel mensual
    2. Si generate_docx=true, genera el .docx y devuelve URL de descarga
 
    Body JSON esperado:
    {
      "empleado": "NOMBRE APELLIDOS",
      "dni": "12345678A",
      "dia_salida": "2026-04-01",
      "dia_regreso": "2026-04-03",
      "motivo": "Viaje a Madrid por reunión con cliente",
      "gastos": [
        {"fecha": "2026-04-01", "num_fra": "129", "proveedor": "RADIO TAXI", "concepto": "TAXI", "importe": 16.50},
        ...
      ]
    }
    """
    data = await request.json()
 
    empleado = data.get('empleado', '').strip()
    gastos   = data.get('gastos', [])
 
    if not empleado:
        raise HTTPException(400, "El campo 'empleado' es obligatorio")
    if not gastos:
        raise HTTPException(400, "Debe incluir al menos un gasto")
 
    # ── 1. Guardar en Excel ────────────────────────────────
    excel_file = EXCEL_STORAGE_DIR / "liquidaciones-gastos.xlsx"
 
    if excel_file.exists():
        wb = load_workbook(excel_file)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Liquidaciones"
        cols = [
            "Fecha Registro", "Empleado", "DNI",
            "Día Salida", "Día Regreso", "Motivo",
            "Fecha Gasto", "Nº Fra/Ticket", "Proveedor", "Concepto", "Importe (€)"
        ]
        ws.append(cols)
        hdr_fill = PatternFill(start_color="1a4d7e", end_color="1a4d7e", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF")
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center")
 
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for gasto in gastos:
        try:
            importe = float(str(gasto.get('importe', 0)).replace(',', '.'))
        except (ValueError, TypeError):
            importe = 0.0
        ws.append([
            now_str,
            empleado,
            data.get('dni', ''),
            data.get('dia_salida', ''),
            data.get('dia_regreso', ''),
            data.get('motivo', ''),
            gasto.get('fecha', ''),
            str(gasto.get('num_fra', '')),
            gasto.get('proveedor', ''),
            gasto.get('concepto', ''),
            importe,
        ])
 
    # Auto-width columns
    for col in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
 
    wb.save(excel_file)
 
    response = {
        "success": True,
        "message": f"Liquidación guardada: {len(gastos)} gasto(s)",
        "excel_file": excel_file.name,
    }
 
    # ── 2. Generar DOCX (opcional) ─────────────────────────
    if generate_docx:
        try:
            _, docx_filename = _generate_liquidacion_docx(data)
            response["docx_filename"] = docx_filename
            response["download_url"]  = f"/download/{docx_filename}"
        except Exception as e:
            print(f"[liquidacion] Error generando DOCX: {e}")
            response["docx_error"] = str(e)
 
    # ── 3. Email (si está configurado) ────────────────────
    email_to = os.getenv("MAIL_TO_LIQUIDACIONES", "").strip() or os.getenv("MAIL_TO", "")
    if email_to and SMTP_HOST and generate_docx and response.get("download_url"):
        docx_path = OUTPUT_DIR / response["docx_filename"]
        subject = f"Liquidación gastos: {empleado} ({data.get('dia_salida','')} – {data.get('dia_regreso','')})"
        send_email_with_json(
            to_email=email_to,
            subject=subject,
            schema_name="liquidacion-gastos",
            data=data,
            json_path=str(docx_path),
        )
 
    return JSONResponse(response)

# ============================================================
# CONTROL DE PRECINTOS
# Añadir este bloque a app.py (antes del if __name__ == "__main__")
# Añadir al .env:  PRECINTOS_PIN=TUPIN
# ============================================================

import threading as _threading
_precintos_lock = _threading.Lock()

PRECINTOS_FILE = EXCEL_STORAGE_DIR / "control-precintos.xlsx"

PRECINTOS_HEADERS = [
    "Nº Precinto", "Matrícula", "Expediente",
    "Estado", "Fecha Creación", "Fecha Modificación", "Notas"
]

PRECINTOS_COL_WIDTHS = [14, 16, 20, 14, 20, 20, 30]


# ── Helpers ───────────────────────────────────────────────────

def _precintos_check_pin(pin) -> bool:
    expected = os.getenv("PRECINTOS_PIN", "").strip()
    if not expected:
        raise HTTPException(500, "PRECINTOS_PIN no configurado en el fichero .env")
    return str(pin).strip() == expected


def _precintos_ensure_file():
    """Create the Excel file with headers if it doesn't exist."""
    if PRECINTOS_FILE.exists():
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Precintos"
    ws.append(PRECINTOS_HEADERS)
    hdr_fill = PatternFill(start_color="1a4d7e", end_color="1a4d7e", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")
    for i, cell in enumerate(ws[1]):
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[cell.column_letter].width = PRECINTOS_COL_WIDTHS[i]
    wb.save(PRECINTOS_FILE)


def _precintos_load():
    """Load workbook. Always call _precintos_ensure_file() first."""
    _precintos_ensure_file()
    return load_workbook(PRECINTOS_FILE)


def _precintos_get_records(ws) -> list:
    records = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        def _fmt_dt(v):
            if v is None:
                return ""
            if hasattr(v, 'strftime'):
                return v.strftime("%d/%m/%Y %H:%M")
            return str(v)[:16]
        records.append({
            "num":               int(row[0]),
            "matricula":         row[1] or "",
            "expediente":        row[2] or "",
            "estado":            row[3] or "ACTIVO",
            "fecha_creacion":    _fmt_dt(row[4]),
            "fecha_modificacion":_fmt_dt(row[5]),
            "notas":             row[6] or "",
        })
    return records


def _precintos_next_num(records: list) -> int:
    """Next number = max of all non-CANCELADO + 1 (CANCELADO frees the slot conceptually
    but we keep it simple: max of ALL + 1, unless there are no records yet)."""
    active = [r["num"] for r in records if r["estado"] != "CANCELADO"]
    if not active:
        # Check if there are any records at all
        all_nums = [r["num"] for r in records]
        return (max(all_nums) + 1) if all_nums else 1
    return max(active) + 1


def _precintos_find_row_index(ws, num: int):
    """Return 1-based row index for a given precinto number, or None."""
    for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[0] is not None and int(row[0]) == num:
            return idx
    return None


def _precintos_save_autowidth(wb, ws):
    for i, col in enumerate(ws.columns):
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)
    wb.save(PRECINTOS_FILE)


# ── Endpoints ─────────────────────────────────────────────────

@app.get("/precintos", response_class=HTMLResponse)
async def precintos_page(request: Request):
    return templates.TemplateResponse("precintos.html", {"request": request})


@app.get("/api/precintos/list")
def precintos_list():
    with _precintos_lock:
        wb = _precintos_load()
        ws = wb.active
        records = _precintos_get_records(ws)

    records_sorted = sorted(records, key=lambda r: r["num"], reverse=True)
    stats = {
        "total":     len(records),
        "activos":   sum(1 for r in records if r["estado"] == "ACTIVO"),
        "reservados":sum(1 for r in records if r["estado"] == "RESERVADO"),
        "anulados":  sum(1 for r in records if r["estado"] == "ANULADO"),
        "cancelados":sum(1 for r in records if r["estado"] == "CANCELADO"),
    }
    return JSONResponse({
        "records":  records_sorted,
        "next_num": _precintos_next_num(records),
        "stats":    stats,
    })


@app.post("/api/precintos/nuevo")
async def precintos_nuevo(request: Request):
    """Register a new precinto (ACTIVO) or reserve one (RESERVADO)."""
    data = await request.json()
    if not _precintos_check_pin(data.get("pin", "")):
        raise HTTPException(401, "PIN incorrecto")

    matricula  = data.get("matricula",  "").strip()
    expediente = data.get("expediente", "").strip()
    notas      = data.get("notas",      "").strip()
    estado     = "ACTIVO" if (matricula and expediente) else "RESERVADO"
    now        = datetime.now()

    with _precintos_lock:
        wb = _precintos_load()
        ws = wb.active
        records  = _precintos_get_records(ws)
        next_num = _precintos_next_num(records)
        ws.append([next_num, matricula, expediente, estado, now, now, notas])
        _precintos_save_autowidth(wb, ws)

    return JSONResponse({
        "success": True,
        "num":     next_num,
        "estado":  estado,
        "msg":     f"Precinto {next_num} {'registrado' if estado == 'ACTIVO' else 'reservado'} correctamente",
    })


@app.post("/api/precintos/editar")
async def precintos_editar(request: Request):
    """Edit matrícula, expediente or notas for an existing precinto."""
    data = await request.json()
    if not _precintos_check_pin(data.get("pin", "")):
        raise HTTPException(401, "PIN incorrecto")

    num        = int(data.get("num", 0))
    matricula  = data.get("matricula",  "").strip()
    expediente = data.get("expediente", "").strip()
    notas      = data.get("notas",      "").strip()

    with _precintos_lock:
        wb  = _precintos_load()
        ws  = wb.active
        row = _precintos_find_row_index(ws, num)
        if row is None:
            raise HTTPException(404, f"Precinto {num} no encontrado")

        current_estado = ws.cell(row=row, column=4).value
        if current_estado in ("ANULADO",):
            raise HTTPException(400, f"No se puede editar un precinto {current_estado}")

        ws.cell(row=row, column=2).value = matricula
        ws.cell(row=row, column=3).value = expediente
        # If both fields are now filled and it was RESERVADO → promote to ACTIVO
        if current_estado == "RESERVADO" and matricula and expediente:
            ws.cell(row=row, column=4).value = "ACTIVO"
        ws.cell(row=row, column=6).value = datetime.now()
        ws.cell(row=row, column=7).value = notas
        _precintos_save_autowidth(wb, ws)

    nuevo_estado = "ACTIVO" if (current_estado == "RESERVADO" and matricula and expediente) else current_estado
    return JSONResponse({"success": True, "estado": nuevo_estado, "msg": f"Precinto {num} actualizado"})


@app.post("/api/precintos/anular")
async def precintos_anular(request: Request):
    """Mark a precinto as ANULADO (bad/damaged). Number is permanently consumed."""
    data = await request.json()
    if not _precintos_check_pin(data.get("pin", "")):
        raise HTTPException(401, "PIN incorrecto")

    num   = int(data.get("num", 0))
    notas = data.get("notas", "").strip() or "Anulado por mal estado"

    with _precintos_lock:
        wb  = _precintos_load()
        ws  = wb.active
        row = _precintos_find_row_index(ws, num)
        if row is None:
            raise HTTPException(404, f"Precinto {num} no encontrado")
        ws.cell(row=row, column=4).value = "ANULADO"
        ws.cell(row=row, column=6).value = datetime.now()
        ws.cell(row=row, column=7).value = notas
        _precintos_save_autowidth(wb, ws)

    return JSONResponse({"success": True, "msg": f"Precinto {num} marcado como ANULADO"})


@app.post("/api/precintos/cancelar-reserva")
async def precintos_cancelar_reserva(request: Request):
    """Cancel a RESERVADO precinto, freeing it (marked CANCELADO)."""
    data = await request.json()
    if not _precintos_check_pin(data.get("pin", "")):
        raise HTTPException(401, "PIN incorrecto")

    num = int(data.get("num", 0))

    with _precintos_lock:
        wb  = _precintos_load()
        ws  = wb.active
        row = _precintos_find_row_index(ws, num)
        if row is None:
            raise HTTPException(404, f"Precinto {num} no encontrado")
        if ws.cell(row=row, column=4).value != "RESERVADO":
            raise HTTPException(400, "Solo se pueden cancelar precintos en estado RESERVADO")
        ws.cell(row=row, column=4).value  = "CANCELADO"
        ws.cell(row=row, column=6).value  = datetime.now()
        ws.cell(row=row, column=7).value  = "Reserva cancelada"
        _precintos_save_autowidth(wb, ws)

    return JSONResponse({"success": True, "msg": f"Reserva del precinto {num} cancelada"})

# ============================================================
# ENTREGA DE EPIs (FR-58)
# Añadir este bloque a app.py (antes del if __name__ == "__main__")
# ============================================================

RESPONSABLE_PREVENCION = "Juan Antonio Martínez Lázaro"

TEXTO_LEGAL_EPIS = (
    "De acuerdo con el art. 29.2 de la Ley 31/1995 de 8 de Noviembre, de Prevención de Riesgos "
    "Laborales, así como del sistema de gestión de la prevención de riesgos, el trabajador:\n\n"
    "a) Se compromete a utilizar correctamente los medios y equipos de protección facilitados.\n"
    "b) Se responsabiliza de su mantenimiento y conservación.\n"
    "c) Entiende que el equipo se le asigna de manera personal.\n"
    "d) Devuelve el EPI usado o deteriorado antes de recibir otro nuevo, quedando prohibida la "
    "utilización de equipos deteriorados o caducos.\n"
    "e) Colabora con la gestión documentada para la entrega de los equipos.\n"
    "f) Comunicará al responsable del grupo o trabajador-enlace la pérdida, merma, deterioro o "
    "caducidad que pudiera sufrir el equipo de protección individual."
)


def _generate_epis_docx(data: dict):
    """Generate FR-58 EPI delivery DOCX. Returns (Path, filename)."""
    from docx import Document
    from docx.shared import Pt, Cm
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    def set_run(run, bold=False, size=10):
        run.font.name = 'Calibri'
        run.font.size = Pt(size)
        run.font.bold = bold

    def cell_write(cell, text, bold=False, size=10, center=False):
        for child in list(cell.paragraphs[0]._p):
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag in ('r', 'sdt', 'hyperlink'):
                cell.paragraphs[0]._p.remove(child)
        run = cell.paragraphs[0].add_run(str(text) if text else '')
        set_run(run, bold=bold, size=size)
        if center:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    def shade_cell(cell, hex_color):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), hex_color)
        tcPr.append(shd)

    empleado = data.get('empleado', '')
    puesto   = data.get('puesto', '')
    fecha    = data.get('fecha', '')
    epis     = data.get('epis', [])

    # ── HEADER TABLE ──────────────────────────────────────
    t0 = doc.add_table(rows=3, cols=4)
    t0.style = 'Table Grid'

    # Row 0: Puesto | value | DATOS EVALUACIÓN | blank
    cell_write(t0.cell(0, 0), 'Puesto de Trabajo:', bold=True)
    cell_write(t0.cell(0, 1), puesto.upper())
    cell_write(t0.cell(0, 2), 'DATOS DE LA EVALUACIÓN DE RIESGOS', bold=True)
    # Row 1: Trabajador | value | FECHA | date
    cell_write(t0.cell(1, 0), 'Trabajador:', bold=True)
    cell_write(t0.cell(1, 1), empleado.upper())
    cell_write(t0.cell(1, 2), 'FECHA:', bold=True)
    cell_write(t0.cell(1, 3), fecha)
    # Row 2: Fecha | mm/yyyy | REALIZACIÓN | COORDINADOR
    mes_anio = ''
    if fecha:
        try:
            from datetime import datetime as _dt
            mes_anio = _dt.strptime(fecha, '%Y-%m-%d').strftime('%m/%Y')
        except Exception:
            mes_anio = fecha
    cell_write(t0.cell(2, 0), 'Fecha:', bold=True)
    cell_write(t0.cell(2, 1), mes_anio)
    cell_write(t0.cell(2, 2), 'REALIZACIÓN:', bold=True)
    cell_write(t0.cell(2, 3), 'COORDINADOR')

    doc.add_paragraph()

    # ── EPI TABLE ─────────────────────────────────────────
    title_p = doc.add_paragraph()
    r = title_p.add_run("Listado de EPI's")
    set_run(r, bold=True, size=11)

    # Group by category
    from itertools import groupby
    categories = {}
    for epi in epis:
        cat = epi.get('categoria', '')
        categories.setdefault(cat, []).append(epi)

    # Build table: 1 header row + category rows + item rows
    total_rows = 1 + sum(1 + len(v) for v in categories.values())
    t1 = doc.add_table(rows=total_rows, cols=4)
    t1.style = 'Table Grid'

    # Column widths
    for row in t1.rows:
        row.cells[0].width = Cm(2.2)   # Normativa
        row.cells[1].width = Cm(8.5)   # EPI
        row.cells[2].width = Cm(1.2)   # SÍ
        row.cells[3].width = Cm(4.1)   # Observaciones

    # Header row
    cell_write(t1.rows[0].cells[0], 'Normativa', bold=True, center=True)
    cell_write(t1.rows[0].cells[1], 'Equipo de Protección Individual', bold=True)
    cell_write(t1.rows[0].cells[2], 'SÍ', bold=True, center=True)
    cell_write(t1.rows[0].cells[3], 'Observaciones', bold=True)
    for j in range(4):
        shade_cell(t1.rows[0].cells[j], 'd9e1f2')

    row_idx = 1
    for cat_name, items in categories.items():
        # Category header row
        t1.cell(row_idx, 0).merge(t1.cell(row_idx, 3))
        cell_write(t1.rows[row_idx].cells[0], cat_name.upper(), bold=True)
        shade_cell(t1.rows[row_idx].cells[0], 'f2f2f2')
        row_idx += 1

        for epi in items:
            cell_write(t1.rows[row_idx].cells[0], epi.get('normativa', ''), size=9, center=True)
            cell_write(t1.rows[row_idx].cells[1], epi.get('nombre', ''))
            cell_write(t1.rows[row_idx].cells[2],
                       '✓' if epi.get('entregado') else '', center=True, size=12)
            cell_write(t1.rows[row_idx].cells[3], epi.get('observaciones', ''), size=9)
            row_idx += 1

    doc.add_paragraph()

    # ── LEGAL TEXT ────────────────────────────────────────
    for line in TEXTO_LEGAL_EPIS.split('\n'):
        if not line.strip():
            continue
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(line)
        set_run(run, size=9)

    doc.add_paragraph()

    # ── SIGNATURES ────────────────────────────────────────
    t2 = doc.add_table(rows=2, cols=2)
    t2.style = 'Table Grid'
    cell_write(t2.cell(0, 0), 'Fdo: Responsable Prevención', bold=True)
    cell_write(t2.cell(0, 1), 'Fdo: Trabajador', bold=True)
    cell_write(t2.cell(1, 0), f'Nombre: {RESPONSABLE_PREVENCION}')
    cell_write(t2.cell(1, 1), f'Nombre: {empleado.upper()}')

    # ── SAVE ─────────────────────────────────────────────
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    slug = empleado.replace(' ', '_').upper()
    filename = f"FR58_EPIs_{slug}_{timestamp}.docx"
    filepath = OUTPUT_DIR / filename
    doc.save(str(filepath))
    return filepath, filename


@app.post("/api/save-entrega-epis")
async def save_entrega_epis(request: Request, generate_docx: bool = False):
    """
    Guarda una entrega de EPIs:
    1. Añade los EPIs entregados al Excel histórico (una fila por EPI)
    2. Si generate_docx=true, genera el FR-58 y devuelve URL de descarga

    Body JSON:
    {
      "empleado": "NOMBRE APELLIDOS",
      "puesto": "MOTERO",
      "fecha": "2026-04-09",
      "epis": [
        {"categoria": "Vestuario laboral", "normativa": "", "nombre": "Pantalón mono",
         "entregado": true, "observaciones": ""},
        ...
      ]
    }
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

    # ── 1. Excel ──────────────────────────────────────────
    excel_file = EXCEL_STORAGE_DIR / "entrega-epis.xlsx"

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

    # ── 2. DOCX ───────────────────────────────────────────
    if generate_docx:
        try:
            _, docx_filename = _generate_epis_docx(data)
            response["docx_filename"] = docx_filename
            response["download_url"]  = f"/download/{docx_filename}"
        except Exception as e:
            print(f"[epis] Error generando DOCX: {e}")
            response["docx_error"] = str(e)

    return JSONResponse(response)

if __name__ == "__main__":
    print("=" * 60)
    print("🚀 JSON Schema Form Generator")
    print("=" * 60)
    print(f"📁 Schemas directory: {SCHEMAS_DIR}")
    print(f"📁 Output directory: {OUTPUT_DIR}")
    print(f"🌐 Server starting at: http://localhost:8200")
    print("=" * 60)
    print()
    # Leer puerto de variable de entorno
    port = int(os.getenv("FORM_UI_PORT", "8200"))
    print(f"[INFO] Starting Form Generator on port {port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
