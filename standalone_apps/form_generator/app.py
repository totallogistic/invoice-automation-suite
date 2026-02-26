#!/usr/bin/env python3
"""
Universal JSON Schema Form Generator - FastAPI Server
Generates HTML forms from JSON Schema and outputs JSON files.

Usage:
    python app.py
    
Then open: http://localhost:8200
"""

import json
from pathlib import Path
from typing import Any, Dict

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from jsonschema import Draft202012Validator, ValidationError, validate, Draft7Validator

import uvicorn
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
                    "description": schema.get("description", "")
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
