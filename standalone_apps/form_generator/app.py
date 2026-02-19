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
from jsonschema import Draft202012Validator, ValidationError
import uvicorn
import os

# Configuration
SCHEMAS_DIR = Path(__file__).parent / "schemas"
OUTPUT_DIR = Path(__file__).parent / "output"
TEMPLATES_DIR = Path(__file__).parent / "templates"
STATIC_DIR = Path(__file__).parent / "static"

# Ensure directories exist
OUTPUT_DIR.mkdir(exist_ok=True)

# Initialize FastAPI
app = FastAPI(title="JSON Schema Form Generator", version="1.0")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)


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
    
    return templates.TemplateResponse("form.html", {
        "request": request,
        "schema_name": schema_name,
        "schema": schema,
        "schema_json": json.dumps(schema, ensure_ascii=False)
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
