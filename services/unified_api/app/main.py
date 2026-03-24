"""Unified API for all tools."""
from __future__ import annotations

import os
import zipfile
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import List
import random
import string

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Path as PathParam
from fastapi.responses import JSONResponse

from .tool_registry import ToolRegistry
from iasuite_common.status import StatusManager
import re as _re


# Configuration
CONFIG_PATH = os.getenv("CONFIG_PATH", "/config/tools.yaml")
DATA_ROOT = Path(os.getenv("DATA_ROOT", "/data"))

# Initialize
app = FastAPI(title="Invoice Automation Suite API", version="2.0")
registry = ToolRegistry.from_yaml(CONFIG_PATH)


def generate_batch_id() -> str:
    """Generate unique batch ID."""
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    suffix = "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(4))
    return f"{timestamp}_{suffix}"


@app.get("/health")
def health():
    """Health check."""
    return {
        "ok": True,
        "tools": registry.list_tools(),
        "data_root": str(DATA_ROOT)
    }


@app.get("/tools")
def list_tools():
    """List all tools."""
    return {
        "tools": [
            {
                "name": t.name,
                "display_name": t.display_name,
                "description": t.description,
                "input_formats": t.input_formats
            }
            for t in registry.tools
        ]
    }

def _read_script_version(path: str) -> str:
    try:
        text = Path(path).read_text()
        m = _re.search(r'SCRIPT_VERSION\s*=\s*["\']([^"\']+)["\']', text)
        return m.group(1) if m else "unknown"
    except Exception:
        return "unknown"

def _read_script_changelog(path: str) -> str:
    try:
        with open(path, encoding="utf-8") as f:
            src = f.read()
        m = _re.search(r'SCRIPT_CHANGELOG\s*=\s*"""(.*?)"""', src, _re.DOTALL)
        return m.group(1).strip() if m else ""
    except Exception:
        return ""

@app.get("/api/lear_rabat/version")
def lear_rabat_version():
    path = "/app/apps/lear_rabat/extractor/extract_lear_rabat.py"
    return {
        "version": _read_script_version(path),
        "changelog": _read_script_changelog(path),
    }

@app.get("/api/lear_cable/version")
def lear_cable_version():
    path = "/app/apps/lear_cable/extractor/extract_lear_fields.py"
    return {
        "version": _read_script_version(path),
        "changelog": _read_script_changelog(path),
        }

@app.get("/api/import_partida/version")
def import_partida_version():
    path = "/app/apps/import_partida/extractor/extract_import_partida_fields.py"
    return {
        "version": _read_script_version(path),
        "changelog": _read_script_changelog(path),
    }

@app.get("/api/croton/version")
def croton_version():
    path = "/app/apps/croton/extractor/extract_croton.py"
    return {
        "version": _read_script_version(path),
        "changelog": _read_script_changelog(path),
    }

@app.get("/api/cuadre_asientos/version")
def cuadre_asientos_version():
    path = "/app/apps/cuadre_asientos/extractor/cuadre_asientos_wrapper.py"
    return {
        "version": _read_script_version(path),
        "changelog": _read_script_changelog(path),
    }

@app.get("/api/camion/version")
def camion_version():
    path = "/app/apps/camion/extractor/camion_export_processor.py"
    return {
        "version": _read_script_version(path),
        "changelog": _read_script_changelog(path),
    }

@app.post("/api/{tool_name}/batches")
async def create_batch(
    tool_name: str = PathParam(...),
    files: List[UploadFile] = File(...),
    skip_validation: str = Form("0"),
):
    """Create new batch."""
    tool = registry.get_tool(tool_name)
    if not tool:
        raise HTTPException(404, f"Tool not found: {tool_name}")
    
    if not files:
        raise HTTPException(400, "No files provided")
    
    batch_id = generate_batch_id()
    batch_inbox = tool.inbox_dir / batch_id
    
    try:
        batch_inbox.mkdir(parents=True, exist_ok=True)

        skip_validation_flag = str(skip_validation).strip().lower() in {"1", "true", "yes", "on"}
        if tool_name == "camion" and skip_validation_flag:
            (batch_inbox / "_SKIP_VALIDATION").write_text("1", encoding="utf-8")
        
        # Save all files
        file_count = 0
        for upload_file in files:
            content = await upload_file.read()
            if not content:
                continue
            
            # Use the original filename (which includes extension)
            filename = upload_file.filename or f"file_{file_count}"
            safe_filename = Path(filename).name
            
            # Validate file extension against tool's accepted formats
            file_ext = Path(safe_filename).suffix.lower().lstrip('.')
            if file_ext not in tool.input_formats:
                raise HTTPException(
                    400, 
                    f"Invalid file type .{file_ext}. Tool '{tool_name}' accepts: {', '.join(tool.input_formats)}"
                )
            
            (batch_inbox / safe_filename).write_bytes(content)
            file_count += 1
        
        if file_count == 0:
            raise HTTPException(400, "No valid files uploaded")
        
        # Create status
        status_mgr = StatusManager(tool.status_dir)
        status_mgr.create_status(
            batch_id=batch_id,
            state="UPLOADED",
            stage="WAITING",
            total_files=file_count,
            message="Batch uploaded"
        )
        
        # Mark as done
        (batch_inbox / "_DONE").touch()
        
        return {
            "batch_id": batch_id,
            "tool": tool_name,
            "files_uploaded": file_count,
            "status": "UPLOADED",
            "skip_validation": skip_validation_flag if tool_name == "camion" else False,
        }
    
    except HTTPException:
        raise
    except Exception as e:
        if batch_inbox.exists():
            shutil.rmtree(batch_inbox)
        raise HTTPException(500, f"Error: {str(e)}")


@app.get("/api/{tool_name}/batches/{batch_id}/status")
def get_status(
    tool_name: str = PathParam(...),
    batch_id: str = PathParam(...)
):
    """Get batch status."""
    tool = registry.get_tool(tool_name)
    if not tool:
        raise HTTPException(404, f"Tool not found: {tool_name}")
    
    status_mgr = StatusManager(tool.status_dir)
    status = status_mgr.get_status(batch_id)
    
    if not status:
        raise HTTPException(404, f"Batch not found: {batch_id}")
    
    return {
        "batch_id": status.batch_id,
        "state": status.state,
        "stage": status.stage,
        "total_files": status.total_files,
        "processed_files": status.processed_files,
        "message": status.message
    }

