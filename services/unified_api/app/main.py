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

from fastapi import FastAPI, UploadFile, File, HTTPException, Path as PathParam
from fastapi.responses import JSONResponse

from .tool_registry import ToolRegistry
from iasuite_common.status import StatusManager


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


@app.post("/api/{tool_name}/batches")
async def create_batch(
    tool_name: str = PathParam(...),
    files: List[UploadFile] = File(...)
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
        
        # Save all files
        file_count = 0
        for upload_file in files:
            content = await upload_file.read()
            if not content:
                continue
            
            filename = upload_file.filename or f"file_{file_count}.pdf"
            safe_filename = Path(filename).name
            
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
            "status": "UPLOADED"
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
