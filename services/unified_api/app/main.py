"""Unified API for all tools."""
from __future__ import annotations

import io
import os
import zipfile
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import List
import random
import string

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Path as PathParam # pyright: ignore[reportMissingImports]
from fastapi.responses import JSONResponse, StreamingResponse # pyright: ignore[reportMissingImports]

from .tool_registry import ToolRegistry
from iasuite_common.status import StatusManager # pyright: ignore[reportMissingImports]
import re as _re


CONFIG_PATH = os.getenv("CONFIG_PATH", "/config/tools.yaml")
DATA_ROOT = Path(os.getenv("DATA_ROOT", "/data"))

app = FastAPI(title="Invoice Automation Suite API", version="2.0")
registry = ToolRegistry.from_yaml(CONFIG_PATH)


def generate_batch_id() -> str:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    suffix = "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(4))
    return f"{timestamp}_{suffix}"


@app.get("/health")
def health():
    return {"ok": True, "tools": registry.list_tools(), "data_root": str(DATA_ROOT)}


@app.get("/tools")
def list_tools():
    return {
        "tools": [
            {"name": t.name, "display_name": t.display_name,
             "description": t.description, "input_formats": t.input_formats}
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


# ── Version endpoints ─────────────────────────────────────────────────────────

@app.get("/api/lear_rabat/version")
def lear_rabat_version():
    path = "/app/apps/lear_rabat/extractor/extract_lear_rabat.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/lear_cable/version")
def lear_cable_version():
    path = "/app/apps/lear_cable/extractor/extract_lear_fields.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/import_partida/version")
def import_partida_version():
    path = "/app/apps/import_partida/extractor/extract_import_partida_fields.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/croton/version")
def croton_version():
    path = "/app/apps/croton/extractor/extract_croton.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/cuadre_asientos/version")
def cuadre_asientos_version():
    path = "/app/apps/cuadre_asientos/extractor/cuadre_asientos_wrapper.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/camion/version")
def camion_version():
    path = "/app/apps/camion/extractor/camion_export_processor.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/croton_import/version")
def croton_import_version():
    path = "/app/apps/croton_import/extractor/extract_croton_import.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/split_nominas/version")
def split_nominas_version():
    path = "/app/apps/split_nominas/extractor/split_nominas.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/split_cotizaciones/version")
def split_cotizaciones_version():
    path = "/app/apps/split_cotizaciones/extractor/split_cotizaciones.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}


# ── Shared ZIP download helper ────────────────────────────────────────────────

def _zip_batch_download(tool_name: str, batch_id: str, zip_prefix: str) -> StreamingResponse:
    """Shared ZIP download for split_* tools (no email, direct browser download)."""
    tool = registry.get_tool(tool_name)
    if not tool:
        raise HTTPException(404, f"Tool {tool_name} not found")

    status_mgr = StatusManager(tool.status_dir)
    status = status_mgr.get_status(batch_id)
    if not status:
        raise HTTPException(404, f"Batch not found: {batch_id}")
    if status.state not in ("DONE", "done"):
        raise HTTPException(409, f"Batch not ready for download (state={status.state})")

    output_path = tool.output_dir / batch_id
    if not output_path.exists():
        raise HTTPException(404, f"Output directory not found for batch {batch_id}")

    pdf_files = sorted(output_path.glob("*.pdf"))
    if not pdf_files:
        raise HTTPException(404, "No PDF files found in batch output")

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for pdf_path in pdf_files:
            zf.write(pdf_path, arcname=pdf_path.name)
    zip_buffer.seek(0)

    zip_filename = f"{zip_prefix}_{batch_id}.zip"
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={
            "Content-Disposition": f'attachment; filename="{zip_filename}"',
            "X-File-Count": str(len(pdf_files)),
        },
    )


# ── Download endpoints ────────────────────────────────────────────────────────

@app.get("/api/split_nominas/batches/{batch_id}/download")
def split_nominas_download(batch_id: str = PathParam(...)):
    """Stream ZIP of split nóminas PDFs."""
    return _zip_batch_download("split_nominas", batch_id, "nominas")


@app.get("/api/split_cotizaciones/batches/{batch_id}/download")
def split_cotizaciones_download(batch_id: str = PathParam(...)):
    """Stream ZIP of split cotizaciones PDFs."""
    return _zip_batch_download("split_cotizaciones", batch_id, "cotizaciones")


# ── Batch creation ────────────────────────────────────────────────────────────

@app.post("/api/{tool_name}/batches")
async def create_batch(
    tool_name: str = PathParam(...),
    files: List[UploadFile] = File(...),
    skip_validation: str = Form("0"),
    run_t1:          str = Form("1"),
    run_dae:         str = Form("1"),
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
        run_t1_flag  = str(run_t1).strip().lower()  not in {"0", "false", "no", "off"}
        run_dae_flag = str(run_dae).strip().lower()  not in {"0", "false", "no", "off"}
        if tool_name == "camion":
            if skip_validation_flag:
                (batch_inbox / "_SKIP_VALIDATION").write_text("1", encoding="utf-8")
            if not run_t1_flag:
                (batch_inbox / "_SKIP_T1").write_text("1", encoding="utf-8")
            if not run_dae_flag:
                (batch_inbox / "_SKIP_DAE").write_text("1", encoding="utf-8")

        file_count = 0
        for upload_file in files:
            content = await upload_file.read()
            if not content:
                continue

            filename = upload_file.filename or f"file_{file_count}"
            safe_filename = Path(filename).name

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

        status_mgr = StatusManager(tool.status_dir)
        status_mgr.create_status(
            batch_id=batch_id,
            state="UPLOADED",
            stage="WAITING",
            total_files=file_count,
            message="Batch uploaded"
        )

        (batch_inbox / "_DONE").touch()

        return {
            "batch_id": batch_id,
            "tool": tool_name,
            "files_uploaded": file_count,
            "status": "UPLOADED",
            "skip_validation": skip_validation_flag if tool_name == "camion" else False,
            "run_t1":  run_t1_flag  if tool_name == "camion" else True,
            "run_dae": run_dae_flag if tool_name == "camion" else True,
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