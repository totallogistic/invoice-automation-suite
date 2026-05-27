"""Unified API for all tools."""
from __future__ import annotations

import subprocess
import io
import os
import csv
import json
import zipfile
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import List
import random
import string

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Path as PathParam, Query, Request # pyright: ignore[reportMissingImports]
from fastapi.responses import JSONResponse, StreamingResponse, FileResponse # pyright: ignore[reportMissingImports]

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


def _cleanup_orphan_inbox(batch_inbox: Path) -> None:
    """
    Borra una carpeta de batch si quedó huérfana tras un fallo de validación.

    Una carpeta se considera huérfana si:
      - No existe (no hay nada que limpiar).
      - Está vacía.
      - Solo contiene markers internos (archivos cuyo nombre empieza por '_'),
        sin archivos de datos reales del cliente.

    Si hay archivos de datos (XLSX/PDF/CSV), NO se borra para preservar
    evidencia de uploads parciales (útil para diagnóstico).
    """
    if not batch_inbox.exists():
        return
    try:
        # Listar archivos NO-marker (los marker empiezan por '_')
        data_files = [
            f for f in batch_inbox.iterdir()
            if f.is_file() and not f.name.startswith("_")
        ]
        if not data_files:
            shutil.rmtree(batch_inbox, ignore_errors=True)
    except Exception:
        # Si algo falla durante el cleanup, lo ignoramos:
        # mejor dejar la carpeta huérfana que romper el response error original.
        pass


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

@app.get("/api/intrastat/version")
def intrastat_version():
    # La versión / changelog viven en el script real, no en el wrapper.
    path = "/app/apps/intrastat/extractor/intrastat_generator.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/lear_rabat/version")
def lear_rabat_version():
    path = "/app/apps/lear_rabat/extractor/extract_lear_rabat.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/lear_cable/version")
def lear_cable_version():
    path = "/app/apps/lear_cable/extractor/extract_lear_fields.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/lear_tac/version")
def lear_tac_version():
    path = "/app/apps/lear_tac/extractor/extract_lear_tac_fields.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/lear_kenitra/version")
def lear_kenitra_version():
    path = "/app/apps/lear_kenitra/extractor/extract_lear_kenitra_fields.py"
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
    path = "/app/apps/camion/extractor/run_processors.py"
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

@app.get("/api/bl/version")
def bl_version():
    path = "/app/apps/bl/extractor/extract_bl.py"
    return {"version": _read_script_version(path), "changelog": _read_script_changelog(path)}

@app.get("/api/export_visual/version")
def export_visual_version():
    path = "/app/apps/export_visual/extractor/export_visual.py"
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
    run_dae:         str = Form("1"),
    cliente: str = Form("aldi"),
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
        run_dae_flag = str(run_dae).strip().lower()  not in {"0", "false", "no", "off"}
        if tool_name == "camion":
            # Validación cruzada: --no-dae requiere --t1
            if not run_dae_flag:
                has_t1 = any(
                    "t1" in (f.filename or "").lower() and (f.filename or "").lower().endswith(".pdf")
                    for f in files
                )
                if not has_t1:
                    raise HTTPException(
                        400,
                        "Si desmarcas 'Generar DAE' debes subir al menos un T1 PDF "
                        "(no se generaría ningún output útil)."
                    )
            if skip_validation_flag:
                (batch_inbox / "_SKIP_VALIDATION").write_text("1", encoding="utf-8")
            if not run_dae_flag:
                (batch_inbox / "_SKIP_DAE").write_text("1", encoding="utf-8")
        if tool_name == "export_visual":
            cliente_val = str(cliente).strip() or "aldi"
            (batch_inbox / "_CLIENTE.txt").write_text(cliente_val, encoding="utf-8")
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
            "run_dae": run_dae_flag if tool_name == "camion" else True,
        }

    except HTTPException:
        # Limpiar carpeta huérfana si no contiene archivos útiles
        # (validación falló antes de escribir contenido — solo markers como _SKIP_*).
        _cleanup_orphan_inbox(batch_inbox)
        raise
    except Exception as e:
        if batch_inbox.exists():
            shutil.rmtree(batch_inbox, ignore_errors=True)
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


# ── BL endpoints ──────────────────────────────────────────────────────────────

BL_CSV_DIR   = Path(os.getenv("BL_CSV_DIR",   "/data/bl/csv"))
BL_DATA_ROOT = Path(os.getenv("BL_DATA_ROOT", "/data/bl"))
_ALGECIRAS   = "ALGECIRAS"
_HECHO_FILE  = BL_CSV_DIR / "bl_hecho.json"


def _load_hecho() -> set:
    try:
        return set(json.loads(_HECHO_FILE.read_text(encoding="utf-8")))
    except Exception:
        return set()


def _save_hecho(hechos: set) -> None:
    BL_CSV_DIR.mkdir(parents=True, exist_ok=True)
    _HECHO_FILE.write_text(json.dumps(sorted(hechos)), encoding="utf-8")


def _read_bl_csvs(fecha_desde: str | None, fecha_hasta: str | None) -> list[dict]:
    """Lee todos los CSVs bl_YYYYMMDD.csv cuya fecha de procesado esté en el rango."""
    records: list[dict] = []
    for csv_path in sorted(BL_CSV_DIR.glob("bl_????????.csv")):
        m = _re.match(r"bl_(\d{4})(\d{2})(\d{2})\.csv$", csv_path.name)
        if not m:
            continue
        file_date_iso = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
        if fecha_desde and file_date_iso < fecha_desde:
            continue
        if fecha_hasta and file_date_iso > fecha_hasta:
            continue
        try:
            with open(csv_path, newline="", encoding="utf-8") as fh:
                records.extend(csv.DictReader(fh))
        except Exception:
            pass
    return records


def _apply_bl_filters(
    records: list[dict],
    naviera: str,
    q: str,
    solo_hecho: str,
    tipo: str,
    puerto_destino: str,
    hechos: set,
) -> list[dict]:
    result = []
    q_low = q.strip().lower()
    pd_low = puerto_destino.strip().upper()

    for r in records:
        # Naviera
        if naviera and r.get("naviera", "").upper() != naviera.upper():
            continue
        # Tipo: exp = ALGECIRAS, imp = otros
        dest = (r.get("puerto_destino") or "").upper()
        if tipo == "exp" and _ALGECIRAS not in dest:
            continue
        if tipo == "imp" and _ALGECIRAS in dest:
            continue
        # Puerto destino libre
        if pd_low and pd_low not in dest:
            continue
        # Hecho state
        clave = f"{r.get('archivo','')}|{r.get('naviera','')}|{r.get('num_bl','')}"
        r["hecho"] = clave in hechos
        if solo_hecho == "1" and not r["hecho"]:
            continue
        if solo_hecho == "0" and r["hecho"]:
            continue
        # Búsqueda libre
        if q_low:
            haystack = " ".join([
                r.get("num_bl", ""), r.get("nombre", ""),
                r.get("matricula", ""), r.get("archivo", ""),
                r.get("buque", ""), r.get("puerto_destino", ""),
            ]).lower()
            if q_low not in haystack:
                continue
        result.append(r)

    result.sort(key=lambda r: (r.get("fecha", ""), r.get("hora", "")), reverse=True)
    return result


@app.get("/api/bl/registros")
def bl_registros(
    fecha_desde:    str = Query(default=""),
    fecha_hasta:    str = Query(default=""),
    naviera:        str = Query(default=""),
    q:              str = Query(default=""),
    solo_hecho:     str = Query(default=""),
    tipo:           str = Query(default=""),   # "exp" | "imp" | ""
    puerto_destino: str = Query(default=""),
):
    """Devuelve registros BL con filtros. Usado por el viewer para stats y tabla."""
    records  = _read_bl_csvs(fecha_desde or None, fecha_hasta or None)
    hechos   = _load_hecho()
    filtered = _apply_bl_filters(records, naviera, q, solo_hecho, tipo, puerto_destino, hechos)
    return {"total": len(filtered), "registros": filtered}


@app.get("/api/bl/stats")
def bl_stats():
    """Devuelve info sobre los CSVs disponibles."""
    csvs = sorted(
        [p.name for p in BL_CSV_DIR.glob("bl_????????.csv")],
        reverse=True,
    )
    return {"csvs": csvs}


@app.get("/api/bl/pdf/{filename}")
def bl_pdf(filename: str = PathParam(...)):
    """Sirve un PDF de BL buscándolo en las carpetas de datos."""
    safe = Path(filename).name  # evitar path traversal
    search_dirs = [
        BL_DATA_ROOT / "processed",
        BL_DATA_ROOT / "inbox",
        BL_DATA_ROOT / "processing",
        BL_DATA_ROOT / "error",
    ]
    for base in search_dirs:
        if not base.exists():
            continue
        # Buscar recursivamente en subdirectorios de batch
        for candidate in base.rglob(safe):
            if candidate.is_file():
                return FileResponse(
                    path=str(candidate),
                    media_type="application/pdf",
                    filename=safe,
                )
    raise HTTPException(404, f"PDF no encontrado: {safe}")


@app.post("/api/bl/toggle")
async def bl_toggle(request: Request):
    """Alterna el estado hecho/pendiente de un registro BL."""
    try:
        payload = await request.json()
    except Exception:
        raise HTTPException(400, "JSON inválido")
    clave = (payload.get("clave") or "").strip()
    if not clave:
        raise HTTPException(400, "clave requerida")
    hechos = _load_hecho()
    if clave in hechos:
        hechos.discard(clave)
        nuevo = False
    else:
        hechos.add(clave)
        nuevo = True
    _save_hecho(hechos)
    return {"clave": clave, "hecho": nuevo}


@app.post("/api/bl/sync")
async def bl_sync():
    """Fuerza la copia de BLs desde Google Drive al inbox local."""
    sync_script = Path(os.getenv("BL_SYNC_SCRIPT", "/opt/bl_sync.sh"))
    if not sync_script.exists():
        raise HTTPException(404, f"Script de sync no encontrado: {sync_script}")
    try:
        result = subprocess.run(
            ["bash", str(sync_script)],
            capture_output=True, text=True, timeout=120,
            env={
                **os.environ,
                "BL_INBOX":    os.getenv("BL_INBOX",    "/data/bl/inbox"),
                "RCLONE_CONF": os.getenv("RCLONE_CONF", "/root/.config/rclone/rclone.conf"),
            }
        )
        if result.returncode == 0:
            return {"ok": True, "message": "Sync completado correctamente"}
        else:
            raise HTTPException(500, f"Error en sync: {result.stderr.strip() or result.stdout.strip()}")
    except subprocess.TimeoutExpired:
        raise HTTPException(504, "Sync tardó demasiado (>120s)")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, str(e))
