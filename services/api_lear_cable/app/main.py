from __future__ import annotations

import json
import os
import random
import string
import time
import zipfile
import shutil
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse

app = FastAPI()

SERVICE_ROOT = Path(os.getenv("SERVICE_ROOT", "/data"))
INBOX_DIR = Path(os.getenv("INBOX_DIR", str(SERVICE_ROOT / "inbox")))
STATUS_DIR = Path(os.getenv("STATUS_DIR", str(SERVICE_ROOT / "status")))

ALLOWED_EXT = {".pdf"}  # filtra solo PDFs


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _rand(n: int = 4) -> str:
    return "".join(random.choice(string.ascii_lowercase + string.digits) for _ in range(n))


def new_batch_id() -> str:
    # Ej: 20260126_120233_ab12
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S") + "_" + _rand(4)


def atomic_write_json(path: Path, data: Dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def safe_extract_zip(zip_path: Path, dest_dir: Path) -> int:
    """
    Extrae ZIP evitando Zip Slip y filtrando solo PDFs.
    Devuelve el número de PDFs extraídos.
    """
    if not zipfile.is_zipfile(zip_path):
        raise HTTPException(status_code=400, detail="El fichero no es un ZIP válido.")

    dest_dir.mkdir(parents=True, exist_ok=True)
    base = dest_dir.resolve()

    extracted = 0
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue

            # Ignora basura típica de macOS
            name = info.filename
            if name.startswith("__MACOSX/") or name.endswith(".DS_Store"):
                continue

            rel = Path(name)

            # Si viene sin nombre “real”
            if not rel.name:
                continue

            # Filtra por extensión
            if rel.suffix.lower() not in ALLOWED_EXT:
                continue

            target = (dest_dir / rel).resolve()

            # Zip Slip protection
            if not str(target).startswith(str(base) + os.sep):
                raise HTTPException(status_code=400, detail="ZIP contiene rutas no permitidas (zip slip).")

            target.parent.mkdir(parents=True, exist_ok=True)

            with zf.open(info, "r") as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)

            extracted += 1

    return extracted


@app.get("/api/lear_cable/health")
def health():
    return {"ok": True, "inbox": str(INBOX_DIR)}


@app.post("/api/lear_cable/batches")
async def create_batch(files: List[UploadFile] = File(...)):
    """
    Sube uno o varios PDFs, o un ZIP con PDFs, lo procesa en /data/inbox/<batch_id>/ y deja status inicial en /data/status/<batch_id>/status.json
    """
    if not files:
        raise HTTPException(status_code=400, detail="No se recibieron archivos.")
    
    batch_id = new_batch_id()
    batch_inbox = INBOX_DIR / batch_id
    status_file = STATUS_DIR / batch_id / "status.json"

    extracted = 0
    tmp_file = None

    try:
        # If single file and it's a ZIP, extract it
        if len(files) == 1:
            file = files[0]
            content = await file.read()
            if not content:
                raise HTTPException(status_code=400, detail="Archivo vacío.")

            filename = file.filename or ""
            file_ext = Path(filename).suffix.lower()

            if file_ext == ".zip":
                # Handle ZIP file
                tmp_file = Path("/tmp") / f"{batch_id}.zip"
                with open(tmp_file, "wb") as f:
                    f.write(content)

                extracted = safe_extract_zip(tmp_file, batch_inbox)
                if extracted == 0:
                    raise HTTPException(status_code=400, detail="No se encontraron PDFs en el ZIP.")

            elif file_ext == ".pdf":
                # Handle single PDF file
                batch_inbox.mkdir(parents=True, exist_ok=True)
                safe_filename = Path(filename).name
                pdf_path = batch_inbox / safe_filename
                with open(pdf_path, "wb") as f:
                    f.write(content)
                extracted = 1

            else:
                raise HTTPException(
                    status_code=400,
                    detail=f"Tipo de archivo no permitido: {file_ext}. Solo se aceptan .pdf o .zip"
                )
        
        else:
            # Handle multiple files - all must be PDFs
            batch_inbox.mkdir(parents=True, exist_ok=True)
            
            for file in files:
                content = await file.read()
                filename = file.filename or ""
                
                if not content:
                    raise HTTPException(
                        status_code=400,
                        detail=f"Archivo vacío detectado: {filename}"
                    )
                
                file_ext = Path(filename).suffix.lower()
                
                if file_ext != ".pdf":
                    raise HTTPException(
                        status_code=400,
                        detail=f"Cuando se suben múltiples archivos, todos deben ser PDF. Archivo inválido: {filename}"
                    )
                
                # Sanitize filename to prevent directory traversal
                safe_filename = Path(filename).name
                pdf_path = batch_inbox / safe_filename
                with open(pdf_path, "wb") as f:
                    f.write(content)
                extracted += 1
            
            if extracted == 0:
                raise HTTPException(status_code=400, detail="No se recibieron archivos PDF válidos.")

        # status inicial (el watcher lo irá actualizando)
        status = {
            "batch_id": batch_id,
            "state": "UPLOADED",          # luego: RUNNING / DONE / ERROR
            "stage": "WAITING",           # luego: EXTRACTING / EMAILING / FINALIZING
            "total_files": extracted,     # PDFs detectados
            "processed_files": 0,
            "started_at": _now_iso(),
            "updated_at": _now_iso(),
            "recipients": [],             # el watcher lo rellena
            "message": "Lote subido. En espera de procesamiento.",
        }
        atomic_write_json(status_file, status)

        return JSONResponse(
            {
                "batch_id": batch_id,
                "extracted_pdfs": extracted,
                "status_url": f"/api/lear_cable/batches/{batch_id}/status",
            }
        )
    finally:
        if tmp_file:
            try:
                tmp_file.unlink(missing_ok=True)
            except Exception:
                pass


@app.get("/api/lear_cable/batches/{batch_id}/status")
def get_batch_status(batch_id: str):
    status_file = STATUS_DIR / batch_id / "status.json"

    if not status_file.exists():
        # Aún no hay status (p.ej. watcher no lo ha creado)
        return {
            "batch_id": batch_id,
            "state": "PENDING",
            "stage": "WAITING",
            "total_files": 0,
            "processed_files": 0,
            "started_at": None,
            "updated_at": _now_iso(),
            "recipients": [],
            "message": "Sin estado aún (pendiente).",
        }

    try:
        data = json.loads(status_file.read_text(encoding="utf-8"))
        return data
    except Exception:
        raise HTTPException(status_code=500, detail="No se pudo leer status.json.")
