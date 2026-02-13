from __future__ import annotations

import json
import os
import random
import string
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse

app = FastAPI()

SERVICE_ROOT = Path(os.getenv("SERVICE_ROOT", "/data"))
INBOX_DIR = Path(os.getenv("INBOX_DIR", str(SERVICE_ROOT / "inbox")))
STATUS_DIR = Path(os.getenv("STATUS_DIR", str(SERVICE_ROOT / "status")))

ALLOWED_EXT = {".pdf"}  # SOLO PDFs (import_partida)


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


@app.get("/api/import_partida/health")
def health():
    return {"ok": True, "inbox": str(INBOX_DIR)}


@app.post("/api/import_partida/batches")
async def create_batch(files: List[UploadFile] = File(...)):
    """
    Partida: acepta EXACTAMENTE 1 PDF.
    Lo deja en /data/inbox/<batch_id>/ y crea status inicial en /data/status/<batch_id>/status.json
    """
    if not files or len(files) != 1:
        raise HTTPException(status_code=400, detail="Debes subir exactamente 1 PDF.")

    file = files[0]
    filename = file.filename or ""
    file_ext = Path(filename).suffix.lower()

    if file_ext not in ALLOWED_EXT:
        raise HTTPException(status_code=400, detail="Tipo de archivo no permitido. Solo se acepta 1 PDF.")

    content = await file.read()
    if not content:
        raise HTTPException(status_code=400, detail="Archivo vacío.")

    batch_id = new_batch_id()
    batch_inbox = INBOX_DIR / batch_id
    status_file = STATUS_DIR / batch_id / "status.json"

    batch_inbox.mkdir(parents=True, exist_ok=True)

    # Sanitiza filename (sin rutas)
    safe_filename = Path(filename).name
    pdf_path = batch_inbox / safe_filename

    with open(pdf_path, "wb") as f:
        f.write(content)

    extracted = 1

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
            "status_url": f"/api/import_partida/batches/{batch_id}/status",
        }
    )


@app.get("/api/import_partida/batches/{batch_id}/status")
def get_batch_status(batch_id: str):
    status_file = STATUS_DIR / batch_id / "status.json"

    if not status_file.exists():
        # Aún no hay status
        return {
            "batch_id": batch_id,
            "state": "PENDING",
            "stage": "WAITING",
            "total_files": 0,
            "processed_files": 0,
            "started_at": None,
            "updated_at": None,
            "recipients": [],
            "message": "Sin status todavía.",
        }

    try:
        return json.loads(status_file.read_text(encoding="utf-8"))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error leyendo status: {e}")