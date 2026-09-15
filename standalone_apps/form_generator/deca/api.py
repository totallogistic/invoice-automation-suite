"""API del prototipo DeCA (FastAPI).

En el prototipo corre como servicio propio. Para integrarlo en el stack, estos
endpoints se añaden a `services/unified_api/app/main.py` (ver README §Integración).

Sigue la convención del stack: `GET /api/deca/version` devuelve
`{version, changelog}` leídos de las constantes SCRIPT_VERSION/SCRIPT_CHANGELOG.
"""

from __future__ import annotations

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import JSONResponse

from .models import DecaInput
from .service import DecaService
from .version import SCRIPT_CHANGELOG, SCRIPT_VERSION

app = FastAPI(title="DeCA — Documento Electrónico de Control Administrativo", version="1.0")
svc = DecaService()


@app.get("/health")
def health():
    return {"ok": True, "storage_mode": svc.storage.mode}


@app.get("/api/deca/version")
def deca_version():
    return {"version": SCRIPT_VERSION, "changelog": SCRIPT_CHANGELOG.strip()}


@app.post("/api/deca")
def create_deca(data: DecaInput):
    """Genera un DeCA: PDF+QR, lo sube al bucket y devuelve la URL pública."""
    try:
        rec = svc.create(data)
    except ValueError as e:
        raise HTTPException(400, str(e))
    return {
        "uuid": rec.uuid,
        "url": rec.url_publica,
        "creado_en": rec.creado_en.isoformat(),
        "url_activa_hasta": rec.url_activa_hasta.isoformat() if rec.url_activa_hasta else None,
        "storage_mode": svc.storage.mode,
    }


@app.get("/api/deca/{uuid}")
def get_deca(uuid: str):
    rec = svc.get(uuid)
    if not rec:
        raise HTTPException(404, "DeCA no encontrado")
    return rec


@app.patch("/api/deca/{uuid}")
def modify_deca(uuid: str):
    """Registra una modificación en ruta (nuevo timestamp), como contempla la norma."""
    ts = svc.modify(uuid)
    if ts is None:
        raise HTTPException(404, "DeCA no encontrado")
    return {"uuid": uuid, "modificado_en": ts.isoformat()}


# ── DEV ONLY ──────────────────────────────────────────────────────────────────
# En producción el bucket sirve el PDF (host fuera del camino crítico). En dev,
# sin bucket, esta ruta hace de "bucket" para poder probar el escaneo en local.
@app.get("/d/{name}")
def dev_serve(name: str):
    data = svc.storage.read_local(f"d/{name}")
    if data is None:
        raise HTTPException(404, "objeto no encontrado (¿modo s3?)")
    return Response(content=data, media_type="application/pdf")
