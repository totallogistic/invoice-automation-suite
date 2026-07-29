"""Handler de la operativa DeCA para form_generator (patrón hoja-control).

Bloque para añadir a `standalone_apps/form_generator/app.py`. Sigue exactamente
el patrón de `save_hoja_control_expedientes`: recibe el JSON del formulario,
ejecuta la operativa definida y devuelve JSON. Diferencia: la operativa es
generar PDF+QR y **subirlo al bucket** (no xlsx→PDF→email), devolviendo la URL
que se enseña/envía al conductor.

Requisitos:
  - Copiar el paquete motor a `standalone_apps/form_generator/deca/`
    (models, pdf, storage, repository, service, version).
  - Añadir `boto3`, `reportlab`, `segno` a form_generator/requirements.txt.
  - Variables DECA_* en el entorno del servicio (systemd / .env.prod).

Versionado: al ser un formulario, la versión sale del schema JSON y lo recoge
`/api/forms/status` automáticamente (ver claude/convenciones-version-y-health.md).
No necesita entrada en TOOL_VERSION_FILES ni /api/<tool>/version.
"""

import base64

from fastapi import HTTPException, Request           # ya importados en app.py
from fastapi.responses import JSONResponse           # ya importado en app.py

from deca.models import DecaInput
from deca.pdf import qr_png_bytes
from deca.service import DecaService

_deca_svc = DecaService()


@app.post("/api/save-deca")                          # noqa: F821  (app definido en app.py)
async def save_deca(request: Request):
    """Genera el DeCA (PDF nativo + QR), lo sube al bucket y devuelve la URL
    pública + el QR para enseñar/enviar al conductor."""
    payload = await request.json()
    try:
        data = DecaInput(**payload)
    except Exception as e:
        raise HTTPException(400, f"Datos del DeCA inválidos: {e}")

    try:
        rec = _deca_svc.create(data)                 # datos→URL→PDF+QR→PUT bucket→registro
    except ValueError as e:                           # p.ej. PDF > 5 MB
        raise HTTPException(400, str(e))

    qr_b64 = base64.b64encode(qr_png_bytes(rec.url_publica)).decode()
    return JSONResponse({
        "success": True,
        "uuid": rec.uuid,
        "url": rec.url_publica,
        "qr_data_uri": f"data:image/png;base64,{qr_b64}",
        "url_activa_hasta": rec.url_activa_hasta.isoformat() if rec.url_activa_hasta else None,
        "message": "DeCA generado y subido. Envía el QR o el enlace al conductor.",
    })
