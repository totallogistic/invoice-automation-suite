"""Orquestación del alta de un DeCA: datos -> URL -> PDF+QR -> subida -> registro.

Orden importante: la URL es DETERMINISTA (base + clave), así que se conoce ANTES
de renderizar → se incrusta en el QR → se sube el PDF a esa misma clave.
"""

from __future__ import annotations

import base64
import datetime as dt
import functools
import os
from pathlib import Path

from .models import DecaInput, DecaRecord, Firma
from .pdf import render_pdf
from .repository import Repository
from .storage import Storage

MAX_PDF_BYTES = 5 * 1024 * 1024  # 5 MB (límite de la norma)

# ── Firma corporativa de Total Logistic (sello/imagen) ────────────────────────
# Se inyecta SOLA en carta de porte / CMR (nunca en el DeCA, que no lleva firma).
# La imagen vive en una ruta del host (config), NO se sube desde el navegador:
# así no es manipulable desde el cliente y es una sola fuente de verdad.
#   DECA_FIRMA_TOTALLOGISTIC  → ruta del PNG/JPG (por defecto: deca/assets/firma_totallogistic.png)
#   DECA_FIRMA_TL_ROL         → etiqueta del rol (por defecto: "Expedidor";
#                               es la casilla "Firma y sello del Expedidor" del
#                               modelo de carta de porte. Pon "Almacén" si prefieres.)
#   DECA_FIRMA_TL_NOMBRE      → nombre mostrado (por defecto: "Total Logistic Services, S.L.")
_FIRMA_TL_PATH = os.getenv("DECA_FIRMA_TOTALLOGISTIC", "")
_FIRMA_TL_ROL = os.getenv("DECA_FIRMA_TL_ROL", "Expedidor")
_FIRMA_TL_NOMBRE = os.getenv("DECA_FIRMA_TL_NOMBRE", "Total Logistic Services, S.L.")


@functools.lru_cache(maxsize=1)
def _firma_tl_datauri():
    """Carga la imagen de la firma corporativa como data URI (o None si no existe)."""
    path = _FIRMA_TL_PATH or str(Path(__file__).with_name("assets") / "firma_totallogistic.png")
    p = Path(path)
    if not p.exists():
        return None
    ext = p.suffix.lstrip(".").lower()
    mime = "image/jpeg" if ext in ("jpg", "jpeg") else "image/png"
    return f"data:{mime};base64," + base64.b64encode(p.read_bytes()).decode()


class DecaService:
    def __init__(self) -> None:
        self.storage = Storage()
        self.repo = Repository()

    def _with_org_firma(self, data: DecaInput) -> DecaInput:
        """Añade la firma corporativa de Total Logistic a carta de porte / CMR.
        El DeCA no lleva firma. La del camionero (canvas manual) ya viene en
        data.firmas; la corporativa se antepone (emisor/almacén primero)."""
        if data.tipo_documento == "deca":
            return data
        datauri = _firma_tl_datauri()
        if not datauri:
            return data  # sin imagen configurada -> no se inyecta nada
        if any((f.nombre or "").strip() == _FIRMA_TL_NOMBRE for f in data.firmas):
            return data  # ya presente (evita duplicar)
        firma = Firma(rol=_FIRMA_TL_ROL, nombre=_FIRMA_TL_NOMBRE, firma_png=datauri,
                      firmado_en=dt.datetime.now(dt.timezone.utc))
        return data.model_copy(update={"firmas": [firma, *data.firmas]})

    def create(self, data: DecaInput) -> DecaRecord:
        data = self._with_org_firma(data)
        record = DecaRecord(datos=data)
        # 1) URL determinista (antes de generar el QR)
        record.url_publica = self.storage.public_url(record.object_key())
        record.creado_en = dt.datetime.now(dt.timezone.utc)
        # 2) PDF nativo con el QR que codifica esa URL
        pdf_bytes = render_pdf(record)
        if len(pdf_bytes) > MAX_PDF_BYTES:
            raise ValueError(f"PDF de {len(pdf_bytes)} bytes supera el máximo de 5 MB")
        # 3) Subida (PUT saliente) al bucket público (o disco local en dev)
        self.storage.put_pdf(record.object_key(), pdf_bytes)
        # 4) Registro de metadatos + timestamp + copia de retención
        self.repo.save_new(record, pdf_bytes)
        return record

    def create_many(self, data: DecaInput, tipos):
        """Emite varios documentos (deca/carta_porte/cmr) desde los MISMOS datos.
        Se meten los datos una vez y salen todos, cada uno con su UUID/URL/QR.
        Las firmas no se aplican al DeCA (no las necesita)."""
        recs = []
        for t in tipos:
            firmas = [] if t == "deca" else data.firmas
            recs.append(self.create(data.model_copy(update={"tipo_documento": t, "firmas": firmas})))
        return recs

    def modify(self, uuid: str):
        return self.repo.register_modification(uuid)

    def get(self, uuid: str):
        return self.repo.get(uuid)
