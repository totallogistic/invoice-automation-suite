"""Orquestación del alta de un DeCA: datos -> URL -> PDF+QR -> subida -> registro.

Orden importante: la URL es DETERMINISTA (base + clave), así que se conoce ANTES
de renderizar → se incrusta en el QR → se sube el PDF a esa misma clave.
"""

from __future__ import annotations

import datetime as dt

from .models import DecaInput, DecaRecord
from .pdf import render_pdf
from .repository import Repository
from .storage import Storage

MAX_PDF_BYTES = 5 * 1024 * 1024  # 5 MB (límite de la norma)


class DecaService:
    def __init__(self) -> None:
        self.storage = Storage()
        self.repo = Repository()

    def create(self, data: DecaInput) -> DecaRecord:
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
