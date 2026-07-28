"""Repositorio de metadatos y timestamps (SQLite).

La Resolución (Apartado Primero) exige registrar la fecha/hora de creación del
fichero y de cada modificación. Aquí se guardan los metadatos; el PDF vive en el
bucket público + una copia de retención (>= 1 año) que en producción se conserva
en on-prem.
"""

from __future__ import annotations

import datetime as dt
import json
import os
import pathlib
import sqlite3
from typing import Optional

from .models import DecaRecord

DB_PATH = os.getenv("DECA_DB", "deca.db")
RETENTION_DIR = pathlib.Path(os.getenv("DECA_RETENTION_DIR", "_deca_retencion"))


def _utcnow() -> dt.datetime:
    return dt.datetime.now(dt.timezone.utc)


class Repository:
    def __init__(self, db_path: str = DB_PATH) -> None:
        # Crea la carpeta de la BD si no existe (sqlite no crea el directorio padre).
        parent = pathlib.Path(db_path).parent
        if str(parent) not in ("", "."):
            parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(db_path)
        self.conn.execute(
            """CREATE TABLE IF NOT EXISTS deca (
                uuid TEXT PRIMARY KEY,
                creado_en TEXT NOT NULL,
                modificado_en TEXT NOT NULL,
                servicio_fin TEXT,
                url_publica TEXT,
                url_activa_hasta TEXT,
                estado TEXT NOT NULL,
                payload_json TEXT NOT NULL
            )"""
        )
        self.conn.commit()

    def save_new(self, record: DecaRecord, pdf_bytes: bytes) -> None:
        if record.creado_en is None:
            record.creado_en = _utcnow()
        # Ventana de la URL pública: fin de servicio + 7 días (o creación + 7 si no hay fin).
        fin = record.datos.servicio_fin or record.creado_en
        record.url_activa_hasta = fin + dt.timedelta(days=int(os.getenv("DECA_URL_TTL_DAYS", "7")))
        # Copia de retención interna (>= 1 año en producción). Se añade el tipo de
        # documento como sufijo para distinguir DeCA / carta de porte en la misma
        # carpeta (p. ej. <uuid>_deca.pdf / <uuid>_carta_porte.pdf). Ambos se
        # conservan en local; al bucket público solo va el DeCA (ver service/storage).
        RETENTION_DIR.mkdir(parents=True, exist_ok=True)
        _tipo = "".join(c for c in (record.datos.tipo_documento or "doc") if c.isalnum() or c == "_") or "doc"
        (RETENTION_DIR / f"{record.uuid}_{_tipo}.pdf").write_bytes(pdf_bytes)
        self.conn.execute(
            "INSERT INTO deca VALUES (?,?,?,?,?,?,?,?)",
            (record.uuid, record.creado_en.isoformat(), json.dumps([]),
             record.datos.servicio_fin.isoformat() if record.datos.servicio_fin else None,
             record.url_publica, record.url_activa_hasta.isoformat(),
             record.estado, record.datos.model_dump_json()),
        )
        self.conn.commit()

    def register_modification(self, uuid: str) -> Optional[dt.datetime]:
        row = self.conn.execute(
            "SELECT modificado_en FROM deca WHERE uuid=?", (uuid,)
        ).fetchone()
        if not row:
            return None
        mods = json.loads(row[0])
        ts = _utcnow()
        mods.append(ts.isoformat())
        self.conn.execute(
            "UPDATE deca SET modificado_en=?, estado='modificado' WHERE uuid=?",
            (json.dumps(mods), uuid),
        )
        self.conn.commit()
        return ts

    def get(self, uuid: str) -> Optional[dict]:
        cols = ["uuid", "creado_en", "modificado_en", "servicio_fin",
                "url_publica", "url_activa_hasta", "estado", "payload_json"]
        row = self.conn.execute(
            f"SELECT {','.join(cols)} FROM deca WHERE uuid=?", (uuid,)
        ).fetchone()
        if not row:
            return None
        rec = dict(zip(cols, row))
        rec["modificado_en"] = json.loads(rec["modificado_en"])
        rec["payload"] = json.loads(rec.pop("payload_json"))
        return rec

    def stats(self) -> dict:
        total = self.conn.execute("SELECT COUNT(*) FROM deca").fetchone()[0]
        return {"total": total}
