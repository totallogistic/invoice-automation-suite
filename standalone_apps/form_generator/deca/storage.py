"""Almacenamiento del PDF público.

Producción: PUT saliente a un bucket S3-compatible (Hetzner Object Storage).
Sin credenciales configuradas cae a un store en disco local, para poder
ejecutar el prototipo end-to-end sin contratar nada.

En ambos casos la URL pública es DETERMINISTA (base + clave), de modo que se
conoce ANTES de subir → se puede incrustar en el QR y luego subir el PDF.
"""

from __future__ import annotations

import os
import pathlib
from typing import Optional


class Storage:
    def __init__(self) -> None:
        self.endpoint = os.getenv("DECA_S3_ENDPOINT")
        self.bucket = os.getenv("DECA_S3_BUCKET", "deca-prod")
        self.key = os.getenv("DECA_S3_KEY")
        self.secret = os.getenv("DECA_S3_SECRET")
        # Base pública. Por defecto, el endpoint nativo del bucket (TLS válido del
        # proveedor, sin tocar el DNS propio). Ej. Hetzner:
        #   https://<bucket>.<region>.your-objectstorage.com
        self.public_base = os.getenv("DECA_PUBLIC_BASEURL", "").rstrip("/")
        self.local_dir = pathlib.Path(os.getenv("DECA_LOCAL_STORE", "_deca_store"))
        self._s3 = None
        if self.endpoint and self.key and self.secret:
            import boto3  # import perezoso
            from botocore.config import Config

            self._s3 = boto3.client(
                "s3", endpoint_url=self.endpoint,
                aws_access_key_id=self.key, aws_secret_access_key=self.secret,
                # path-style: obligatorio para MinIO/endpoint por IP, y compatible
                # con Hetzner (evita que boto3 intente el virtual-host bucket.<host>).
                config=Config(s3={"addressing_style": "path"}),
            )
            if not self.public_base:
                self.public_base = f"{self.endpoint.rstrip('/')}/{self.bucket}"

    @property
    def mode(self) -> str:
        return "s3" if self._s3 else "local"

    def public_url(self, object_key: str) -> str:
        base = self.public_base or "http://localhost:8202"  # dev: lo sirve la API
        return f"{base}/{object_key}"

    def put_pdf(self, object_key: str, pdf_bytes: bytes) -> str:
        """Sube el PDF y devuelve su URL pública (ya conocida de antemano)."""
        if self._s3:
            self._s3.put_object(
                Bucket=self.bucket, Key=object_key, Body=pdf_bytes,
                ContentType="application/pdf",
            )
        else:
            dest = self.local_dir / object_key
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(pdf_bytes)
        return self.public_url(object_key)

    def delete(self, object_key: str) -> None:
        """Desactiva la descarga pública (a `fin_de_servicio + 7 días`)."""
        if self._s3:
            self._s3.delete_object(Bucket=self.bucket, Key=object_key)
        else:
            p = self.local_dir / object_key
            if p.exists():
                p.unlink()

    def read_local(self, object_key: str) -> Optional[bytes]:
        """Solo modo local/dev: sirve el PDF desde disco (la API hace de bucket)."""
        p = self.local_dir / object_key
        return p.read_bytes() if p.exists() else None
