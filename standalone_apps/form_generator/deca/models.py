"""Modelo de datos del DeCA (mercancías).

Campos obligatorios del art. 6 de la Orden FOM/2861/2012, tal y como los exige
la Resolución de 5-jun-2026 (BOE-A-2026-12784). La norma obliga a identificar
de forma expresa y diferenciada al cargador contractual y al transportista
efectivo (Apartado Octavo).
"""

from __future__ import annotations

import datetime as dt
import uuid as uuidlib
from typing import List, Optional

from pydantic import BaseModel, Field


class Cargador(BaseModel):
    nombre: str = Field(..., description="Nombre o razón social del cargador contractual")
    nif: str = Field(..., description="NIF del cargador contractual")
    domicilio: str = Field(..., description="Domicilio del cargador contractual")


class Transportista(BaseModel):
    nombre: str = Field(..., description="Nombre o razón social del transportista efectivo")
    nif: str = Field(..., description="NIF del transportista efectivo")


class Mercancia(BaseModel):
    naturaleza: str = Field(..., description="Naturaleza de la mercancía")
    peso_kg: float = Field(..., ge=0, description="Peso de la mercancía en kg")


class DecaInput(BaseModel):
    """Datos de entrada de un DeCA (lo que introduce el usuario / se pre-rellena)."""

    cargador_contractual: Cargador
    transportista_efectivo: Transportista
    origen: str = Field(..., description="Lugar de origen del transporte")
    destino: str = Field(..., description="Lugar de destino del transporte")
    mercancia: Mercancia
    fecha_transporte: dt.date = Field(..., description="Fecha del transporte")
    matricula_tractora: str = Field(..., description="Matrícula de la tractora / vehículo")
    matricula_remolque: Optional[str] = Field(
        None, description="Matrícula del remolque (si es conjunto articulado)"
    )
    autorizaciones_especiales: Optional[str] = Field(
        None, description="Autorizaciones especiales, si se requieren"
    )
    observaciones: Optional[str] = Field(None, description="Observaciones / reservas")
    # Ventana del servicio: define cuándo se puede desactivar la URL pública (fin + 7 días).
    servicio_inicio: Optional[dt.datetime] = None
    servicio_fin: Optional[dt.datetime] = None


class DecaRecord(BaseModel):
    """DeCA ya generado: datos + identidad + URL pública + timestamps."""

    uuid: str = Field(default_factory=lambda: uuidlib.uuid4().hex)
    datos: DecaInput
    url_publica: Optional[str] = None
    creado_en: Optional[dt.datetime] = None
    modificado_en: List[dt.datetime] = Field(default_factory=list)
    url_activa_hasta: Optional[dt.datetime] = None
    estado: str = "creado"

    def object_key(self) -> str:
        """Clave del objeto en el bucket (y sufijo de la URL pública)."""
        return f"d/{self.uuid}.pdf"
