"""Modelo de datos de los documentos de transporte.

Superconjunto único que sirve para los tres documentos, con `tipo_documento`:
  - "deca"        → Documento Electrónico de Control Administrativo (art. 6
                    Orden FOM/2861/2012, Resolución 5-jun-2026 / BOE-A-2026-12784).
  - "carta_porte" → Carta de porte nacional (documento de control + campos
                    contractuales: portes, pago, valor declarado, aduana...).
  - "cmr"         → CMR internacional (Convenio de Ginebra 1956): añade
                    destinatario, reembolso, seguro, transportistas sucesivos...

El DeCA es el subconjunto; carta de porte y CMR lo contienen (modelo anidado,
alineado con el conjunto de datos eFTI). Los campos extra son opcionales: en un
DeCA puro quedan vacíos.
"""

from __future__ import annotations

import datetime as dt
import uuid as uuidlib
from typing import List, Optional

from pydantic import BaseModel, Field


class Cargador(BaseModel):
    nombre: str = Field(..., description="Nombre o razón social del cargador / remitente")
    nif: str = Field(..., description="NIF")
    domicilio: str = Field(..., description="Domicilio")


class Transportista(BaseModel):
    nombre: str = Field(..., description="Nombre o razón social del transportista efectivo")
    nif: str = Field(..., description="NIF")


class Destinatario(BaseModel):
    nombre: str = Field(..., description="Nombre o razón social del destinatario")
    nif: Optional[str] = None
    domicilio: Optional[str] = None


class Mercancia(BaseModel):
    naturaleza: str = Field(..., description="Naturaleza de la mercancía")
    peso_kg: float = Field(..., ge=0, description="Peso en kg")
    bultos: Optional[str] = Field(None, description="Nº de bultos y marcas")
    embalaje: Optional[str] = Field(None, description="Tipo de embalaje")


class Firma(BaseModel):
    """Firma electrónica simple (NO cualificada): imagen del trazo + sello propio."""
    rol: str = Field(..., description="cargador | transportista | destinatario")
    nombre: str = Field(..., description="Nombre de quien firma")
    firma_png: Optional[str] = Field(None, description="Imagen del trazo en base64 (o data URI)")
    lugar: Optional[str] = None
    firmado_en: Optional[dt.datetime] = None


class DecaInput(BaseModel):
    """Datos de entrada del documento de transporte (superconjunto DeCA/CdP/CMR)."""

    tipo_documento: str = Field("deca", description="deca | carta_porte | cmr")

    # ── Núcleo compartido (art. 6 FOM/2861/2012) ──
    cargador_contractual: Cargador
    transportista_efectivo: Transportista
    origen: str
    destino: str
    mercancia: Mercancia
    fecha_transporte: dt.date
    matricula_tractora: str
    matricula_remolque: Optional[str] = None
    autorizaciones_especiales: Optional[str] = None
    observaciones: Optional[str] = None
    servicio_inicio: Optional[dt.datetime] = None
    servicio_fin: Optional[dt.datetime] = None

    # ── Añade Carta de porte / CMR ──
    destinatario: Optional[Destinatario] = None
    portes: Optional[str] = Field(None, description="Pagados / Debidos")
    condiciones_pago: Optional[str] = None
    valor_declarado: Optional[str] = None
    instrucciones_aduana: Optional[str] = None

    # ── Añade CMR internacional ──
    lugar_carga: Optional[str] = Field(None, description="Lugar y fecha de toma de la mercancía")
    lugar_entrega: Optional[str] = None
    reembolso: Optional[str] = None
    instrucciones_seguro: Optional[str] = None
    prohibicion_trasbordo: Optional[bool] = None
    plazo_entrega: Optional[str] = None
    documentos_anexos: Optional[str] = None
    transportistas_sucesivos: Optional[str] = None

    # ── Firma simple propia (opcional; no cualificada) ──
    firmas: List[Firma] = Field(default_factory=list)


class DecaRecord(BaseModel):
    """Documento ya generado: datos + identidad + URL pública + timestamps."""

    uuid: str = Field(default_factory=lambda: uuidlib.uuid4().hex)
    datos: DecaInput
    url_publica: Optional[str] = None
    creado_en: Optional[dt.datetime] = None
    modificado_en: List[dt.datetime] = Field(default_factory=list)
    url_activa_hasta: Optional[dt.datetime] = None
    estado: str = "creado"

    def object_key(self) -> str:
        return f"d/{self.uuid}.pdf"
