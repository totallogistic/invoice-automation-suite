"""Perfil de formularios habilitados por despliegue (allow-list explícita).

Permite desplegar el form_generator sirviendo SOLO un subconjunto de formularios
(p. ej. únicamente el DeCA + su importación por CSV) sin tocar el código: se
controla por la variable de entorno `FORMS_ENABLED`.

  FORMS_ENABLED = lista de schemas separada por comas, p. ej.
                  "documento-transporte"  ó  "documento-transporte,carta-porte".
  Vacío / no definido  →  TODOS los formularios habilitados (comportamiento por
                          defecto; Total Logistic no cambia en nada).

En modo restringido (FORMS_ENABLED definido):
  - la landing solo lista los schemas habilitados (filtro en `index()`),
  - un middleware en app.py devuelve 404 a CUALQUIER ruta de formulario no
    habilitada (no basta con ocultar la ficha: hay que cerrar la ruta),
  - los routers pesados (p. ej. extintores) solo se montan si procede.

Este módulo contiene solo lógica pura (sin dependencias de FastAPI) para poder
testearlo de forma aislada; el middleware y el filtrado viven en app.py.
"""

from __future__ import annotations

import os
import re


def enabled_set(raw: str | None = None):
    """Conjunto de schemas habilitados, o None si están TODOS (sin restricción).
    Acepta separadores ',' o ';'."""
    raw = (os.getenv("FORMS_ENABLED", "") if raw is None else (raw or "")).strip()
    if not raw:
        return None
    return {s.strip() for s in raw.replace(";", ",").split(",") if s.strip()}


def form_enabled(name: str, enabled) -> bool:
    """¿Está habilitado este schema concreto?"""
    return enabled is None or name in enabled


def feature_enabled(feature: str, enabled) -> bool:
    """¿Está habilitada una 'familia'/funcionalidad con rutas propias cableadas?
    - 'deca'  → si está el formulario unificado `documento-transporte` (o `deca`).
    - resto   → si algún schema habilitado coincide o empieza por el prefijo
                (cubre familias por sede: revision-estanterias-alg/mlg, etc.)."""
    if enabled is None:
        return True
    if feature == "deca":
        return "documento-transporte" in enabled or "deca" in enabled
    return any(s == feature or s.startswith(feature) for s in enabled)


# Rutas "custom" (no siguen el patrón /form/{schema}) → feature a la que pertenecen.
_ROUTE_FEATURE = (
    ("/form-custom/estanterias", "revision-estanterias"),
    ("/api/save-estanterias-completo", "revision-estanterias"),
    ("/api/save-maquinaria-mlg", "mantenimiento-maquinas-mlg"),
    ("/api/save-liquidacion", "liquidacion-gastos"),
    ("/api/save-entrega-epis", "entrega-epis"),
    ("/api/save-hoja-control-expedientes", "hoja-control-expedientes"),
    ("/api/hoja-control", "hoja-control-expedientes"),
    ("/api/precintos", "precintos"),
    ("/precintos", "precintos"),
    ("/api/bl", "bl"),
    ("/bl", "bl"),
    ("/api/save-deca", "deca"),
    ("/api/deca", "deca"),
    ("/deca/", "deca"),
)

# Rutas de infraestructura SIEMPRE permitidas (no son formularios).
_ALWAYS_OK = ("/health", "/api/forms/status", "/politica-privacidad",
              "/static", "/favicon", "/download/", "/download-storage/")

# Rutas con el schema en el path (más largas primero para el match correcto).
_SCHEMA_RE = re.compile(
    r"^/(?:form|api/schema|api/validate|api/save-to-excel|api/save-and-email|"
    r"api/save|api/submit)/([^/]+)"
)
_VERSION_RE = re.compile(r"^/api/form/([^/]+)/version")


def path_allowed(path: str, enabled) -> bool:
    """¿Puede atenderse esta ruta con el perfil actual? True siempre que no haya
    restricción (enabled is None)."""
    if enabled is None:
        return True
    if path == "/" or path.startswith(_ALWAYS_OK):
        return True
    m = _SCHEMA_RE.match(path) or _VERSION_RE.match(path)
    if m:
        return form_enabled(m.group(1), enabled)
    for prefix, feat in _ROUTE_FEATURE:
        if path.startswith(prefix):
            return feature_enabled(feat, enabled)
    return False  # en modo restringido, lo no reconocido se deniega
