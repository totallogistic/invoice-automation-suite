"""Versión y changelog de la tool `deca`.

Sigue la convención del stack (ver claude/convenciones-version-y-health.md):
`unified_api` lee estas constantes por regex y las expone en
`GET /api/deca/version` y en el agregador `GET /api/versions`.
"""

SCRIPT_VERSION = "2026-07-16.v1"

SCRIPT_CHANGELOG = """
## 2026-07-16.v1
Prototipo inicial del Documento Electrónico de Control Administrativo (DeCA):
- Modelo de datos del art. 6 Orden FOM/2861/2012 (mercancías).
- Generación de PDF nativo (<=5 MB) con QR incrustado que codifica la URL
  determinista del documento.
- Subida (PUT) a bucket S3-compatible (Hetzner Object Storage), con fallback a
  disco local cuando no hay credenciales.
- Registro de metadatos y timestamps (creación / modificaciones) en SQLite.
- Caducidad de la URL pública a `fin_de_servicio + 7 días`; copia interna con
  retención >= 1 año.
"""
