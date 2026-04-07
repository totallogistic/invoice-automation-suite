"""
Generador de Situación Financiera.

Estructura de directorios:
  config/            → situacion_config.yaml (Samba compartido para edición)
  data/              → saldos_actuales.yaml + ficheros generados (Samba lectura)
  templates_excel/   → SITUACION_FINANCIERA_TEMPLATE.xlsx (protegido, sin Samba)

Lógica de Dispuesto (K) por prioridad:
  1. saldos_actuales.yaml (sync bancario automático) — si IBAN coincide
  2. campo 'dispuesto' en situacion_config.yaml — para líneas sin API
"""

import logging
import yaml
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

BASE_DIR      = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "templates_excel" / "SITUACION_FINANCIERA_TEMPLATE.xlsx"
OUTPUT_PATH   = BASE_DIR / "data" / "situacion_financiera.xlsx"
CONFIG_PATH   = BASE_DIR / "config" / "situacion_config.yaml"
SALDOS_PATH   = BASE_DIR / "data" / "saldos_actuales.yaml"

# Columnas del template
COL = {
    "banco":       1,   # A
    "cta_aux":     2,   # B
    "producto":    3,   # C
    "or":          4,   # D
    "tipo":        5,   # E
    "euribor_lbl": 6,   # F
    "diferencial": 7,   # G
    "nd":          8,   # H
    "importe":     10,  # J
    "dispuesto":   11,  # K
    "reservado":   18,  # R
    "label_tabla": 19,  # S
}


def _load_yaml(path: Path) -> dict:
    if not path.exists():
        return {}
    data = yaml.safe_load(path.read_text())
    return data if isinstance(data, dict) else {}


def _iban(value) -> str:
    return str(value).replace(" ", "").upper() if value else ""


async def generate_situacion() -> Path:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Template no encontrado: {TEMPLATE_PATH}\n"
            "Copia SITUACION_FINANCIERA_TEMPLATE.xlsx a la carpeta templates_excel/."
        )
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Config no encontrado: {CONFIG_PATH}\n"
            "Asegúrate de que situacion_config.yaml está en la carpeta config/."
        )

    config = _load_yaml(CONFIG_PATH)
    saldos = _load_yaml(SALDOS_PATH)

    if not config:
        raise ValueError("situacion_config.yaml vacío.")

    # Índice IBAN → dispuesto desde sync bancario (prioridad 1)
    iban_index: dict[str, float] = {}
    for s in saldos.get("saldos", []):
        iban_key = _iban(s.get("iban", ""))
        if iban_key:
            iban_index[iban_key] = float(s.get("dispuesto", 0))

    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active

    # ── 1. Fecha ──────────────────────────────────────────────────────
    ws["F1"] = f"SITUACIÓN FINANCIERA {datetime.now().strftime('%d/%m/%y')}"

    # ── 2. Euribor ────────────────────────────────────────────────────
    eur = config.get("euribor", {})
    if eur.get("Euribor 3"):  ws["B19"] = eur["Euribor 3"]
    if eur.get("Euribor 6"):  ws["B20"] = eur["Euribor 6"]
    if eur.get("Euribor 12"): ws["B21"] = eur["Euribor 12"]

    # ── 3. R8 ─────────────────────────────────────────────────────────
    if config.get("disponible_importacion"):
        ws["R8"] = config["disponible_importacion"]

    # ── 4. Anticipos ──────────────────────────────────────────────────
    for i, ant in enumerate(config.get("anticipos", []), start=23):
        ws.cell(row=i, column=1).value = ant.get("banco", "")
        ws.cell(row=i, column=2).value = ant.get("numero")

    # ── 5. Cuentas ────────────────────────────────────────────────────
    cuentas = config.get("cuentas") or []
    actualizados_api    = 0
    actualizados_manual = 0
    sin_dispuesto       = []

    for cfg in cuentas:
        fila = cfg.get("fila")
        if not fila:
            continue

        # Escribe campos del config
        ws.cell(row=fila, column=COL["banco"]).value    = cfg.get("banco")
        ws.cell(row=fila, column=COL["cta_aux"]).value  = cfg.get("cta_auxiliar")
        ws.cell(row=fila, column=COL["or"]).value       = cfg.get("or")
        ws.cell(row=fila, column=COL["tipo"]).value     = cfg.get("tipo")
        ws.cell(row=fila, column=COL["producto"]).value = cfg.get("producto")

        if cfg.get("euribor_tipo"):
            ws.cell(row=fila, column=COL["euribor_lbl"]).value = cfg["euribor_tipo"]
        if cfg.get("diferencial") is not None:
            ws.cell(row=fila, column=COL["diferencial"]).value = cfg["diferencial"]
        if cfg.get("nd"):
            ws.cell(row=fila, column=COL["nd"]).value = cfg["nd"]
        if cfg.get("importe") is not None:
            cell = ws.cell(row=fila, column=COL["importe"])
            cell.value        = cfg["importe"]
            cell.number_format = '#,##0'
        if cfg.get("reservado") is not None:
            cell = ws.cell(row=fila, column=COL["reservado"])
            cell.value        = cfg["reservado"]
            cell.number_format = '#,##0'
            ws.cell(row=fila, column=COL["label_tabla"]).value = cfg.get("banco")

        # ── Dispuesto (K) — prioridad: API > config manual ────────────
        dispuesto = None

        # Prioridad 1: API (por IBAN)
        iban = _iban(cfg.get("iban"))
        if iban:
            api_val = iban_index.get(iban)
            if api_val is not None:
                dispuesto = api_val
                actualizados_api += 1
                logger.info("F%d %s → K=%.2f € [API]", fila, cfg.get("banco", ""), dispuesto)

        # Prioridad 2: config manual (campo dispuesto)
        if dispuesto is None and cfg.get("dispuesto") is not None:
            dispuesto = float(cfg["dispuesto"])
            actualizados_manual += 1
            logger.info("F%d %s → K=%.2f € [manual config]", fila, cfg.get("banco", ""), dispuesto)

        if dispuesto is not None:
            cell = ws.cell(row=fila, column=COL["dispuesto"])
            cell.value        = round(dispuesto, 2)
            cell.number_format = '#,##0.00'
        else:
            sin_dispuesto.append(f"  F{fila} {cfg.get('banco')} — sin dispuesto")

    logger.info(
        "Dispuesto: %d via API, %d manual config, %d sin datos",
        actualizados_api, actualizados_manual, len(sin_dispuesto)
    )
    if sin_dispuesto:
        logger.debug("Sin dispuesto:\n%s", "\n".join(sin_dispuesto))

    # ── 6. Guarda ─────────────────────────────────────────────────────
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    logger.info("Guardado: %s", OUTPUT_PATH)
    return OUTPUT_PATH