"""
Generador de Situación Financiera.

Flujo completo:
  1. Carga el template limpio (solo estructura + fórmulas)
  2. Aplica situacion_config.yaml:
       - Euribor actual → B19, B20, B21
       - Por fila: producto (C), euribor_tipo (F), diferencial (G), importe (J)
  3. Escribe Dispuesto (K) desde API bancaria — matching por IBAN
  4. Actualiza fecha en F1
  5. Guarda → data/situacion_financiera.xlsx
     (el template nunca se modifica)

El resto de columnas (L, M, N, I, totales, tabla R) se recalculan
solas en Excel al abrir el fichero.
"""

import logging
import yaml
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

TEMPLATE_PATH = Path(__file__).parent / "data" / "SITUACION_FINANCIERA_TEMPLATE.xlsx"
OUTPUT_PATH   = Path(__file__).parent / "data" / "situacion_financiera.xlsx"
CONFIG_PATH   = Path(__file__).parent / "situacion_config.yaml"

# Columnas
COL_PRODUCTO    = 3   # C — producto / IBAN líneas
COL_EURIBOR_LBL = 6  # F — etiqueta tipo Euribor
COL_DIFERENCIAL = 7   # G — diferencial sobre Euribor
COL_IMPORTE     = 10  # J — límite del crédito
COL_DISPUESTO   = 11  # K — dispuesto ← API

# Celdas Euribor
EURIBOR_CELDAS = {
    "Euribor 3":  "B19",
    "Euribor 6":  "B20",
    "Euribor 12": "B21",
}


def _load_config() -> dict:
    if not CONFIG_PATH.exists():
        logger.warning("situacion_config.yaml no encontrado")
        return {}
    data = yaml.safe_load(CONFIG_PATH.read_text())
    return data if isinstance(data, dict) else {}


def _iban(value) -> str:
    return str(value).replace(" ", "").upper() if value else ""


async def generate_situacion(latest_balances: list[dict]) -> Path:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Template no encontrado: {TEMPLATE_PATH}\n"
            "Copia SITUACION_FINANCIERA_TEMPLATE.xlsx a la carpeta data/."
        )

    wb     = load_workbook(TEMPLATE_PATH)
    ws     = wb.active
    config = _load_config()

    # ── 1. Fecha ───────────────────────────────────────────────────────
    ws["F1"] = f"SITUACIÓN FINANCIERA {datetime.now().strftime('%d/%m/%y')}"

    # ── 2. Euribor (B19, B20, B21) ────────────────────────────────────
    for nombre, celda in EURIBOR_CELDAS.items():
        valor = config.get("euribor", {}).get(nombre)
        if valor is not None:
            ws[celda] = valor
            logger.info("Euribor %s → %s = %.5f", nombre, celda, valor)

    # ── 3. Índice IBAN → balance (API) ────────────────────────────────
    iban_index: dict[str, float] = {
        _iban(b["iban"]): b["amount"]
        for b in latest_balances
        if b.get("iban")
    }

    # ── 4. Aplica config por fila ──────────────────────────────────────
    cuentas   = config.get("cuentas") or []
    cfg_map   = {c["fila"]: c for c in cuentas}
    actualizados = 0
    sin_datos    = []

    for fila in range(4, 17):
        tipo = ws.cell(row=fila, column=5).value  # E
        if not tipo:
            continue

        cfg = cfg_map.get(fila, {})

        # Producto (C) — solo para líneas/anticipos (pólizas mantienen IBAN en C)
        if cfg.get("producto") is not None:
            ws.cell(row=fila, column=COL_PRODUCTO).value = cfg["producto"]

        # Tipo Euribor (F) — etiqueta informativa
        if cfg.get("euribor_tipo"):
            ws.cell(row=fila, column=COL_EURIBOR_LBL).value = cfg["euribor_tipo"]

        # Diferencial (G)
        if cfg.get("diferencial") is not None:
            ws.cell(row=fila, column=COL_DIFERENCIAL).value = cfg["diferencial"]

        # Importe (J)
        if cfg.get("importe") is not None:
            ws.cell(row=fila, column=COL_IMPORTE).value = cfg["importe"]

        # Dispuesto (K) — desde API por IBAN
        # IBAN: primero config, luego columna C del template (pólizas)
        iban_cfg  = _iban(cfg.get("iban"))
        iban_tmpl = _iban(ws.cell(row=fila, column=COL_PRODUCTO).value)
        iban      = iban_cfg or iban_tmpl

        if not iban:
            sin_datos.append(f"  F{fila} ({cfg.get('banco','?')}): sin IBAN — Dispuesto queda vacío")
            continue

        amount = iban_index.get(iban)
        if amount is not None:
            cell = ws.cell(row=fila, column=COL_DISPUESTO)
            cell.value         = round(amount, 2)
            cell.number_format = '#,##0.00'
            actualizados += 1
            logger.info("F%d %s → K=%.2f €", fila, cfg.get("banco", ""), amount)
        else:
            sin_datos.append(
                f"  F{fila} ({cfg.get('banco','?')}) IBAN {iban[:16]}...: "
                "no encontrado en API (banco no conectado aún)"
            )

    if sin_datos:
        logger.info("Filas sin Dispuesto:\n%s", "\n".join(sin_datos))
    logger.info("Dispuesto actualizado: %d/%d filas con IBAN", actualizados, len(cuentas))

    # ── 5. Guarda ──────────────────────────────────────────────────────
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUTPUT_PATH)
    logger.info("Guardado: %s", OUTPUT_PATH)
    return OUTPUT_PATH