"""
Google Sheets sync — escribe saldos e histórico en una hoja compartida.

Setup necesario (una sola vez):
  1. Google Cloud Console → crear Service Account → descargar JSON de credenciales
  2. Compartir la Google Sheet con el email del Service Account (editor)
  3. Añadir al .env:
       GOOGLE_SHEETS_CREDENTIALS=/ruta/al/credentials.json
       GOOGLE_SHEETS_ID=1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms

Dependencias: gspread (ya en requirements.txt)
"""

import logging
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]


def _get_sheet(credentials_path: str, spreadsheet_id: str):
    creds  = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(spreadsheet_id)


def _ensure_worksheet(spreadsheet, title: str, rows: int = 500, cols: int = 20):
    try:
        ws = spreadsheet.worksheet(title)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)
    return ws


async def sync_to_sheets(
    credentials_path: str,
    spreadsheet_id: str,
    latest: list[dict],
    history: list[dict],
    daily_totals: list[dict],
):
    try:
        spreadsheet = _get_sheet(credentials_path, spreadsheet_id)

        # ------------------------------------------------------------------
        # Hoja 1: Saldos actuales
        # ------------------------------------------------------------------
        ws1 = _ensure_worksheet(spreadsheet, "Saldos actuales")

        headers = ["Banco", "IBAN", "Cuenta", "Saldo (EUR)", "Última sync", "Actualizado"]
        rows = [headers]

        total = 0.0
        for row in latest:
            iban = f"****{row['iban'][-4:]}" if row.get("iban") and len(row["iban"]) >= 4 else row.get("iban", "—")
            rows.append([
                row["bank_name"],
                iban,
                row.get("account_name", ""),
                round(row["amount"], 2),
                row["recorded_at"][:10] if row.get("recorded_at") else "—",
                datetime.now().strftime("%d/%m/%Y %H:%M"),
            ])
            total += row["amount"]

        # Fila total
        rows.append(["TOTAL", "", "", round(total, 2), "", ""])

        ws1.update(rows, value_input_option="USER_ENTERED")

        # Formato cabecera (negrita)
        ws1.format("A1:F1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.1, "green": 0.1, "blue": 0.18},
        })
        # Formato total
        total_row = len(latest) + 2
        ws1.format(f"A{total_row}:F{total_row}", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.91, "green": 0.94, "blue": 0.99},
        })
        # Ancho de columnas
        ws1.set_basic_filter()
        logger.info("Sheets: hoja 'Saldos actuales' actualizada (%d cuentas)", len(latest))

        # ------------------------------------------------------------------
        # Hoja 2: Histórico diario
        # ------------------------------------------------------------------
        ws2 = _ensure_worksheet(spreadsheet, "Histórico")

        bank_names = sorted(set(r["bank_name"] for r in history))
        all_days   = sorted(set(r["day"] for r in history))
        day_map    = {}
        for r in history:
            day_map.setdefault(r["day"], {})[r["bank_name"]] = r["total"]

        hist_headers = ["Fecha"] + bank_names + ["Total EUR"]
        hist_rows    = [hist_headers]

        for day in all_days:
            row_vals = [day]
            row_total = 0.0
            for bank in bank_names:
                val = day_map.get(day, {}).get(bank)
                row_vals.append(round(val, 2) if val is not None else "")
                if val:
                    row_total += val
            row_vals.append(round(row_total, 2))
            hist_rows.append(row_vals)

        ws2.update(hist_rows, value_input_option="USER_ENTERED")
        ws2.format("A1:Z1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.1, "green": 0.1, "blue": 0.18},
        })
        ws2.set_basic_filter()
        logger.info("Sheets: hoja 'Histórico' actualizada (%d días)", len(all_days))

        # ------------------------------------------------------------------
        # Hoja 3: Totales diarios (para gráfico rápido)
        # ------------------------------------------------------------------
        ws3 = _ensure_worksheet(spreadsheet, "Totales diarios")
        total_rows = [["Fecha", "Total EUR"]]
        for r in daily_totals:
            total_rows.append([r["day"], round(r["total"], 2)])
        ws3.update(total_rows, value_input_option="USER_ENTERED")
        ws3.format("A1:B1", {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.1, "green": 0.1, "blue": 0.18},
        })
        logger.info("Sheets: hoja 'Totales diarios' actualizada")

    except Exception as exc:
        logger.error("Google Sheets sync failed: %s", exc)
        raise