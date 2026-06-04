"""
Daily sync job — 07:00 via APScheduler.
Post-sync: actualiza Excel local, sincroniza Google Sheets, envía email.
"""

import os
import logging
from datetime import datetime, timezone

from bank_fetcher import EnableBankingClient
from db import get_all_banks, save_balances

logger = logging.getLogger(__name__)


def _get_client() -> EnableBankingClient:
    app_id   = os.environ["ENABLE_BANKING_APP_ID"]
    key_path = os.environ.get("ENABLE_BANKING_KEY_PATH", "keys/private.key")
    return EnableBankingClient(app_id, open(key_path).read())


async def sync_all_banks() -> dict:
    client    = _get_client()
    banks     = await get_all_banks()
    connected = [b for b in banks if b["status"] == "connected" and b["session_id"]]
    results   = {"synced": [], "failed": [], "skipped": []}

    if not connected:
        logger.warning("No connected banks to sync.")
        return results

    for bank in connected:
        try:
            import aiosqlite
            from db import DB_PATH
            async with aiosqlite.connect(DB_PATH) as conn:
                conn.row_factory = aiosqlite.Row
                async with conn.execute("SELECT * FROM accounts WHERE bank_id=?", (bank["id"],)) as cur:
                    accounts = [dict(r) for r in await cur.fetchall()]

            if not accounts:
                logger.warning("Bank %s has no accounts, skipping.", bank["name"])
                results["skipped"].append(bank["name"])
                continue

            accounts_payload = [
                {"id": a["id"], "account_id": {"iban": a["iban"]}, "name": a["name"]}
                for a in accounts
            ]
            balances = await client.fetch_all_balances(accounts_payload, bank["session_id"])
            await save_balances(bank["id"], balances)
            results["synced"].append({"bank": bank["name"], "accounts": len(balances)})
            logger.info("Synced %s: %d accounts", bank["name"], len(balances))

        except Exception as exc:
            logger.error("Sync failed for %s: %s", bank["name"], exc)
            results["failed"].append({"bank": bank["name"], "error": str(exc)})

    await _post_sync_exports(results)
    return results


async def _post_sync_exports(results: dict):
    from db import get_latest_balances, get_history_per_bank, get_daily_totals
    latest  = await get_latest_balances()
    history = await get_history_per_bank(days=90)
    daily   = await get_daily_totals(days=90)

    if not latest:
        return

    # 1. saldos_actuales.yaml — nombre fijo, puente para Situación Financiera
    try:
        import yaml
        from pathlib import Path
        yaml_data = {
            "fecha": datetime.now().strftime("%Y-%m-%d"),
            "generado_por": "sync_automatico",
            "saldos": [
                {
                    "iban": b["iban"],
                    "banco": b["bank_name"],
                    "cuenta": b.get("account_name", ""),
                    "dispuesto": round(b["amount"], 2),
                    "moneda": b.get("currency", "EUR"),
                }
                for b in latest if b.get("iban")
            ]
        }
        yaml_path = Path(__file__).parent / "data" / "saldos_actuales.yaml"
        yaml_path.write_text(yaml.dump(yaml_data, allow_unicode=True, default_flow_style=False))
        logger.info("saldos_actuales.yaml actualizado (%d cuentas)", len(yaml_data["saldos"]))
    except Exception as exc:
        logger.warning("saldos_actuales.yaml failed: %s", exc)

    # 2. Excel saldos diario + histórico
    try:
        from excel_export import generate_excel
        today_str = datetime.now().strftime("%Y%m%d")
        await generate_excel(latest, history, daily, filename=f"saldos_{today_str}.xlsx")
        await generate_excel(latest, history, daily, filename="saldos_historico.xlsx")
        logger.info("Excel saldos diario y histórico actualizados")
    except Exception as exc:
        logger.warning("Excel saldos failed: %s", exc)

    # 3. Excel bancos por empresa (replica modelo manual)
    try:
        from bancos_export import generate_bancos
        today_str = datetime.now().strftime("%Y%m%d")
        await generate_bancos(history=history, filename=f"bancos_{today_str}.xlsx")
        logger.info("Excel bancos actualizado: bancos_%s.xlsx", today_str)
    except Exception as exc:
        logger.warning("Excel bancos failed: %s", exc)

    # 3. Google Sheets (si configurado en .env)
    creds_path = os.environ.get("GOOGLE_SHEETS_CREDENTIALS", "")
    sheet_id   = os.environ.get("GOOGLE_SHEETS_ID", "")
    if creds_path and sheet_id:
        try:
            from sheets_export import sync_to_sheets
            await sync_to_sheets(creds_path, sheet_id, latest, history, daily)
            logger.info("Google Sheets sincronizado")
        except Exception as exc:
            logger.error("Google Sheets sync failed: %s", exc)

    # 3. Email resumen
    await _send_summary_email(results, latest)


async def _send_summary_email(results: dict, latest: list[dict]):
    try:
        import sys, pathlib
        sys.path.insert(0, str(pathlib.Path(__file__).parents[3] / "libs"))
        from iasuite_common.email import send_email  # type: ignore

        total = sum(r["amount"] for r in latest)
        lines = [f"<h2>Saldos {datetime.now().strftime('%d/%m/%Y')}</h2>"]
        lines.append(f"<p><strong>Total: {total:,.2f} EUR</strong></p><hr>")
        lines.append("<table cellpadding='6' style='font-family:monospace'>")
        lines.append("<tr><th>Banco</th><th>IBAN</th><th>Saldo</th></tr>")
        for row in latest:
            iban = f"****{row['iban'][-4:]}" if row.get("iban") else "—"
            lines.append(f"<tr><td>{row['bank_name']}</td><td>{iban}</td><td align='right'>{row['amount']:,.2f} {row['currency']}</td></tr>")
        lines.append("</table>")
        if results["failed"]:
            lines.append("<p style='color:red'>⚠️ Errores:</p><ul>")
            for f in results["failed"]:
                lines.append(f"<li>{f['bank']}: {f['error']}</li>")
            lines.append("</ul>")

        recipient = os.environ.get("REPORT_EMAIL", "")
        if recipient:
            send_email(
                to=recipient,
                subject=f"💰 Saldos {datetime.now().strftime('%d/%m/%Y')} — {total:,.0f} EUR",
                body_html="\n".join(lines),
            )
    except ImportError:
        logger.warning("iasuite_common no encontrado — sin email")
    except Exception as exc:
        logger.error("Email failed: %s", exc)