"""
Daily sync job.
Runs at 07:00 via APScheduler (started from app.py).
Fetches balances from all connected banks, stores them,
and sends a summary email via iasuite_common.email.
"""

import os
import logging
from datetime import datetime, timezone

from bank_fetcher import EnableBankingClient
from db import get_all_banks, save_balances

logger = logging.getLogger(__name__)

# ------------------------------------------------------------------
# Load client once (credentials from env)
# ------------------------------------------------------------------
def _get_client() -> EnableBankingClient:
    app_id = os.environ["ENABLE_BANKING_APP_ID"]
    key_path = os.environ.get("ENABLE_BANKING_KEY_PATH", "keys/private.pem")
    private_key = open(key_path).read()
    return EnableBankingClient(app_id, private_key)


# ------------------------------------------------------------------
# Core sync logic
# ------------------------------------------------------------------
async def sync_all_banks() -> dict:
    client = _get_client()
    banks = await get_all_banks()
    connected = [b for b in banks if b["status"] == "connected" and b["session_id"]]

    results = {"synced": [], "failed": [], "skipped": []}

    if not connected:
        logger.warning("No connected banks to sync.")
        return results

    for bank in connected:
        try:
            # Re-fetch accounts for this bank from the session
            # In real usage you'd store accounts list; here we re-query
            from db import DB_PATH
            import aiosqlite
            async with aiosqlite.connect(DB_PATH) as db_conn:
                db_conn.row_factory = aiosqlite.Row
                async with db_conn.execute(
                    "SELECT * FROM accounts WHERE bank_id=?", (bank["id"],)
                ) as cur:
                    accounts = [dict(r) for r in await cur.fetchall()]

            if not accounts:
                logger.warning("Bank %s has no accounts stored, skipping.", bank["name"])
                results["skipped"].append(bank["name"])
                continue

            # Normalise to the shape expected by fetch_all_balances
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

    # Send daily email summary
    await _send_summary_email(results)
    return results


# ------------------------------------------------------------------
# Email report via iasuite_common
# ------------------------------------------------------------------
async def _send_summary_email(results: dict):
    try:
        # Try to import iasuite_common from the parent libs directory
        import sys, pathlib
        sys.path.insert(0, str(pathlib.Path(__file__).parents[3] / "libs"))
        from iasuite_common.email import send_email  # type: ignore

        from db import get_latest_balances, get_daily_totals
        latest = await get_latest_balances()
        totals = await get_daily_totals(days=1)
        total_today = sum(r["total"] for r in totals) if totals else 0

        lines = [f"<h2>Balance summary — {datetime.now().strftime('%d/%m/%Y')}</h2>"]
        lines.append(f"<p><strong>Total: {total_today:,.2f} EUR</strong></p><hr>")
        lines.append("<table border='0' cellpadding='6' style='font-family:monospace'>")
        lines.append("<tr><th>Banco</th><th>IBAN</th><th>Saldo</th></tr>")
        for row in latest:
            iban_short = f"****{row['iban'][-4:]}" if row.get("iban") else "—"
            lines.append(
                f"<tr><td>{row['bank_name']}</td>"
                f"<td>{iban_short}</td>"
                f"<td align='right'>{row['amount']:,.2f} {row['currency']}</td></tr>"
            )
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
                subject=f"💰 Saldos {datetime.now().strftime('%d/%m/%Y')} — {total_today:,.0f} EUR",
                body_html="\n".join(lines),
            )
            logger.info("Summary email sent to %s", recipient)

    except ImportError:
        logger.warning("iasuite_common not found — skipping email report")
    except Exception as exc:
        logger.error("Failed to send summary email: %s", exc)
