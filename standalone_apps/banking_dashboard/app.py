"""
Banking Dashboard — FastAPI app.

Endpoints:
  GET  /              → dashboard (balances + charts)
  GET  /setup         → bank setup page (connect/reconnect banks)
  GET  /setup/connect/{bank_id}  → start Enable Banking auth flow
  GET  /callback      → OAuth callback from Enable Banking
  POST /sync          → manual sync trigger
  GET  /api/data      → JSON for chart refresh

Banks are defined in banks.yaml (copy from banks.yaml.example).
"""

import os
import uuid
import logging
import yaml
from pathlib import Path
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from apscheduler.schedulers.asyncio import AsyncIOScheduler

import db
from bank_fetcher import EnableBankingClient
from scheduler import sync_all_banks

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).parent
BANKS_CONFIG = BASE_DIR / "banks.yaml"

# ------------------------------------------------------------------
# App startup / shutdown
# ------------------------------------------------------------------
scheduler = AsyncIOScheduler()

@asynccontextmanager
async def lifespan(app: FastAPI):
    await db.init_db()
    await _seed_banks_from_config()

    # Daily sync at 07:00
    scheduler.add_job(
        sync_all_banks,
        trigger="cron",
        hour=int(os.environ.get("SYNC_HOUR", "7")),
        minute=0,
        id="daily_sync",
        replace_existing=True,
    )
    scheduler.start()
    logger.info("Scheduler started — daily sync at %s:00", os.environ.get("SYNC_HOUR", "7"))
    yield
    scheduler.shutdown()


app = FastAPI(title="Banking Dashboard", lifespan=lifespan)
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _get_client() -> EnableBankingClient:
    app_id = os.environ["ENABLE_BANKING_APP_ID"]
    key_path = os.environ.get("ENABLE_BANKING_KEY_PATH", str(BASE_DIR / "keys/private.pem"))
    return EnableBankingClient(app_id, open(key_path).read())


async def _seed_banks_from_config():
    if not BANKS_CONFIG.exists():
        logger.warning("banks.yaml not found — no banks seeded. Copy banks.yaml.example.")
        return
    config = yaml.safe_load(BANKS_CONFIG.read_text())
    for bank in config.get("banks", []):
        await db.upsert_bank(
            bank_id=bank["id"],
            name=bank["name"],
            aspsp_name=bank["aspsp_name"],
            country=bank.get("country", "ES"),
        )
    logger.info("Seeded %d banks from config", len(config.get("banks", [])))


# Temporary store for pending auth states (in-memory is fine — small app)
_pending_states: dict[str, str] = {}  # state → bank_id


# ------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    latest = await db.get_latest_balances()
    history = await db.get_history_per_bank(days=60)
    daily = await db.get_daily_totals(days=60)
    total = sum(r["amount"] for r in latest)

    # Structure for Chart.js
    bank_names = sorted(set(r["bank_name"] for r in history))
    all_days = sorted(set(r["day"] for r in history))
    series = []
    for bank in bank_names:
        day_map = {r["day"]: r["total"] for r in history if r["bank_name"] == bank}
        series.append({
            "label": bank,
            "data": [day_map.get(d, None) for d in all_days],
        })

    return templates.TemplateResponse("dashboard.html", {
        "request": request,
        "latest": latest,
        "total": total,
        "bank_names": bank_names,
        "days": all_days,
        "series": series,
        "daily_totals": [{"day": r["day"], "total": r["total"]} for r in daily],
    })


@app.get("/setup", response_class=HTMLResponse)
async def setup_page(request: Request):
    banks = await db.get_all_banks()
    return templates.TemplateResponse("setup.html", {
        "request": request,
        "banks": banks,
    })


@app.get("/setup/connect/{bank_id}")
async def connect_bank(bank_id: str, request: Request):
    banks = await db.get_all_banks()
    bank = next((b for b in banks if b["id"] == bank_id), None)
    if not bank:
        raise HTTPException(404, "Bank not found")

    state = str(uuid.uuid4())
    _pending_states[state] = bank_id

    redirect_url = str(request.url_for("oauth_callback"))
    client = _get_client()
    auth_url = await client.start_auth(
        aspsp_name=bank["aspsp_name"],
        aspsp_country=bank["country"],
        redirect_url=redirect_url,
        state=state,
    )
    return RedirectResponse(auth_url)


@app.get("/callback")
async def oauth_callback(code: str, state: str):
    bank_id = _pending_states.pop(state, None)
    if not bank_id:
        raise HTTPException(400, "Unknown or expired state. Please retry from Setup.")

    client = _get_client()
    session_data = await client.create_session(code)

    session_id = session_data["session_id"]
    accounts = session_data.get("accounts", [])
    expires_at = session_data.get("access", {}).get("valid_until", "")

    await db.save_session(bank_id, session_id, expires_at, accounts)
    logger.info("Bank %s connected — %d accounts", bank_id, len(accounts))

    # Trigger an immediate sync for this bank
    await sync_all_banks()

    return RedirectResponse("/setup")


@app.post("/sync")
async def manual_sync():
    result = await sync_all_banks()
    return JSONResponse(result)


@app.get("/api/data")
async def api_data():
    latest = await db.get_latest_balances()
    totals = await db.get_daily_totals(days=60)
    history = await db.get_history_per_bank(days=60)
    return {
        "latest": latest,
        "totals": totals,
        "history": history,
        "total": sum(r["amount"] for r in latest),
    }
