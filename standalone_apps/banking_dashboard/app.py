"""
Banking Dashboard — FastAPI app.

Endpoints:
  GET  /              → dashboard
  GET  /setup         → conectar/reconectar bancos
  GET  /setup/connect/{bank_id}  → inicia flujo OAuth
  GET  /callback      → callback de Enable Banking
  POST /sync          → sync manual
  GET  /export/excel  → descarga saldos.xlsx
  GET  /api/data      → JSON para el dashboard
"""

import os
import uuid
import logging
import yaml
from pathlib import Path
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse, FileResponse
from fastapi.templating import Jinja2Templates
from apscheduler.schedulers.asyncio import AsyncIOScheduler

import db
from bank_fetcher import EnableBankingClient
from scheduler import sync_all_banks

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
logger = logging.getLogger(__name__)

BASE_DIR     = Path(__file__).parent
BANKS_CONFIG = BASE_DIR / "banks.yaml"

scheduler = AsyncIOScheduler()

@asynccontextmanager
async def lifespan(app: FastAPI):
    await db.init_db()
    await _seed_banks_from_config()
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

_pending_states: dict[str, str] = {}


def _get_client() -> EnableBankingClient:
    app_id   = os.environ["ENABLE_BANKING_APP_ID"]
    key_path = os.environ.get("ENABLE_BANKING_KEY_PATH", str(BASE_DIR / "keys/private.key"))
    return EnableBankingClient(app_id, open(key_path).read())


async def _seed_banks_from_config():
    if not BANKS_CONFIG.exists():
        logger.warning("banks.yaml not found — no banks seeded.")
        return
    config = yaml.safe_load(BANKS_CONFIG.read_text())
    for bank in config.get("banks", []):
        await db.upsert_bank(bank["id"], bank["name"], bank["aspsp_name"], bank.get("country", "ES"))
    logger.info("Seeded %d banks from config", len(config.get("banks", [])))


# ------------------------------------------------------------------
# Dashboard
# ------------------------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def dashboard(request: Request):
    latest  = await db.get_latest_balances()
    history = await db.get_history_per_bank(days=60)
    daily   = await db.get_daily_totals(days=60)
    total   = sum(r["amount"] for r in latest)

    bank_names = sorted(set(r["bank_name"] for r in history))
    all_days   = sorted(set(r["day"] for r in history))
    series = []
    for bank in bank_names:
        day_map = {r["day"]: r["total"] for r in history if r["bank_name"] == bank}
        series.append({"label": bank, "data": [day_map.get(d) for d in all_days]})

    return templates.TemplateResponse("dashboard.html", {
        "request":      request,
        "latest":       latest,
        "total":        total,
        "bank_names":   bank_names,
        "days":         all_days,
        "series":       series,
        "daily_totals": [{"day": r["day"], "total": r["total"]} for r in daily],
        "excel_ready":  (BASE_DIR / "data" / "saldos.xlsx").exists(),
    })


# ------------------------------------------------------------------
# Setup
# ------------------------------------------------------------------

@app.get("/setup", response_class=HTMLResponse)
async def setup_page(request: Request):
    banks = await db.get_all_banks()
    return templates.TemplateResponse("setup.html", {"request": request, "banks": banks})


@app.get("/setup/connect/{bank_id}")
async def connect_bank(bank_id: str, request: Request):
    banks = await db.get_all_banks()
    bank  = next((b for b in banks if b["id"] == bank_id), None)
    if not bank:
        raise HTTPException(404, "Bank not found")

    state = str(uuid.uuid4())
    _pending_states[state] = bank_id
    redirect_url = str(request.url_for("oauth_callback"))
    auth_url = await _get_client().start_auth(bank["aspsp_name"], bank["country"], redirect_url, state)
    return RedirectResponse(auth_url)


@app.get("/callback")
async def oauth_callback(code: str, state: str):
    bank_id = _pending_states.pop(state, None)
    if not bank_id:
        raise HTTPException(400, "Estado desconocido o expirado. Inténtalo de nuevo desde Setup.")

    session_data = await _get_client().create_session(code)
    session_id   = session_data["session_id"]
    accounts     = session_data.get("accounts", [])
    expires_at   = session_data.get("access", {}).get("valid_until", "")

    await db.save_session(bank_id, session_id, expires_at, accounts)
    logger.info("Bank %s connected — %d accounts", bank_id, len(accounts))
    await sync_all_banks()
    return RedirectResponse("/setup")


# ------------------------------------------------------------------
# Sync manual
# ------------------------------------------------------------------

@app.post("/sync")
async def manual_sync():
    result = await sync_all_banks()
    return JSONResponse(result)


# ------------------------------------------------------------------
# Export Situación Financiera
# ------------------------------------------------------------------

@app.get("/export/situacion")
async def export_situacion():
    from bancos_export import generate_bancos
    from db import get_history_per_bank
    try:
        path = await generate_bancos(history=None)
    except FileNotFoundError as e:
        raise HTTPException(404, str(e))
    except ValueError as e:
        raise HTTPException(400, str(e))
    today = __import__("datetime").date.today().strftime("%Y%m%d")
    return FileResponse(
        path=str(path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"situacion_financiera_{today}.xlsx",
    )


# ------------------------------------------------------------------
# Export Excel
# ------------------------------------------------------------------

@app.get("/export/excel")
async def export_excel():
    from excel_export import generate_excel
    latest  = await db.get_latest_balances()
    history = await db.get_history_per_bank(days=90)
    daily   = await db.get_daily_totals(days=90)

    if not latest:
        raise HTTPException(404, "No hay datos. Realiza un sync primero.")

    excel_path = await generate_excel(latest, history, daily)
    import datetime
    today = datetime.date.today().strftime("%Y%m%d")
    return FileResponse(
        path=str(excel_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"saldos_bancarios_{today}.xlsx",
    )


# ------------------------------------------------------------------
# API JSON
# ------------------------------------------------------------------

@app.get("/api/data")
async def api_data():
    latest  = await db.get_latest_balances()
    totals  = await db.get_daily_totals(days=60)
    history = await db.get_history_per_bank(days=60)
    return {"latest": latest, "totals": totals, "history": history, "total": sum(r["amount"] for r in latest)}
