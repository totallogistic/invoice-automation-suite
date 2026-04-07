"""
Database layer. Uses a single SQLite file at data/balances.db.
Tables:
  banks    — configured banks and their session state
  accounts — accounts discovered after authorisation
  balances — daily balance snapshots
"""

import logging
import aiosqlite
from datetime import datetime, timezone
from pathlib import Path

logger = logging.getLogger(__name__)

DB_PATH = Path(__file__).parent / "data" / "balances.db"


async def init_db():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    async with aiosqlite.connect(DB_PATH) as db:
        await db.executescript("""
            CREATE TABLE IF NOT EXISTS banks (
                id          TEXT PRIMARY KEY,
                name        TEXT NOT NULL,
                aspsp_name  TEXT NOT NULL,
                country     TEXT DEFAULT 'ES',
                session_id  TEXT,
                session_expires_at TEXT,
                connected_at TEXT,
                last_sync_at TEXT,
                status      TEXT DEFAULT 'pending'
            );

            CREATE TABLE IF NOT EXISTS accounts (
                id          TEXT PRIMARY KEY,
                bank_id     TEXT NOT NULL,
                iban        TEXT,
                name        TEXT,
                currency    TEXT DEFAULT 'EUR',
                FOREIGN KEY (bank_id) REFERENCES banks(id)
            );

            CREATE TABLE IF NOT EXISTS balances (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                account_id   TEXT NOT NULL,
                bank_id      TEXT NOT NULL,
                amount       REAL NOT NULL,
                currency     TEXT DEFAULT 'EUR',
                balance_type TEXT DEFAULT 'closingBooked',
                recorded_at  TEXT NOT NULL,
                FOREIGN KEY (account_id) REFERENCES accounts(id)
            );
        """)
        await db.commit()
    logger.info("DB initialised at %s", DB_PATH)


# ------------------------------------------------------------------
# Banks
# ------------------------------------------------------------------

async def get_all_banks() -> list[dict]:
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("SELECT * FROM banks ORDER BY name") as cur:
            rows = await cur.fetchall()
    return [dict(r) for r in rows]


async def upsert_bank(bank_id: str, name: str, aspsp_name: str, country: str = "ES"):
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            INSERT INTO banks (id, name, aspsp_name, country)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET name=excluded.name, aspsp_name=excluded.aspsp_name
        """, (bank_id, name, aspsp_name, country))
        await db.commit()


async def save_session(bank_id: str, session_id: str, expires_at: str, accounts: list[dict]):
    now = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute("""
            UPDATE banks
            SET session_id=?, session_expires_at=?, connected_at=?, status='connected'
            WHERE id=?
        """, (session_id, expires_at, now, bank_id))

        for acc in accounts:
            # Enable Banking usa 'uid' como identificador de cuenta
            acc_id = acc.get("uid") or acc.get("id") or acc.get("resource_id", "")
            iban = acc.get("account_id", {}).get("iban", "") if isinstance(acc.get("account_id"), dict) else ""
            await db.execute("""
                INSERT INTO accounts (id, bank_id, iban, name, currency)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET iban=excluded.iban, name=excluded.name
            """, (
                acc_id,
                bank_id,
                iban,
                acc.get("name", ""),
                acc.get("currency", "EUR"),
            ))
        await db.commit()


async def save_balances(bank_id: str, balances: list[dict]):
    now = datetime.now(timezone.utc).isoformat()
    async with aiosqlite.connect(DB_PATH) as db:
        for b in balances:
            await db.execute("""
                INSERT INTO balances (account_id, bank_id, amount, currency, balance_type, recorded_at)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (b["account_id"], bank_id, b["amount"], b["currency"], b.get("balance_type", ""), now))
        await db.execute(
            "UPDATE banks SET last_sync_at=? WHERE id=?", (now, bank_id)
        )
        await db.commit()


# ------------------------------------------------------------------
# Dashboard queries
# ------------------------------------------------------------------

async def get_latest_balances() -> list[dict]:
    """Latest balance per account, with bank info."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("""
            SELECT b.name AS bank_name,
                   a.iban,
                   a.name AS account_name,
                   bl.amount,
                   bl.currency,
                   bl.recorded_at
            FROM balances bl
            JOIN accounts a ON a.id = bl.account_id
            JOIN banks b    ON b.id = bl.bank_id
            WHERE bl.id IN (
                SELECT MAX(id) FROM balances GROUP BY account_id
            )
            ORDER BY b.name, a.iban
        """) as cur:
            rows = await cur.fetchall()
    return [dict(r) for r in rows]


async def get_daily_totals(days: int = 60) -> list[dict]:
    """Total balance per day across all accounts (EUR)."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("""
            SELECT DATE(recorded_at) AS day, SUM(amount) AS total
            FROM (
                SELECT account_id, MAX(id) AS mid
                FROM balances
                WHERE recorded_at >= DATE('now', ?)
                GROUP BY account_id, DATE(recorded_at)
            ) latest
            JOIN balances ON balances.id = latest.mid
            GROUP BY day
            ORDER BY day
        """, (f"-{days} days",)) as cur:
            rows = await cur.fetchall()
    return [dict(r) for r in rows]


async def get_history_per_bank(days: int = 60) -> list[dict]:
    """Daily total per bank for chart series."""
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        async with db.execute("""
            SELECT bk.name AS bank_name,
                   DATE(bl.recorded_at) AS day,
                   SUM(bl.amount) AS total
            FROM balances bl
            JOIN banks bk ON bk.id = bl.bank_id
            WHERE bl.recorded_at >= DATE('now', ?)
              AND bl.id IN (
                  SELECT MAX(id) FROM balances
                  WHERE recorded_at >= DATE('now', ?)
                  GROUP BY account_id, DATE(recorded_at)
              )
            GROUP BY bk.name, day
            ORDER BY bk.name, day
        """, (f"-{days} days", f"-{days} days")) as cur:
            rows = await cur.fetchall()
    return [dict(r) for r in rows]
