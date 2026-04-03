"""
Enable Banking API client.
Docs: https://enablebanking.com/accounts-api/
Auth: JWT signed with your RSA private key (generated on enablebanking.com dashboard).
"""

import uuid
import time
import logging
from datetime import datetime, timedelta, timezone
from pathlib import Path

import httpx
import jwt as pyjwt

logger = logging.getLogger(__name__)

ENABLE_BANKING_BASE = "https://api.tilisy.com"


class EnableBankingClient:
    def __init__(self, app_id: str, private_key_pem: str):
        self.app_id = app_id
        self.private_key = private_key_pem

    # ------------------------------------------------------------------
    # JWT auth token (short-lived, signed with RSA key)
    # ------------------------------------------------------------------
    def _make_token(self) -> str:
        now = int(time.time())
        payload = {
            "iss": self.app_id,
            "iat": now,
            "exp": now + 3600,
            "jti": str(uuid.uuid4()),
        }
        return pyjwt.encode(payload, self.private_key, algorithm="RS256")

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self._make_token()}",
            "Content-Type": "application/json",
        }

    # ------------------------------------------------------------------
    # Step 1: Start bank authorisation — returns redirect URL for the user
    # ------------------------------------------------------------------
    async def start_auth(
        self,
        aspsp_name: str,
        aspsp_country: str,
        redirect_url: str,
        state: str,
    ) -> str:
        valid_until = (
            datetime.now(timezone.utc) + timedelta(days=90)
        ).strftime("%Y-%m-%dT%H:%M:%SZ")

        body = {
            "aspsp": {"name": aspsp_name, "country": aspsp_country},
            "state": state,
            "redirect_url": redirect_url,
            "access": {"valid_until": valid_until},
        }

        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(
                f"{ENABLE_BANKING_BASE}/auth",
                headers=self._headers(),
                json=body,
            )
            resp.raise_for_status()
            data = resp.json()
            logger.info("Auth started for %s → %s", aspsp_name, data.get("url", "")[:60])
            return data["url"]

    # ------------------------------------------------------------------
    # Step 2: Exchange auth code for a session (after user redirect back)
    # ------------------------------------------------------------------
    async def create_session(self, code: str) -> dict:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(
                f"{ENABLE_BANKING_BASE}/sessions",
                headers=self._headers(),
                json={"code": code},
            )
            resp.raise_for_status()
            data = resp.json()
            logger.info("Session created: %s accounts", len(data.get("accounts", [])))
            return data  # {session_id, accounts: [{id, iban, name, ...}]}

    # ------------------------------------------------------------------
    # Step 3: Fetch balances for one account using a stored session
    # ------------------------------------------------------------------
    async def get_account_balances(self, account_id: str, session_id: str) -> list[dict]:
        headers = self._headers()
        headers["X-Session-Id"] = session_id

        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.get(
                f"{ENABLE_BANKING_BASE}/accounts/{account_id}/balances",
                headers=headers,
            )
            resp.raise_for_status()
            return resp.json().get("balances", [])

    # ------------------------------------------------------------------
    # Convenience: fetch balances for all accounts in a session
    # ------------------------------------------------------------------
    async def fetch_all_balances(self, accounts: list[dict], session_id: str) -> list[dict]:
        results = []
        for account in accounts:
            try:
                balances = await self.get_account_balances(account["id"], session_id)
                # Prefer closingBooked, fall back to interimAvailable
                balance = next(
                    (b for b in balances if b.get("name") == "closingBooked"),
                    next(iter(balances), None),
                )
                if balance:
                    results.append(
                        {
                            "account_id": account["id"],
                            "iban": account.get("account_id", {}).get("iban", ""),
                            "name": account.get("name", ""),
                            "amount": float(balance["balanceAmount"]["amount"]),
                            "currency": balance["balanceAmount"]["currency"],
                            "balance_type": balance.get("name", ""),
                        }
                    )
            except Exception as exc:
                logger.error("Failed fetching balance for %s: %s", account.get("id"), exc)
        return results
