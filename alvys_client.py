"""Minimal client for the Alvys TMS public API (OAuth2 client-credentials flow)."""

import os
import time

import requests

TOKEN_URL = "https://auth.alvys.com/oauth/token"
API_BASE_URL = "https://integrations.alvys.com"
AUDIENCE = "https://api.alvys.com/public/"

# Alvys requires at least one non-date search parameter on both endpoints; a
# date range alone isn't enough (verified via the API's own validation
# errors). Defaulting `status` to the full enum below when the caller
# doesn't pass one keeps a plain "everything in this date range" call
# working without every caller needing to know that quirk.
ALL_LOAD_STATUSES = [
    "TONU", "Dispatched", "In Transit", "Queued", "Paid", "Open", "Completed",
    "Financed", "Cancelled", "Covered", "In Review", "Released-Carrier Paid",
    "Reserved", "Trip Completed", "En-Route", "Quoted", "Admin", "Invoiced",
    "Carrier Paid", "Released", "Delivered",
]
ALL_INVOICE_STATUSES = ["Draft", "AwaitingPayment", "Paid"]


class AlvysAuthError(Exception):
    pass


class AlvysClient:
    def __init__(self, client_id=None, client_secret=None):
        self.client_id = client_id or os.environ.get("ALVYS_CLIENT_ID")
        self.client_secret = client_secret or os.environ.get("ALVYS_CLIENT_SECRET")
        if not self.client_id or not self.client_secret:
            raise AlvysAuthError(
                "Missing Alvys credentials. Set ALVYS_CLIENT_ID and ALVYS_CLIENT_SECRET "
                "as environment variables (see .env.example)."
            )
        self._token = None
        self._token_expires_at = 0

    def _fetch_token(self):
        response = requests.post(
            TOKEN_URL,
            json={
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "audience": AUDIENCE,
                "grant_type": "client_credentials",
            },
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()
        self._token = payload["access_token"]
        self._token_expires_at = time.time() + payload.get("expires_in", 3600) - 60

    def _get_token(self):
        if not self._token or time.time() >= self._token_expires_at:
            self._fetch_token()
        return self._token

    def post(self, path, body=None):
        """Low-level POST against the Alvys integrations API. `path` is relative, e.g. '/api/p/v1/loads/search'."""
        headers = {"Authorization": f"Bearer {self._get_token()}"}
        response = requests.post(f"{API_BASE_URL}{path}", headers=headers, json=body or {}, timeout=30)
        response.raise_for_status()
        return response.json()

    def search_loads(self, start_date=None, end_date=None, page=0, page_size=200, status=None, **body):
        """POST /api/p/v1/loads/search. `start_date`/`end_date` are ISO-8601 datetimes."""
        payload = dict(body)
        payload["page"] = page
        payload["pageSize"] = page_size
        if start_date or end_date:
            payload["dateRange"] = {"startDate": start_date, "endDate": end_date}
        payload["status"] = status if status else ALL_LOAD_STATUSES
        return self.post("/api/p/v1/loads/search", body=payload)

    def search_invoices(self, start_date=None, end_date=None, page=0, page_size=200, status=None, **body):
        """POST /api/p/v1/invoices/search. `start_date`/`end_date` are ISO-8601 datetimes, applied to invoicedDateRange."""
        payload = dict(body)
        payload["page"] = page
        payload["pageSize"] = page_size
        if start_date or end_date:
            payload["invoicedDateRange"] = {"start": start_date, "end": end_date}
        payload["status"] = status if status else ALL_INVOICE_STATUSES
        return self.post("/api/p/v1/invoices/search", body=payload)
