"""Client for the Alvys TMS public API (OAuth2 client-credentials flow).

Endpoint/auth verified against a live call on 2026-08-26:
  - Token: POST https://auth.alvys.com/oauth/token
  - Loads search: POST https://integrations.alvys.com/api/p/v1/loads/search
    (a *search* body -- it 400s unless at least one of Status, PONumbers,
    UpdatedBy, CustomerId, LoadNumbers, OrderNumbers or
    CustomerSalesAgentId is provided). `dateRange` is accepted by the API
    but does NOT filter results server-side (confirmed empirically -- Total
    stayed identical across wildly different ranges), so date filtering has
    to happen client-side; see `detention.filter_rows_by_date`.
  - Max `pageSize` is 500.
"""

import os
import time

import requests

TOKEN_URL = "https://auth.alvys.com/oauth/token"
AUDIENCE = "https://api.alvys.com/public/"
INTEGRATIONS_BASE_URL = "https://integrations.alvys.com"

MAX_PAGE_SIZE = 500

# All Status values the /loads/search endpoint accepts (from its own
# validation error message). Statuses that imply a truck was actually
# dispatched/underway/completed -- i.e. the ones that can have stop-level
# ArrivedAt/DepartedAt data worth checking for detention.
EXECUTED_STATUSES = [
    "Dispatched",
    "In Transit",
    "En-Route",
    "TONU",
    "Trip Completed",
    "Delivered",
    "Invoiced",
    "Paid",
    "Carrier Paid",
    "Released",
    "Released-Carrier Paid",
    "Completed",
    "Financed",
]


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

    def _post(self, path, body):
        headers = {"Authorization": f"Bearer {self._get_token()}"}
        response = requests.post(f"{INTEGRATIONS_BASE_URL}{path}", headers=headers, json=body, timeout=30)
        response.raise_for_status()
        return response.json()

    def search_loads_page(self, page=0, page_size=MAX_PAGE_SIZE, statuses=None, **extra_filters):
        """One page of /loads/search. At least one filter (statuses here) is required by the API."""
        body = {"page": page, "pageSize": page_size, "status": statuses or EXECUTED_STATUSES}
        body.update(extra_filters)
        return self._post("/api/p/v1/loads/search", body)

    def search_all_loads(self, statuses=None, **extra_filters):
        """Page through /loads/search and return every matching Load dict.

        Client-side date filtering happens downstream (see
        `detention.filter_rows_by_date`) since the API's `dateRange` filter
        doesn't actually narrow results.
        """
        loads = []
        page = 0
        while True:
            result = self.search_loads_page(page=page, statuses=statuses, **extra_filters)
            items = result.get("Items", [])
            loads.extend(items)
            total = result.get("Total", len(loads))
            if len(loads) >= total or not items:
                break
            page += 1
        return loads

    def search_invoices(self, **params):
        # NOTE: unlike /loads/search, this endpoint has NOT been verified against
        # a live call. Only the base URL (integrations.alvys.com, confirmed via
        # /loads/search) has been updated here; the path and body shape are still
        # a best-effort guess -- check a live response and adjust before relying
        # on this for anything.
        body = {"page": 0, "pageSize": MAX_PAGE_SIZE}
        body.update(params)
        return self._post("/api/p/v1/invoices/search", body)
