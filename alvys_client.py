"""Client for the Alvys TMS public API (OAuth2 client-credentials flow).

Confirmed against Alvys's Power BI integration docs
(docs.alvys.com/docs/create-queries-and-retrieve-data-from-the-alvys-api):
search endpoints are POST requests with a JSON body (page/pageSize plus
filters), against https://integrations.alvys.com/api/p/v1/..., and return
paginated results under an "Items" key. Response field names for Loads
specifically weren't shown in that doc (only Drivers/Fuel examples expand
their fields) -- confirm/adjust ALVYS_ITEMS_KEYS and the field-name
constants in generate_report.py against a real response.
"""

import os
import time

import requests

TOKEN_URL = "https://auth.alvys.com/oauth/token"
API_BASE_URL = "https://integrations.alvys.com/api/p/v1"
AUDIENCE = "https://api.alvys.com/public/"

MAX_PAGES = 50  # safety cap: 50 pages * pageSize records, in case pagination never terminates


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

    def post(self, path, body):
        """Low-level POST against the Alvys public API. `path` is relative, e.g. '/loads/search'."""
        headers = {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }
        response = requests.post(f"{API_BASE_URL}{path}", headers=headers, json=body, timeout=30)
        response.raise_for_status()
        return response.json()

    def _extract_items(self, payload):
        for key in ("Items", "items", "Results", "results", "data"):
            if isinstance(payload, dict) and key in payload:
                return payload[key]
        return payload if isinstance(payload, list) else []

    def _search_paginated(self, path, body, page_size):
        items = []
        for page in range(MAX_PAGES):
            page_body = dict(body, page=page, pageSize=page_size)
            payload = self.post(path, page_body)
            page_items = self._extract_items(payload)
            items.extend(page_items)
            if len(page_items) < page_size:
                break
        return {"data": items}

    def search_loads(self, start_date=None, end_date=None, status=None, page_size=200, **extra):
        # NOTE: confirmed request shape from Alvys's Power BI docs example;
        # response field names for individual loads are unconfirmed -- check
        # a real response and adjust generate_report.py's field-name lists.
        body = dict(extra)
        if start_date or end_date:
            body["dateRange"] = {"startDate": start_date, "endDate": end_date}
        if status:
            body["status"] = status
        return self._search_paginated("/loads/search", body, page_size)

    def search_invoices(self, start_date=None, end_date=None, status=None, page_size=100,
                         date_field="invoicedDateRange", **extra):
        # NOTE: same caveat as search_loads. Alvys's example shows both
        # invoicedDateRange and paidDateRange -- date_field picks which one
        # start_date/end_date apply to (defaults to invoicedDateRange).
        body = dict(extra)
        if start_date or end_date:
            body[date_field] = {"start": start_date, "end": end_date}
        if status:
            body["status"] = status
        return self._search_paginated("/invoices/search", body, page_size)
