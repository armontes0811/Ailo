"""Minimal client for the Alvys TMS public API (OAuth2 client-credentials flow)."""

import os
import time

import requests

TOKEN_URL = "https://auth.alvys.com/oauth/token"
API_BASE_URL = "https://api.alvys.com/public"
AUDIENCE = "https://api.alvys.com/public/"


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

    def get(self, path, params=None):
        """Low-level GET against the Alvys public API. `path` is relative, e.g. '/loads/search'."""
        headers = {"Authorization": f"Bearer {self._get_token()}"}
        response = requests.get(f"{API_BASE_URL}{path}", headers=headers, params=params, timeout=30)
        response.raise_for_status()
        return response.json()

    def search_loads(self, start_date=None, end_date=None, **params):
        # NOTE: endpoint path/params are best-effort from Alvys's public docs
        # (docs.alvys.com/docs/create-queries-and-retrieve-data-from-the-alvys-api).
        # Verify against a live call / the docs and adjust if the field names differ.
        query = dict(params)
        if start_date:
            query["startDate"] = start_date
        if end_date:
            query["endDate"] = end_date
        return self.get("/loads/search", params=query)

    def search_invoices(self, start_date=None, end_date=None, **params):
        # NOTE: same caveat as search_loads -- verify against live docs/response shape.
        query = dict(params)
        if start_date:
            query["startDate"] = start_date
        if end_date:
            query["endDate"] = end_date
        return self.get("/invoices/search", params=query)
