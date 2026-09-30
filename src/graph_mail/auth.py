"""Token providers. Both expose get_token() so the rest of the code doesn't care which is used."""
from __future__ import annotations

import os
from typing import Protocol

import msal

from .config import MailConfig

APP_SCOPE = ["https://graph.microsoft.com/.default"]


class AuthError(RuntimeError):
    pass


class TokenProvider(Protocol):
    def get_token(self) -> str: ...


class ClientSecretAuth:
    """Application (app-only) permissions via client credentials."""

    def __init__(self, cfg: MailConfig) -> None:
        self._app = msal.ConfidentialClientApplication(
            cfg.client_id, authority=cfg.authority, client_credential=cfg.client_secret
        )

    def get_token(self) -> str:
        # MSAL caches in-memory and only hits the network when the token is near expiry.
        result = self._app.acquire_token_for_client(scopes=APP_SCOPE)
        if "access_token" not in result:
            raise AuthError(f"Client-credentials failed: {result.get('error')}: {result.get('error_description')}")
        return result["access_token"]


class DelegatedAuth:
    """Delegated (signed-in user) permissions using device-code flow + a persistent token cache.

    First run prompts the user to sign in; later runs refresh silently from the cache file.
    The app registration must have 'Allow public client flows' enabled.
    """

    def __init__(self, cfg: MailConfig, interactive: bool = True) -> None:
        self._cfg = cfg
        self._interactive = interactive
        self._cache = msal.SerializableTokenCache()
        self._path = cfg.token_cache_path
        if self._path and self._path.exists():
            self._cache.deserialize(self._path.read_text())
        self._app = msal.PublicClientApplication(cfg.client_id, authority=cfg.authority, token_cache=self._cache)

    def _persist(self) -> None:
        if self._path and self._cache.has_state_changed:
            self._path.parent.mkdir(parents=True, exist_ok=True)
            self._path.write_text(self._cache.serialize())
            try:
                os.chmod(self._path, 0o600)
            except OSError:
                pass

    def login(self) -> str:
        flow = self._app.initiate_device_flow(scopes=list(self._cfg.delegated_scopes))
        if "user_code" not in flow:
            raise AuthError(f"Could not start device flow: {flow}")
        print(flow["message"])
        result = self._app.acquire_token_by_device_flow(flow)
        self._persist()
        if "access_token" not in result:
            raise AuthError(f"Device login failed: {result.get('error')}: {result.get('error_description')}")
        return result["access_token"]

    def get_token(self) -> str:
        accounts = self._app.get_accounts()
        if self._cfg.mailbox.lower() != "me":
            accounts = [a for a in accounts if a.get("username", "").lower() == self._cfg.mailbox.lower()] or accounts
        if accounts:
            result = self._app.acquire_token_silent(list(self._cfg.delegated_scopes), account=accounts[0])
            self._persist()
            if result and "access_token" in result:
                return result["access_token"]
        if not self._interactive:
            raise AuthError("No valid cached token; run an interactive login first (DelegatedAuth.login()).")
        return self.login()


def build_auth(cfg: MailConfig) -> TokenProvider:
    return ClientSecretAuth(cfg) if cfg.auth_mode == "client_secret" else DelegatedAuth(cfg)
