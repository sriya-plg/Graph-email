"""Configuration for one mailbox / app registration."""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Mapping

AuthMode = Literal["client_secret", "delegated"]


@dataclass(frozen=True)
class MailConfig:
    tenant_id: str
    client_id: str
    mailbox: str  # UPN/email of the mailbox; use "me" with delegated auth
    auth_mode: AuthMode = "client_secret"
    client_secret: str | None = None
    token_cache_path: Path | None = None  # delegated only
    delegated_scopes: tuple[str, ...] = ("Mail.Read", "Mail.ReadWrite")
    graph_base_url: str = "https://graph.microsoft.com/v1.0"
    authority_host: str = "https://login.microsoftonline.com"
    timeout: float = 30.0
    max_retries: int = 5

    def __post_init__(self) -> None:
        if self.auth_mode not in ("client_secret", "delegated"):
            raise ValueError(f"auth_mode must be 'client_secret' or 'delegated', got {self.auth_mode!r}")
        for name in ("tenant_id", "client_id", "mailbox"):
            if not getattr(self, name):
                raise ValueError(f"{name} is required")
        if self.auth_mode == "client_secret":
            if not self.client_secret:
                raise ValueError("client_secret is required when auth_mode='client_secret'")
            if self.mailbox.lower() == "me":
                raise ValueError("App-only auth has no 'me'; set mailbox to the user's email/UPN")

    @property
    def authority(self) -> str:
        return f"{self.authority_host}/{self.tenant_id}"

    @classmethod
    def from_dict(cls, data: Mapping[str, object]) -> "MailConfig":
        data = dict(data)
        if data.get("token_cache_path"):
            data["token_cache_path"] = Path(str(data["token_cache_path"]))
        if isinstance(data.get("delegated_scopes"), (list, str)):
            scopes = data["delegated_scopes"]
            data["delegated_scopes"] = tuple(
                s.strip() for s in (scopes.split(",") if isinstance(scopes, str) else scopes) if s.strip()
            )
        return cls(**data)  # type: ignore[arg-type]

    @classmethod
    def from_env(cls, prefix: str = "GRAPH_MAIL_", env: Mapping[str, str] | None = None) -> "MailConfig":
        """Read e.g. GRAPH_MAIL_TENANT_ID, GRAPH_MAIL_CLIENT_ID, GRAPH_MAIL_MAILBOX,
        GRAPH_MAIL_AUTH_MODE, GRAPH_MAIL_CLIENT_SECRET, GRAPH_MAIL_TOKEN_CACHE_PATH.
        Use a different prefix per project/mailbox (e.g. 'ORDERS_MAIL_')."""
        env = os.environ if env is None else env
        data: dict[str, object] = {}
        for field in ("tenant_id", "client_id", "mailbox", "auth_mode", "client_secret",
                      "token_cache_path", "delegated_scopes"):
            value = env.get(f"{prefix}{field.upper()}")
            if value:
                data[field] = value
        return cls.from_dict(data)
