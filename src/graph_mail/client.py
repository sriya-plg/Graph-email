"""Thin Microsoft Graph mail client: list messages, fetch attachments, mark read, move."""
from __future__ import annotations

import re
import time
from datetime import datetime, timezone
from typing import Any, Iterator
from urllib.parse import quote

import requests

from .auth import TokenProvider, build_auth
from .config import MailConfig
from .models import Attachment

DEFAULT_SELECT = "id,subject,from,receivedDateTime,hasAttachments,isRead,conversationId,internetMessageId,webLink"
RETRY_STATUS = {429, 500, 502, 503, 504}
EPOCH = datetime(1970, 1, 1, tzinfo=timezone.utc)


class GraphError(RuntimeError):
    pass


def _iso(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


class GraphMailClient:
    def __init__(self, cfg: MailConfig, auth: TokenProvider | None = None) -> None:
        self.cfg = cfg
        self.auth = auth or build_auth(cfg)
        self._http = requests.Session()
        if cfg.mailbox.lower() == "me":
            self._user = "/me"
        else:
            self._user = f"/users/{quote(cfg.mailbox)}"

    # ---------- low-level ----------
    def _request(self, method: str, url_or_path: str, **kw: Any) -> requests.Response:
        url = url_or_path if url_or_path.startswith("http") else f"{self.cfg.graph_base_url}{url_or_path}"
        refreshed = False
        for attempt in range(self.cfg.max_retries + 1):
            headers = {"Authorization": f"Bearer {self.auth.get_token()}", **kw.pop("headers", {})}
            resp = self._http.request(method, url, headers=headers, timeout=self.cfg.timeout, **kw)
            if resp.status_code == 401 and not refreshed:
                refreshed = True
                continue
            if resp.status_code in RETRY_STATUS and attempt < self.cfg.max_retries:
                wait = float(resp.headers.get("Retry-After", 2 ** attempt))
                time.sleep(min(wait, 60))
                continue
            if not resp.ok:
                raise GraphError(f"{method} {url} -> {resp.status_code}: {resp.text[:500]}")
            return resp
        raise GraphError(f"{method} {url}: retries exhausted")

    def _paged(self, path: str, params: dict[str, Any] | None = None) -> Iterator[dict[str, Any]]:
        url: str | None = path
        while url:
            data = self._request("GET", url, params=params if url == path else None).json()
            yield from data.get("value", [])
            url = data.get("@odata.nextLink")

    # ---------- folders ----------
    def find_folder_id(self, display_name: str) -> str:
        """Look up a top-level folder ID by name. Well-known names ('inbox', 'archive', ...) need no lookup."""
        safe = display_name.replace("'", "''")
        for f in self._paged(f"{self._user}/mailFolders", {"$filter": f"displayName eq '{safe}'"}):
            return f["id"]
        raise GraphError(f"Folder not found: {display_name}")

    # ---------- messages ----------
    def iter_messages(
        self,
        folder: str = "inbox",
        *,
        since: datetime | None = None,
        unread_only: bool = False,
        has_attachments: bool | None = None,
        subject_regex: str | None = None,
        sender_contains: str | None = None,
        page_size: int = 50,
        limit: int | None = None,
        select: str = DEFAULT_SELECT,
    ) -> Iterator[dict[str, Any]]:
        # receivedDateTime must lead the $filter when it's used in $orderby.
        filters = [f"receivedDateTime ge {_iso(since or EPOCH)}"]
        if unread_only:
            filters.append("isRead eq false")
        if has_attachments is not None:
            filters.append(f"hasAttachments eq {str(has_attachments).lower()}")
        params = {
            "$filter": " and ".join(filters),
            "$orderby": "receivedDateTime asc",
            "$top": page_size,
            "$select": select,
        }
        subject_re = re.compile(subject_regex, re.I) if subject_regex else None
        sender_q = sender_contains.lower() if sender_contains else None
        count = 0
        for msg in self._paged(f"{self._user}/mailFolders/{quote(folder)}/messages", params):
            if subject_re and not subject_re.search(msg.get("subject") or ""):
                continue
            if sender_q:
                addr = ((msg.get("from") or {}).get("emailAddress") or {}).get("address", "")
                if sender_q not in addr.lower():
                    continue
            yield msg
            count += 1
            if limit and count >= limit:
                return

    def mark_read(self, message_id: str, is_read: bool = True) -> None:
        self._request("PATCH", f"{self._user}/messages/{message_id}", json={"isRead": is_read})

    def move_message(self, message_id: str, destination: str) -> dict[str, Any]:
        """destination: well-known name ('archive', 'deleteditems') or a folder ID (see find_folder_id)."""
        return self._request("POST", f"{self._user}/messages/{message_id}/move", json={"destinationId": destination}).json()

    # ---------- attachments ----------
    def get_attachments(
        self,
        message_id: str,
        *,
        name_regex: str | None = None,
        extensions: tuple[str, ...] | list[str] | None = None,
        include_inline: bool = False,
    ) -> list[Attachment]:
        """Download file attachments. Metadata is listed first, then each match is fetched via /$value,
        which also works for large attachments (>3 MB)."""
        name_re = re.compile(name_regex, re.I) if name_regex else None
        exts = tuple(e.lower() if e.startswith(".") else f".{e.lower()}" for e in extensions) if extensions else None
        base = f"{self._user}/messages/{message_id}/attachments"
        out: list[Attachment] = []
        for meta in self._paged(base, {"$select": "id,name,contentType,size,isInline"}):
            if meta.get("@odata.type") != "#microsoft.graph.fileAttachment":
                continue
            if meta.get("isInline") and not include_inline:
                continue
            name = meta.get("name") or ""
            if name_re and not name_re.search(name):
                continue
            if exts and not name.lower().endswith(exts):
                continue
            content = self._request("GET", f"{base}/{meta['id']}/$value").content
            out.append(Attachment(meta["id"], message_id, name, meta.get("contentType", ""),
                                  meta.get("size", len(content)), bool(meta.get("isInline")), content))
        return out

    def iter_attachments(
        self,
        folder: str = "inbox",
        *,
        name_regex: str | None = None,
        extensions: tuple[str, ...] | list[str] | None = None,
        **message_filters: Any,
    ) -> Iterator[tuple[dict[str, Any], Attachment]]:
        """Convenience: yield (message, attachment) for every matching attachment in matching messages."""
        message_filters["has_attachments"] = True
        for msg in self.iter_messages(folder, **message_filters):
            for att in self.get_attachments(msg["id"], name_regex=name_regex, extensions=extensions):
                yield msg, att
