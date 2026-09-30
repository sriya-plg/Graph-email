"""Example: pull order-entry and invoice PDFs from a mailbox. Each project owns its routing rules."""
from datetime import datetime, timedelta, timezone
from pathlib import Path

from graph_mail import GraphMailClient, MailConfig

# Option A: from env vars (GRAPH_MAIL_TENANT_ID, GRAPH_MAIL_CLIENT_ID, GRAPH_MAIL_MAILBOX, ...)
cfg = MailConfig.from_env()

# Option B: explicit / from your own settings file
# cfg = MailConfig(tenant_id="...", client_id="...", mailbox="orders@company.com",
#                  auth_mode="client_secret", client_secret="...")

client = GraphMailClient(cfg)
since = datetime.now(timezone.utc) - timedelta(days=1)

# Project-specific classification lives here, not in the shared package.
RULES = {"order_entry": r"order|PO[-_ ]?\d+", "invoice": r"invoice|inv[-_ ]?\d+"}

for msg, att in client.iter_attachments("inbox", since=since, unread_only=True, extensions=[".pdf"]):
    kind = next((k for k, rx in RULES.items() if __import__("re").search(rx, att.name, __import__("re").I)), "other")
    path = att.save(Path("downloads") / kind)
    print(f"{kind:12} {msg['subject']!r} -> {path}")
    client.mark_read(msg["id"])
