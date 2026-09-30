# graph-mail

Reusable Microsoft Graph mailbox reader: list messages, download attachments, mark read, move.
One package, many projects; each project supplies its own config and its own attachment-routing rules.

## Install (from your internal index / git)
    pip install graph-mail            # or: pip install -e path/to/graph_mail_connector

## Config
Per mailbox/app registration. Either env vars (prefix is customizable, so several mailboxes can coexist):

| Var (prefix `GRAPH_MAIL_`) | Notes |
|---|---|
| `TENANT_ID`, `CLIENT_ID` | from the Entra app registration |
| `MAILBOX` | user email/UPN; `me` allowed for delegated |
| `AUTH_MODE` | `client_secret` (default) or `delegated` |
| `CLIENT_SECRET` | required for `client_secret` |
| `TOKEN_CACHE_PATH` | delegated only; where the refresh-token cache lives |
| `DELEGATED_SCOPES` | comma list, default `Mail.Read,Mail.ReadWrite` |

    MailConfig.from_env(prefix="ORDERS_MAIL_")   # a second mailbox in the same process
    MailConfig(tenant_id=..., client_id=..., mailbox=..., auth_mode="delegated", token_cache_path=Path(".cache/orders.json"))

## Entra setup
- **client_secret**: add *Application* permission `Mail.Read` (+ `Mail.ReadWrite` if you mark read/move), grant admin consent,
  then **restrict to specific mailboxes** with an Exchange Application Access Policy or RBAC for Applications;
  otherwise the app can read every mailbox in the tenant.
- **delegated**: add *Delegated* `Mail.Read`/`Mail.ReadWrite`, enable *Allow public client flows*. First run does a device-code sign-in
  (`DelegatedAuth.login()`), later runs refresh silently from `TOKEN_CACHE_PATH`. Use `Mail.Read.Shared` for shared mailboxes.

## Usage
See `examples/fetch_orders_and_invoices.py`.

    client = GraphMailClient(MailConfig.from_env())
    for msg in client.iter_messages("inbox", since=dt, unread_only=True, subject_regex="invoice"): ...
    atts = client.get_attachments(msg["id"], extensions=[".pdf", ".xlsx"])

## Notes
- Retries 429/5xx honoring `Retry-After`; follows paging; attachments fetched via `/$value` so >3 MB files work.
- Only file attachments are returned (item/reference attachments skipped).
- Not included yet: delta queries (incremental sync), webhook subscriptions, sending mail.
