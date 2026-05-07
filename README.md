# Refresh Token Mail Web

A small self-hosted Python web tool for managing Microsoft account refresh tokens and reading recent verification emails.

## Features

- Single-file Python HTTP server based on the standard library
- SQLite-backed account storage
- Microsoft OAuth refresh-token exchange
- Microsoft Graph Mail.Read preferred for verification-code email lookup
- IMAP XOAUTH2 fallback
- Saved accounts and batch import UI
- Tests for OAuth scope, Graph/IMAP mail helpers, page rendering, and security helper behavior

## Security Notice

This project handles highly sensitive data such as refresh tokens and optional mailbox passwords.

Do **not** commit or publish:

- `.env`
- `app.db`
- refresh tokens
- access tokens
- account passwords
- production logs containing secrets

Use a private server, bind to localhost by default, and place it behind your own authenticated reverse proxy if exposing it externally.

## Configuration

Environment variables:

- `RTWEB_DB`: SQLite database path, defaults to `app.db`
- `RTWEB_ADMIN_PASSWORD` or `RTWEB_PASSWORD`: admin password, defaults to `change-me`
- `RTWEB_SESSION_SECRET`: session-signing secret, random by default per process
- `RTWEB_HOST`: bind host, defaults to `127.0.0.1`
- `RTWEB_PORT`: bind port, defaults to `8020`
- `RTWEB_COOKIE_SECURE`: set to `0` only for local plain-HTTP development

Example:

```bash
export RTWEB_PASSWORD='change-this'
export RTWEB_SESSION_SECRET='use-a-long-random-secret'
python3 app.py
```

## Tests

```bash
pytest tests -q
```

## License

Internal utility code. Add a license before accepting external contributions.
