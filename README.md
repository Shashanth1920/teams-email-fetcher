# Teams Email Reader

A Python script that polls a Microsoft 365 mailbox via the Microsoft Graph API and stores new emails in a local SQLite database.

## How It Works

1. Authenticates using the **OAuth 2.0 Device Code Flow** (MSAL) — no password stored.
2. Every 30 seconds, fetches the 20 most recent emails from the signed-in account.
3. Strips HTML from email bodies using BeautifulSoup.
4. Inserts new emails into `emails.db` (duplicates are silently ignored).
5. Prints a summary of any newly discovered emails to the console.

## Requirements

- Python 3.8+
- A registered Azure AD app with `Mail.Read` permission (delegated)

Install dependencies:

```bash
pip install msal requests beautifulsoup4
```

## Usage

```bash
python main.py
```

On first run, a device code prompt will appear. Visit the displayed URL, enter the code, and sign in with your Microsoft 365 account. The script then starts polling continuously.

## Configuration

Edit the constants at the top of [main.py](main.py):

| Variable | Description | Default |
|---|---|---|
| `CLIENT_ID` | Azure AD app (client) ID | `d8f7b98e-...` |
| `SCOPES` | Graph API permission scopes | `["Mail.Read"]` |
| `AUTHORITY` | MSAL authority URL | `organizations` tenant |
| `POLL_INTERVAL` | Seconds between checks | `30` |

## Database Schema

Emails are stored in `emails.db` (SQLite):

| Column | Type | Description |
|---|---|---|
| `id` | TEXT (PK) | Graph message ID |
| `account` | TEXT | Account label |
| `sender` | TEXT | Sender email address |
| `subject` | TEXT | Email subject |
| `received` | TEXT | ISO 8601 timestamp |
| `message` | TEXT | Plain-text body |

## Notes

- Tokens are cached in memory by MSAL and refreshed silently — re-authentication is only needed if the refresh token expires.
- The `emails.db` file is created automatically on first run.
