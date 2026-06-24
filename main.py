from msal import PublicClientApplication
import requests
import sqlite3
import time
from bs4 import BeautifulSoup


# ── Config ──────────────────────────────────────────────────────────────────

CLIENT_ID = "d8f7b98e-3120-4035-9fe2-588c00d7c74c"  
SCOPES    = ["Mail.Read"]
AUTHORITY = "https://login.microsoftonline.com/organizations"

GRAPH_URL = (
    "https://graph.microsoft.com/v1.0/me/messages"
    "?$top=20"
    "&$orderby=receivedDateTime desc"
    "&$select=id,from,subject,receivedDateTime,body"
)

POLL_INTERVAL = 30  # seconds


# ── Database ─────────────────────────────────────────────────────────────────

def setup_db():
    conn = sqlite3.connect("emails.db")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS emails (
            id       TEXT PRIMARY KEY,
            account  TEXT,
            sender   TEXT,
            subject  TEXT,
            received TEXT,
            message  TEXT
        )
    """)
    conn.commit()
    return conn


# ── Auth ──────────────────────────────────────────────────────────────────────

def login(label):
    """Run device flow login and return the MSAL app object."""
    app = PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"Device flow failed for {label}:\n{flow}")

    print(f"\n{'─'*50}")
    print(f"  Login → {label}")
    print(f"{'─'*50}")
    print(flow["message"])
    print()

    token = app.acquire_token_by_device_flow(flow)
    if "access_token" not in token:
        raise Exception(f"Token acquisition failed for {label}:\n{token}")

    print(f"  ✓ {label} logged in successfully.")
    return app


def get_access_token(app):
    """Return a valid access token, refreshing silently if needed."""
    cached_accounts = app.get_accounts()
    if cached_accounts:
        result = app.acquire_token_silent(SCOPES, account=cached_accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    raise Exception("Silent token refresh failed.")


# ── Email Fetching ────────────────────────────────────────────────────────────

def fetch_emails(app, label, conn):
    """Fetch latest emails for one account and store new ones in DB."""
    try:
        token = get_access_token(app)
    except Exception as e:
        print(f"  [{label}] Auth error: {e}")
        return 0

    headers = {"Authorization": f"Bearer {token}"}

    try:
        r    = requests.get(GRAPH_URL, headers=headers, timeout=10)
        data = r.json()
    except Exception as e:
        print(f"  [{label}] Request error: {e}")
        return 0

    if "error" in data:
        print(f"  [{label}] API error: {data['error'].get('message', data['error'])}")
        return 0

    new_count = 0

    for m in data.get("value", []):

        # Sender
        sender = ""
        if m.get("from") and m["from"].get("emailAddress"):
            sender = m["from"]["emailAddress"]["address"]

        # Body — strip HTML
        html    = m.get("body", {}).get("content", "")
        message = BeautifulSoup(html, "html.parser").get_text("\n", strip=True)

        # Insert (ignore duplicates)
        cur = conn.execute(
            "INSERT OR IGNORE INTO emails VALUES (?,?,?,?,?,?)",
            (
                m["id"],
                label,
                sender,
                m.get("subject", ""),
                m.get("receivedDateTime", ""),
                message,
            ),
        )

        if cur.rowcount > 0:
            print(f"\n  ✉  NEW EMAIL  [{label}]")
            print(f"     From   : {sender}")
            print(f"     Subject: {m.get('subject', '(no subject)')}")
            new_count += 1

    conn.commit()
    return new_count


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    conn = setup_db()

    app = login("Account 1")

    print(f"\n{'═'*50}")
    print("  Account ready. Starting sync loop...")
    print(f"{'═'*50}\n")

    while True:
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n[{timestamp}] Checking emails...\n")

        count = fetch_emails(app, "Account 1", conn)
        print(f"  Account 1: {count} new email(s)")

        print(f"\n  Next check in {POLL_INTERVAL}s...")
        time.sleep(POLL_INTERVAL)


if __name__ == "__main__":
    main()in()