#!/usr/bin/env python3
"""
Gmail earnings trigger.
Scans recent Gmail messages for earnings-related keywords from IR / broker senders.
Creates deduplicated trigger_jobs entries.

Requires: data/gmail_token.json (from Gmail API OAuth quickstart).

Setup:
  1. Go to https://console.cloud.google.com/apis/credentials
  2. Create OAuth 2.0 Client ID (Desktop application)
  3. Download client_secret.json → earnings_trigger/data/
  4. Run: python gmail_trigger.py --setup
     This opens a browser for OAuth consent, then saves token to data/gmail_token.json
  5. After that, gmail_trigger.py can run headlessly.
"""

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from db import get_conn, init_db
from utils import utc_now_iso, make_dedupe_key

GMAIL_TOKEN_PATH = "data/gmail_token.json"
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

EARNINGS_KEYWORDS = [
    "earnings call",
    "quarterly results",
    "earnings release",
    "results for the quarter",
    "財報",
]

SENDER_HINTS = [
    "@investorrelations.",
    "@ir.",
    "no-reply@",
    "notification@",
]

TRIGGER_SCORE_GMAIL = 3


def get_gmail_service():
    creds = Credentials.from_authorized_user_file(GMAIL_TOKEN_PATH, GMAIL_SCOPES)
    service = build("gmail", "v1", credentials=creds)
    return service


def list_recent_messages(service, query: str, max_results: int = 50):
    results = (
        service.users()
        .messages()
        .list(userId="me", q=query, maxResults=max_results)
        .execute()
    )
    return results.get("messages", [])


def fetch_message(service, msg_id: str):
    return (
        service.users()
        .messages()
        .get(
            userId="me",
            id=msg_id,
            format="metadata",
            metadataHeaders=["Subject", "From", "Date"],
        )
        .execute()
    )


def detect_gmail_jobs():
    init_db()
    service = get_gmail_service()
    now = utc_now_iso()

    query = "newer_than:1d (earnings OR results OR 財報)"
    messages = list_recent_messages(service, query=query, max_results=100)

    with get_conn() as conn:
        watchlist_tickers = [
            row["ticker"]
            for row in conn.execute(
                "SELECT ticker FROM watchlist WHERE enabled = 1"
            ).fetchall()
        ]

        for msg in messages:
            msg_data = fetch_message(service, msg["id"])
            headers = {
                h["name"]: h["value"]
                for h in msg_data.get("payload", {}).get("headers", [])
            }
            subject = headers.get("Subject", "")
            sender = headers.get("From", "")

            text = f"{subject} {sender}".lower()
            if not any(k.lower() in text for k in EARNINGS_KEYWORDS):
                continue

            # Match ticker from subject against watchlist
            ticker = None
            for tk in watchlist_tickers:
                if tk in subject.upper():
                    ticker = tk
                    break
            if not ticker:
                continue

            dedupe_key = make_dedupe_key(ticker, "gmail", msg["id"])
            conn.execute(
                """
                INSERT OR IGNORE INTO trigger_jobs
                (ticker, trigger_type, trigger_source, trigger_score, dedupe_key,
                 detected_at, status)
                VALUES (?, ?, ?, ?, ?, ?, 'new')
            """,
                (
                    ticker,
                    "gmail_alert",
                    "gmail",
                    TRIGGER_SCORE_GMAIL,
                    dedupe_key,
                    now,
                ),
            )
        conn.commit()


def setup_oauth():
    """Interactive OAuth setup — run once to create gmail_token.json."""
    from google_auth_oauthlib.flow import InstalledAppFlow
    from pathlib import Path

    client_secret = Path("data/client_secret.json")
    if not client_secret.exists():
        print(f"Error: {client_secret} not found.")
        print("Download from Google Cloud Console → OAuth 2.0 Client IDs → Desktop app")
        return

    flow = InstalledAppFlow.from_client_secrets_file(str(client_secret), GMAIL_SCOPES)
    creds = flow.run_local_server(port=0)

    import json
    token_data = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes,
    }
    with open(GMAIL_TOKEN_PATH, "w") as f:
        json.dump(token_data, f, indent=2)
    print(f"Token saved to {GMAIL_TOKEN_PATH}")


if __name__ == "__main__":
    import sys
    if "--setup" in sys.argv:
        setup_oauth()
    else:
        detect_gmail_jobs()
