"""
SEC EDGAR trigger.
Scans company submissions JSON for recent 8-K / 10-Q / 10-K filings
and creates deduplicated trigger jobs.
"""

import csv
from datetime import datetime, timedelta
import requests

from db import get_conn, init_db
from config import (
    WATCHLIST_CSV, SEC_TICKERS_URL, SEC_SUBMISSIONS_URL,
    USER_AGENT, SEC_TRIGGER_FORMS, SEC_TRIGGER_SCORE,
)
from utils import utc_now_iso, make_dedupe_key, normalize_cik

# Only trigger filings from the last N days
LOOKBACK_DAYS = 7

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept-Encoding": "gzip, deflate",
    "Host": "www.sec.gov",
}


def load_watchlist():
    with open(WATCHLIST_CSV, "r", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def fetch_company_tickers():
    r = requests.get(SEC_TICKERS_URL, headers=HEADERS, timeout=30)
    r.raise_for_status()
    data = r.json()
    mapping = {}
    for _, item in data.items():
        mapping[item["ticker"].upper()] = {
            "cik": str(item["cik_str"]),
            "title": item["title"],
        }
    return mapping


def upsert_watchlist_with_cik():
    init_db()
    sec_map = fetch_company_tickers()
    rows = load_watchlist()

    with get_conn() as conn:
        for row in rows:
            ticker = row["ticker"].upper().strip()
            sec_info = sec_map.get(ticker)
            if not sec_info:
                continue
            conn.execute("""
                INSERT INTO watchlist (ticker, company_name, cik, ir_url, enabled)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(ticker) DO UPDATE SET
                    company_name=excluded.company_name,
                    cik=excluded.cik,
                    ir_url=excluded.ir_url,
                    enabled=excluded.enabled
            """, (
                ticker,
                row.get("company_name") or sec_info["title"],
                sec_info["cik"],
                row.get("ir_url", ""),
                int(row.get("enabled", "1")),
            ))
        conn.commit()


def fetch_recent_filings(cik: str):
    url = SEC_SUBMISSIONS_URL.format(cik=normalize_cik(cik))
    headers = {**HEADERS, "Host": "data.sec.gov"}
    r = requests.get(url, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json()


def detect_sec_jobs(lookback_days=None):
    init_db()
    upsert_watchlist_with_cik()
    now = utc_now_iso()
    cutoff = (datetime.now() - timedelta(days=lookback_days or LOOKBACK_DAYS)).strftime("%Y-%m-%d")

    with get_conn() as conn:
        watchlist = conn.execute(
            "SELECT * FROM watchlist WHERE enabled = 1"
        ).fetchall()

        new_count = 0
        for row in watchlist:
            ticker = row["ticker"]
            cik = row["cik"]
            try:
                data = fetch_recent_filings(cik)
                recent = data.get("filings", {}).get("recent", {})
                forms = recent.get("form", [])
                filing_dates = recent.get("filingDate", [])
                accession_numbers = recent.get("accessionNumber", [])

                for form_type, filing_date, accession_number in zip(
                    forms, filing_dates, accession_numbers
                ):
                    if form_type not in SEC_TRIGGER_FORMS:
                        continue
                    if filing_date < cutoff:
                        continue

                    dedupe_key = make_dedupe_key(
                        ticker, "sec", form_type, filing_date, accession_number
                    )

                    cursor = conn.execute("""
                        INSERT OR IGNORE INTO trigger_jobs
                        (ticker, trigger_type, trigger_source, trigger_score, dedupe_key,
                         detected_at, filing_date, accession_number, form_type, status)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'new')
                    """, (
                        ticker,
                        "sec_filing",
                        "sec_submissions_json",
                        SEC_TRIGGER_SCORE.get(form_type, 5),
                        dedupe_key,
                        now,
                        filing_date,
                        accession_number,
                        form_type,
                    ))
                    if cursor.rowcount > 0:
                        new_count += 1

                conn.commit()
            except Exception as e:
                print(f"[sec] {ticker} failed: {e}")

        print(f"[sec] Done — {new_count} new jobs created (cutoff: {cutoff})")


if __name__ == "__main__":
    detect_sec_jobs()
