"""
Earnings calendar updater.
Uses yfinance get_earnings_dates() as a scheduling aid.
Note: future dates from yfinance can be unstable — treat as advisory, not final trigger.
"""

import csv
from datetime import datetime
import yfinance as yf

from db import get_conn, init_db
from config import WATCHLIST_CSV
from utils import utc_now_iso


def load_watchlist():
    with open(WATCHLIST_CSV, "r", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def update_calendar():
    init_db()
    rows = load_watchlist()
    now = utc_now_iso()

    with get_conn() as conn:
        for row in rows:
            ticker = row["ticker"].upper().strip()
            try:
                tk = yf.Ticker(ticker)
                df = tk.get_earnings_dates(limit=8)
                if df is None or df.empty:
                    continue

                for idx, rec in df.iterrows():
                    dt = idx.to_pydatetime() if hasattr(idx, "to_pydatetime") else idx
                    expected_date = dt.strftime("%Y-%m-%d")
                    expected_time = ""
                    conn.execute("""
                        INSERT OR REPLACE INTO earnings_calendar
                        (ticker, expected_date, expected_time, source, last_checked_at)
                        VALUES (?, ?, ?, ?, ?)
                    """, (ticker, expected_date, expected_time, "yfinance", now))
            except Exception as e:
                print(f"[calendar] {ticker} failed: {e}")

        conn.commit()


if __name__ == "__main__":
    update_calendar()
