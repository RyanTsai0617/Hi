#!/usr/bin/env python3
"""
Earnings Trigger Orchestrator — 完整流程入口。

Steps:
  1. Update earnings calendar (yfinance)
  2. Scan SEC EDGAR for new filings
  3. (Optional) Scan Gmail for earnings alerts
  4. Purge stale jobs
  5. Run pending jobs (new + retryable failed)

Usage:
  python main.py                  # full pipeline
  python main.py --scan-only      # only update calendar + detect triggers, no job execution
  python main.py --run-only       # only run pending jobs
  python main.py --status         # print job status summary
"""

import argparse
import sys
from datetime import datetime

from db import get_conn, init_db


def step_calendar():
    print("\n=== Step 1: Update Earnings Calendar ===")
    try:
        from calendar_watcher import update_calendar
        update_calendar()
        print("[calendar] Done")
    except Exception as e:
        print(f"[calendar] Error: {e}")


def step_sec():
    print("\n=== Step 2: Scan SEC EDGAR ===")
    try:
        from sec_trigger import detect_sec_jobs
        detect_sec_jobs()
    except Exception as e:
        print(f"[sec] Error: {e}")


def step_gmail():
    print("\n=== Step 3: Scan Gmail ===")
    try:
        from gmail_trigger import detect_gmail_jobs
        detect_gmail_jobs()
        print("[gmail] Done")
    except FileNotFoundError:
        print("[gmail] Skipped — data/gmail_token.json not found (run OAuth setup first)")
    except Exception as e:
        print(f"[gmail] Error: {e}")


def step_purge():
    print("\n=== Step 4: Purge Stale Jobs ===")
    from job_runner import purge_stale_jobs
    purge_stale_jobs()


def step_run(limit=3):
    print(f"\n=== Step 5: Run Jobs (limit={limit}) ===")
    from job_runner import run_jobs
    run_jobs(limit=limit)


def print_status():
    """Print a summary of job queue and upcoming earnings."""
    init_db()
    conn = get_conn()

    print("\n=== Job Queue Summary ===")
    for row in conn.execute("""
        SELECT status, count(*) as cnt FROM trigger_jobs GROUP BY status ORDER BY status
    """).fetchall():
        print(f"  {row['status']:12s}  {row['cnt']}")

    print("\n=== Upcoming Earnings (next 14 days) ===")
    today = datetime.now().strftime("%Y-%m-%d")
    for row in conn.execute("""
        SELECT ticker, expected_date, source
        FROM earnings_calendar
        WHERE expected_date >= ?
        ORDER BY expected_date ASC
        LIMIT 20
    """, (today,)).fetchall():
        delta = (datetime.strptime(row["expected_date"], "%Y-%m-%d") - datetime.now()).days
        marker = " *** TODAY ***" if delta <= 0 else f" (in {delta}d)"
        print(f"  {row['ticker']:6s}  {row['expected_date']}{marker}")

    print("\n=== Pending Jobs (new, top 10) ===")
    for row in conn.execute("""
        SELECT ticker, form_type, filing_date, trigger_score
        FROM trigger_jobs
        WHERE status = 'new'
        ORDER BY trigger_score DESC, filing_date DESC
        LIMIT 10
    """).fetchall():
        print(f"  {row['ticker']:6s}  {row['form_type'] or 'n/a':6s}  {row['filing_date'] or '?':12s}  score={row['trigger_score']}")

    print("\n=== Recently Failed (top 5) ===")
    for row in conn.execute("""
        SELECT ticker, form_type, filing_date, retry_count, error_message
        FROM trigger_jobs
        WHERE status = 'failed'
        ORDER BY detected_at DESC
        LIMIT 5
    """).fetchall():
        err = (row["error_message"] or "")[:80]
        print(f"  {row['ticker']:6s}  {row['form_type'] or 'n/a':6s}  retries={row['retry_count']}  {err}")

    conn.close()


def main():
    parser = argparse.ArgumentParser(description="Earnings Trigger Orchestrator")
    parser.add_argument("--scan-only", action="store_true", help="Only scan triggers, don't run jobs")
    parser.add_argument("--run-only", action="store_true", help="Only run pending jobs")
    parser.add_argument("--status", action="store_true", help="Print status summary")
    parser.add_argument("--limit", type=int, default=3, help="Max jobs to run per invocation")
    parser.add_argument("--skip-gmail", action="store_true", help="Skip Gmail trigger")
    args = parser.parse_args()

    print(f"[main] Earnings Trigger — {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    if args.status:
        print_status()
        return

    if args.run_only:
        step_purge()
        step_run(limit=args.limit)
        return

    # Full pipeline
    step_calendar()
    step_sec()
    if not args.skip_gmail:
        step_gmail()

    if not args.scan_only:
        step_purge()
        step_run(limit=args.limit)

    print("\n[main] Done.")


if __name__ == "__main__":
    main()
