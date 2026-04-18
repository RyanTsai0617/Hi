"""
Job runner.
Picks up 'new' and retryable 'failed' trigger jobs, invokes Claude earnings workflow.
Supports two modes: Claude Code CLI (full agentic) or API fallback.
"""

import subprocess
import sys
from datetime import datetime, timedelta

from db import get_conn, init_db
from config import (
    RUNNER_OUTPUT_BASE, BASE_DIR, MAX_RETRIES, JOB_RECENCY_DAYS,
)


def run_jobs(limit=3, include_retries=True):
    init_db()
    cutoff = (datetime.now() - timedelta(days=JOB_RECENCY_DAYS)).strftime("%Y-%m-%d")

    with get_conn() as conn:
        # Pick up new jobs + retryable failed jobs
        if include_retries:
            jobs = conn.execute("""
                SELECT * FROM trigger_jobs
                WHERE (status = 'new' OR (status = 'failed' AND retry_count < ?))
                  AND COALESCE(filing_date, detected_at) >= ?
                ORDER BY
                    CASE status WHEN 'new' THEN 0 ELSE 1 END,
                    trigger_score DESC,
                    detected_at ASC
                LIMIT ?
            """, (MAX_RETRIES, cutoff, limit)).fetchall()
        else:
            jobs = conn.execute("""
                SELECT * FROM trigger_jobs
                WHERE status = 'new'
                  AND COALESCE(filing_date, detected_at) >= ?
                ORDER BY trigger_score DESC, detected_at ASC
                LIMIT ?
            """, (cutoff, limit)).fetchall()

        if not jobs:
            print(f"[runner] No jobs to run (cutoff: {cutoff})")
            return

        for job in jobs:
            job_id = job["id"]
            ticker = job["ticker"]
            filing_date = job["filing_date"] or job["detected_at"][:10]
            retry = job["retry_count"] if job["retry_count"] else 0

            output_dir = RUNNER_OUTPUT_BASE / ticker
            output_dir.mkdir(parents=True, exist_ok=True)
            output_path = output_dir / f"{filing_date}_{ticker}_earnings.md"

            cmd = [
                sys.executable,
                str(BASE_DIR / "run_claude_earnings.py"),
                "--ticker", ticker,
                "--date", filing_date,
            ]

            retry_label = f" (retry {retry})" if retry > 0 else ""
            try:
                conn.execute(
                    "UPDATE trigger_jobs SET status = 'running' WHERE id = ?",
                    (job_id,),
                )
                conn.commit()

                print(f"[runner] Starting{retry_label}: {ticker} ({job['form_type'] or 'news'}) → {output_path}")
                result = subprocess.run(
                    cmd,
                    capture_output=True,
                    text=True,
                    timeout=660,
                )

                if result.returncode == 0 and output_path.exists():
                    conn.execute("""
                        UPDATE trigger_jobs
                        SET status = 'done', output_path = ?
                        WHERE id = ?
                    """, (str(output_path), job_id))
                    conn.commit()
                    print(f"[runner] Done: {ticker} → {output_path}")
                else:
                    error_msg = result.stderr[:500] if result.stderr else f"exit code {result.returncode}"
                    conn.execute("""
                        UPDATE trigger_jobs
                        SET status = 'failed', error_message = ?, retry_count = ?
                        WHERE id = ?
                    """, (error_msg, retry + 1, job_id))
                    conn.commit()
                    remaining = MAX_RETRIES - retry - 1
                    print(f"[runner] Failed: {ticker} — {error_msg[:200]} ({remaining} retries left)")

            except subprocess.TimeoutExpired:
                conn.execute("""
                    UPDATE trigger_jobs
                    SET status = 'failed', error_message = 'timeout (11min)', retry_count = ?
                    WHERE id = ?
                """, (retry + 1, job_id))
                conn.commit()
                print(f"[runner] Timeout: {ticker}")

            except Exception as e:
                conn.execute("""
                    UPDATE trigger_jobs
                    SET status = 'failed', error_message = ?, retry_count = ?
                    WHERE id = ?
                """, (str(e)[:500], retry + 1, job_id))
                conn.commit()
                print(f"[runner] Error: {ticker} — {e}")


def purge_stale_jobs(days=60):
    """Mark very old 'new' jobs as 'stale' so they don't clog the queue."""
    init_db()
    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    with get_conn() as conn:
        cursor = conn.execute("""
            UPDATE trigger_jobs
            SET status = 'stale'
            WHERE status = 'new'
              AND COALESCE(filing_date, detected_at) < ?
        """, (cutoff,))
        conn.commit()
        print(f"[runner] Purged {cursor.rowcount} stale jobs (older than {cutoff})")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=3)
    parser.add_argument("--purge", action="store_true", help="Purge stale jobs instead of running")
    parser.add_argument("--no-retry", action="store_true")
    args = parser.parse_args()

    if args.purge:
        purge_stale_jobs()
    else:
        run_jobs(limit=args.limit, include_retries=not args.no_retry)
