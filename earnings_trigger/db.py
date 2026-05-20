import sqlite3
from config import DB_PATH, DATA_DIR

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS watchlist (
    ticker TEXT PRIMARY KEY,
    company_name TEXT,
    cik TEXT,
    ir_url TEXT,
    enabled INTEGER DEFAULT 1
);

CREATE TABLE IF NOT EXISTS earnings_calendar (
    ticker TEXT,
    expected_date TEXT,
    expected_time TEXT,
    source TEXT,
    last_checked_at TEXT,
    PRIMARY KEY (ticker, expected_date, source)
);

CREATE TABLE IF NOT EXISTS trigger_jobs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ticker TEXT NOT NULL,
    trigger_type TEXT NOT NULL,
    trigger_source TEXT NOT NULL,
    trigger_score INTEGER NOT NULL,
    dedupe_key TEXT NOT NULL UNIQUE,
    detected_at TEXT NOT NULL,
    filing_date TEXT,
    accession_number TEXT,
    form_type TEXT,
    status TEXT NOT NULL DEFAULT 'new',
    output_path TEXT,
    error_message TEXT,
    retry_count INTEGER NOT NULL DEFAULT 0
);
"""

MIGRATION_SQL = """
-- Add retry_count if missing (safe to re-run)
ALTER TABLE trigger_jobs ADD COLUMN retry_count INTEGER NOT NULL DEFAULT 0;
"""


def get_conn():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.executescript(SCHEMA_SQL)
        # Run migrations (ignore errors for already-applied)
        for stmt in MIGRATION_SQL.strip().split(";"):
            stmt = stmt.strip()
            if stmt and not stmt.startswith("--"):
                try:
                    conn.execute(stmt)
                except Exception:
                    pass  # column already exists
        conn.commit()
