"""
Earnings Trigger Integration — 整合 SEC 申報偵測 + 深度分析 into market_news 郵件

Usage:
    from earnings_integration import (
        scan_sec_filings,       # SEC EDGAR 掃描
        get_recent_deep_analyses,  # 取得已完成深度分析
        build_sec_filings_html, # SEC 申報 HTML
        build_deep_analysis_html,  # 深度分析 HTML
        queue_deep_analysis,    # 排隊深度分析 job
    )
"""
import csv
import hashlib
import json
import logging
import os
import sqlite3
import html as html_mod
from datetime import datetime, timedelta
from pathlib import Path

import requests

log = logging.getLogger("market_news")

# ── Paths ──
SCRIPT_DIR = Path(__file__).resolve().parent
TRIGGER_DIR = None  # set in init()
TRIGGER_DB_PATH = None

# SEC API
SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
SEC_TRIGGER_FORMS = {"8-K", "10-Q", "10-K"}
SEC_FILING_URL = "https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={cik}&type={form}&dateb=&owner=include&count=10"
SEC_ARCHIVES_URL = "https://www.sec.gov/Archives/edgar/data/{cik}/{accession_formatted}/{primary_doc}"
SEC_HEADERS = {
    "User-Agent": "Cathay Securities earnings monitor ryan@cathaysec.com.tw",
    "Accept-Encoding": "gzip, deflate",
}

# Deep analysis watchlist (高 token 成本，限制範圍)
DEEP_ANALYSIS_CSV = None
_deep_watchlist = None


def init(trigger_base=None):
    """初始化 earnings_trigger 路徑。支援自動偵測。"""
    global TRIGGER_DIR, TRIGGER_DB_PATH, DEEP_ANALYSIS_CSV, _deep_watchlist

    if trigger_base:
        TRIGGER_DIR = Path(trigger_base)
    else:
        # Auto-detect: 同目錄 or OneDrive
        candidates = [
            SCRIPT_DIR / "earnings_trigger",
            Path.home() / "Library" / "CloudStorage" / "OneDrive-個人" / "桌面" / "Coding" / "earnings_trigger",
            Path.home() / "OneDrive" / "桌面" / "Coding" / "earnings_trigger",
            Path.home() / "Code" / "earnings_trigger",
        ]
        for c in candidates:
            if (c / "data").exists():
                TRIGGER_DIR = c
                break

    if TRIGGER_DIR is None:
        log.warning("earnings_trigger directory not found, SEC integration disabled")
        return False

    TRIGGER_DB_PATH = TRIGGER_DIR / "data" / "trigger.db"
    DEEP_ANALYSIS_CSV = TRIGGER_DIR / "data" / "us_earnings_watchlist.csv"

    # Load deep analysis watchlist
    _deep_watchlist = _load_deep_watchlist()
    log.info(f"earnings_trigger initialized: {TRIGGER_DIR}")
    log.info(f"  Deep analysis watchlist: {len(_deep_watchlist)} tickers")
    return True


def _load_deep_watchlist():
    """讀取深度分析 watchlist (CSV)"""
    if not DEEP_ANALYSIS_CSV or not DEEP_ANALYSIS_CSV.exists():
        return {}
    try:
        with open(DEEP_ANALYSIS_CSV, "r", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
        return {r["ticker"].upper(): r for r in rows if r.get("enabled", "1") == "1"}
    except Exception as e:
        log.warning(f"Failed to load deep watchlist: {e}")
        return {}


def _get_trigger_conn():
    """取得 trigger.db 連線"""
    if not TRIGGER_DB_PATH:
        return None
    TRIGGER_DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(TRIGGER_DB_PATH)
    conn.row_factory = sqlite3.Row
    # Ensure tables exist
    conn.executescript("""
        CREATE TABLE IF NOT EXISTS watchlist (
            ticker TEXT PRIMARY KEY,
            company_name TEXT,
            cik TEXT,
            ir_url TEXT,
            enabled INTEGER DEFAULT 1
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
            error_message TEXT
        );
    """)
    return conn


def _normalize_cik(cik):
    return str(cik).zfill(10)


def _make_dedupe_key(*parts):
    return hashlib.sha256("|".join(str(p) for p in parts).encode()).hexdigest()


# ══════════════════════════════════════════════════════════════════════════════
#  SEC EDGAR 掃描
# ══════════════════════════════════════════════════════════════════════════════

def _fetch_company_tickers():
    """SEC ticker → CIK mapping"""
    headers = {**SEC_HEADERS, "Host": "www.sec.gov"}
    r = requests.get(SEC_TICKERS_URL, headers=headers, timeout=30)
    r.raise_for_status()
    data = r.json()
    mapping = {}
    for _, item in data.items():
        mapping[item["ticker"].upper()] = {
            "cik": str(item["cik_str"]),
            "title": item["title"],
        }
    return mapping


def scan_sec_filings(watchlist_dict, lookback_days=3):
    """
    掃描 watchlist 中所有公司的 SEC 申報。
    回傳近期 filings 列表 + 新 trigger jobs 數量。

    Parameters:
        watchlist_dict: {ticker: company_name} — market_news 的 270 檔 watchlist
        lookback_days: 回溯天數

    Returns:
        list of dict: [{ticker, company, form_type, filing_date, accession_number, is_new}, ...]
    """
    if not TRIGGER_DIR:
        return []

    cutoff = (datetime.now() - timedelta(days=lookback_days)).strftime("%Y-%m-%d")
    now_iso = datetime.utcnow().isoformat() + "+00:00"

    # Get CIK mapping
    try:
        sec_map = _fetch_company_tickers()
    except Exception as e:
        log.warning(f"SEC tickers fetch failed: {e}")
        return []

    conn = _get_trigger_conn()
    if not conn:
        return []

    filings = []
    errors = 0

    # Only scan tickers that are in both watchlist AND have SEC data
    for ticker, company in watchlist_dict.items():
        if ticker.endswith((".HK", ".T")):  # Skip non-US
            continue
        sec_info = sec_map.get(ticker.upper())
        if not sec_info:
            continue

        cik = sec_info["cik"]

        try:
            url = SEC_SUBMISSIONS_URL.format(cik=_normalize_cik(cik))
            headers = {**SEC_HEADERS, "Host": "data.sec.gov"}
            r = requests.get(url, headers=headers, timeout=20)
            r.raise_for_status()
            data = r.json()

            recent = data.get("filings", {}).get("recent", {})
            forms = recent.get("form", [])
            filing_dates = recent.get("filingDate", [])
            accession_numbers = recent.get("accessionNumber", [])

            for form_type, filing_date, acc_num in zip(forms, filing_dates, accession_numbers):
                if form_type not in SEC_TRIGGER_FORMS:
                    continue
                if filing_date < cutoff:
                    continue

                dedupe_key = _make_dedupe_key(ticker, "sec", form_type, filing_date, acc_num)

                # Check if already in trigger_jobs
                existing = conn.execute(
                    "SELECT id, status FROM trigger_jobs WHERE dedupe_key = ?",
                    (dedupe_key,)
                ).fetchone()
                is_new = existing is None

                if is_new:
                    # Insert new trigger job (only for deep watchlist tickers)
                    if ticker.upper() in (_deep_watchlist or {}):
                        conn.execute("""
                            INSERT OR IGNORE INTO trigger_jobs
                            (ticker, trigger_type, trigger_source, trigger_score, dedupe_key,
                             detected_at, filing_date, accession_number, form_type, status)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'new')
                        """, (
                            ticker, "sec_filing", "market_news_scan", 5,
                            dedupe_key, now_iso, filing_date, acc_num, form_type,
                        ))

                filings.append({
                    "ticker": ticker.upper(),
                    "company": company,
                    "form_type": form_type,
                    "filing_date": filing_date,
                    "accession_number": acc_num,
                    "is_new": is_new,
                    "status": None if is_new else dict(existing).get("status"),
                })

        except Exception as e:
            errors += 1
            if errors <= 3:
                log.debug(f"  SEC scan {ticker}: {e}")
            if errors > 10:
                log.warning(f"  SEC scan: too many errors ({errors}), stopping early")
                break

    conn.commit()
    conn.close()

    # Sort by date desc
    filings.sort(key=lambda x: x["filing_date"], reverse=True)
    log.info(f"  SEC scan: {len(filings)} filings found ({errors} errors)")
    return filings


# ══════════════════════════════════════════════════════════════════════════════
#  Deep Analysis 結果讀取
# ══════════════════════════════════════════════════════════════════════════════

def get_recent_deep_analyses(days=7):
    """
    從 trigger_jobs 取得近期完成的深度分析。
    讀取輸出 .md 檔取得摘要。

    Returns:
        list of dict: [{ticker, filing_date, form_type, output_path, summary_excerpt}, ...]
    """
    if not TRIGGER_DIR:
        return []

    conn = _get_trigger_conn()
    if not conn:
        return []

    cutoff = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d")
    rows = conn.execute("""
        SELECT ticker, filing_date, form_type, output_path, detected_at
        FROM trigger_jobs
        WHERE status = 'done' AND filing_date >= ?
        ORDER BY filing_date DESC
    """, (cutoff,)).fetchall()
    conn.close()

    results = []
    for row in rows:
        excerpt = ""
        output_path = row["output_path"]
        if output_path and os.path.exists(output_path):
            try:
                with open(output_path, "r", encoding="utf-8") as f:
                    content = f.read()
                # Extract first 500 chars as excerpt
                # Try to find the conclusion section
                for marker in ["## 結論", "## Conclusion", "## 總結", "**結論**"]:
                    idx = content.find(marker)
                    if idx >= 0:
                        excerpt = content[idx:idx+500].strip()
                        break
                if not excerpt:
                    excerpt = content[:500].strip()
            except Exception:
                excerpt = "(無法讀取)"

        results.append({
            "ticker": row["ticker"],
            "filing_date": row["filing_date"],
            "form_type": row["form_type"],
            "output_path": output_path,
            "summary_excerpt": excerpt,
        })

    return results


def queue_deep_analysis(ticker, trigger_type="news_high_score", filing_date=None):
    """
    手動排隊深度分析 job（從 market_news 高分新聞觸發）。
    只有在 deep_watchlist 裡的 ticker 才會排隊。
    """
    if not TRIGGER_DIR or not _deep_watchlist:
        return False
    if ticker.upper() not in _deep_watchlist:
        return False

    conn = _get_trigger_conn()
    if not conn:
        return False

    now_iso = datetime.utcnow().isoformat() + "+00:00"
    filing_date = filing_date or datetime.now().strftime("%Y-%m-%d")
    dedupe_key = _make_dedupe_key(ticker, trigger_type, filing_date)

    try:
        conn.execute("""
            INSERT OR IGNORE INTO trigger_jobs
            (ticker, trigger_type, trigger_source, trigger_score, dedupe_key,
             detected_at, filing_date, form_type, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'new')
        """, (
            ticker.upper(), trigger_type, "market_news_auto", 7,
            dedupe_key, now_iso, filing_date, "news",
        ))
        conn.commit()
        conn.close()
        log.info(f"  Queued deep analysis: {ticker} ({trigger_type})")
        return True
    except Exception as e:
        log.debug(f"  Queue deep analysis failed: {ticker} {e}")
        conn.close()
        return False


# ══════════════════════════════════════════════════════════════════════════════
#  HTML 渲染
# ══════════════════════════════════════════════════════════════════════════════

_FORM_LABELS = {
    "8-K": ("重大事件", "#e74c3c"),
    "10-Q": ("季報", "#2980b9"),
    "10-K": ("年報", "#27ae60"),
}


def build_sec_filings_html(filings, max_show=15):
    """渲染 SEC 最新申報 HTML 區段"""
    if not filings:
        return ""

    h = """<div style="margin-top:22px">
<div style="background:linear-gradient(90deg,#1a5276,#2e86c1);color:white;padding:10px 16px;
border-radius:6px 6px 0 0;font-size:15px;font-weight:bold">📋 SEC 最新申報</div>
<div style="border:1px solid #d5dbdb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">"""

    for f in filings[:max_show]:
        label, color = _FORM_LABELS.get(f["form_type"], ("Other", "#95a5a6"))
        new_badge = ' <span style="background:#e74c3c;color:white;font-size:10px;padding:1px 5px;border-radius:3px">NEW</span>' if f["is_new"] else ""
        h += (
            f'<div style="margin-bottom:8px;padding:8px 12px;background:#f8f9fa;border-left:3px solid {color};border-radius:4px">'
            f'<span style="background:{color};color:white;font-size:11px;padding:2px 6px;border-radius:3px;font-weight:600">{f["form_type"]}</span> '
            f'<strong>{html_mod.escape(f["ticker"])}</strong> '
            f'<span style="color:#666">{html_mod.escape(f["company"])}</span>'
            f'{new_badge} '
            f'<span style="color:#999;font-size:12px;float:right">{f["filing_date"]}</span>'
            f'</div>\n'
        )

    total = len(filings)
    if total > max_show:
        h += f'<div style="color:#999;font-size:12px;text-align:center;padding:6px">... 另有 {total - max_show} 筆申報</div>'

    h += "</div></div>"
    return h


def build_deep_analysis_html(analyses):
    """渲染深度分析結果 HTML 區段"""
    if not analyses:
        return ""

    h = """<div style="margin-top:22px">
<div style="background:linear-gradient(90deg,#6c3483,#a569bd);color:white;padding:10px 16px;
border-radius:6px 6px 0 0;font-size:15px;font-weight:bold">🔬 深度財報分析（Claude）</div>
<div style="border:1px solid #d5dbdb;border-top:none;border-radius:0 0 6px 6px;padding:14px 16px">"""

    for a in analyses:
        label, color = _FORM_LABELS.get(a["form_type"], ("分析", "#8e44ad"))
        excerpt = a["summary_excerpt"]
        if len(excerpt) > 300:
            excerpt = excerpt[:300] + "..."

        h += (
            f'<div style="margin-bottom:14px;padding:12px 14px;background:#faf5ff;border-left:4px solid #8e44ad;border-radius:4px">'
            f'<div style="margin-bottom:6px">'
            f'<strong style="font-size:15px">{html_mod.escape(a["ticker"])}</strong> '
            f'<span style="background:{color};color:white;font-size:10px;padding:1px 5px;border-radius:3px">{a["form_type"]}</span> '
            f'<span style="color:#999;font-size:12px">{a["filing_date"]}</span>'
            f'</div>'
            f'<div style="font-size:13px;line-height:1.7;color:#4a235a;white-space:pre-line">{html_mod.escape(excerpt)}</div>'
            f'</div>\n'
        )

    h += "</div></div>"
    return h
