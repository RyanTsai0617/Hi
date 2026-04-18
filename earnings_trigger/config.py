import platform
import shutil
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
LOG_DIR = BASE_DIR / "logs"
PROMPT_DIR = BASE_DIR / "prompts"

DB_PATH = DATA_DIR / "trigger.db"
WATCHLIST_CSV = DATA_DIR / "us_earnings_watchlist.csv"
WORKFLOW_PROMPT_PATH = PROMPT_DIR / "earnings_workflow_with_nlm.md"

SEC_TICKERS_URL = "https://www.sec.gov/files/company_tickers.json"
SEC_SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"

USER_AGENT = "Cathay Securities earnings monitor ryan@cathaysec.com.tw"

SEC_TRIGGER_FORMS = {"8-K", "10-Q", "10-K"}
SEC_TRIGGER_SCORE = {
    "8-K": 5,
    "10-Q": 5,
    "10-K": 5,
}

# --- Platform-aware paths ---
_system = platform.system()
if _system == "Darwin":
    RUNNER_OUTPUT_BASE = Path.home() / "Desktop" / "財報"
else:
    RUNNER_OUTPUT_BASE = Path.home() / "OneDrive" / "桌面" / "財報"

# --- Job settings ---
MAX_RETRIES = 2  # retry failed jobs up to N times
JOB_RECENCY_DAYS = 30  # only run jobs with filing_date within N days

# --- Runner mode ---
# "claude_cli" = full agentic (WebSearch, NLM, Gmail, Drive)
# "api"        = pure API call (no tools, legacy fallback)
CLAUDE_CLI = shutil.which("claude")
RUNNER_MODE = "claude_cli" if CLAUDE_CLI else "api"

# Claude CLI settings
CLAUDE_MODEL = "sonnet"
CLAUDE_MAX_TURNS = 50
CLAUDE_MAX_BUDGET_USD = 2.0  # per job safety cap
CLAUDE_ALLOWED_TOOLS = [
    "WebSearch", "WebFetch",
    "Read", "Write", "Edit", "Bash", "Glob", "Grep", "Agent",
    "mcp__notebooklm-mcp__*",
    "mcp__24aa53c4-9fa2-4975-b489-14ed2cbb1277__gmail_search_messages",
    "mcp__24aa53c4-9fa2-4975-b489-14ed2cbb1277__gmail_read_message",
    "mcp__c1fc4002-5f49-5f9d-a4e5-93c4ef5d6a75__google_drive_search",
    "mcp__c1fc4002-5f49-5f9d-a4e5-93c4ef5d6a75__google_drive_fetch",
]

# Legacy API fallback
RUN_CLAUDE_SCRIPT = BASE_DIR / "run_claude_earnings.py"
_python = "python3" if _system == "Darwin" else "python"
CLAUDE_API_CMD_TEMPLATE = f'{_python} "{{script}}" --ticker {{ticker}} --date {{date}} > "{{outfile}}"'
