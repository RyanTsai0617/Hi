#!/usr/bin/env bash
set -e

cd "$(dirname "$0")"

echo "Creating virtual environment..."
python3 -m venv .venv

# Cross-platform activate
if [[ "$(uname -s)" == "Darwin" ]]; then
    source .venv/bin/activate
else
    source .venv/Scripts/activate
fi

echo "Installing dependencies..."
pip install --upgrade pip
pip install anthropic google-api-python-client google-auth-httplib2 google-auth-oauthlib yfinance requests

echo "Initializing database..."
python3 db.py

echo ""
echo "earnings_trigger 環境初始化完成"
echo ""
echo "Next steps:"
if [[ "$(uname -s)" == "Darwin" ]]; then
    echo "  1. source .venv/bin/activate"
else
    echo "  1. source .venv/Scripts/activate"
fi
echo "  2. python3 calendar_watcher.py    # update earnings calendar"
echo "  3. python3 sec_trigger.py         # scan SEC filings"
echo "  4. python3 job_runner.py          # run pending jobs"
