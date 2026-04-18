#!/usr/bin/env python3
"""
Claude earnings workflow runner.

Two modes:
  1. Claude Code CLI (default) — full agentic with WebSearch, NLM, Gmail, Drive
  2. API fallback — pure Anthropic SDK, no tools (legacy)

Usage:
  python3 run_claude_earnings.py --ticker NVDA --date 2026-04-11
  python3 run_claude_earnings.py --ticker NVDA --mode api  # force legacy mode
"""

import argparse
import json
import subprocess
import sys
from datetime import datetime
from pathlib import Path

from config import (
    WORKFLOW_PROMPT_PATH, RUNNER_OUTPUT_BASE, RUNNER_MODE, CLAUDE_CLI,
    CLAUDE_MODEL, CLAUDE_MAX_TURNS, CLAUDE_MAX_BUDGET_USD, CLAUDE_ALLOWED_TOOLS,
)


def load_workflow_prompt() -> str:
    with open(WORKFLOW_PROMPT_PATH, "r", encoding="utf-8") as f:
        return f.read()


def run_claude_cli(ticker: str, as_of: str, output_path: Path) -> bool:
    """Run via Claude Code CLI — full agentic mode with MCP tools."""
    system_prompt = load_workflow_prompt()

    user_prompt = (
        f"請以以下 ticker 與日期作為本輪分析對象，完全依照 workflow 承作完整財報分析：\n\n"
        f"- Ticker: {ticker}\n"
        f"- As-of date: {as_of}\n\n"
        f"完成後將最終 Markdown 報告存到：{output_path}\n\n"
        f"如果在該日期前尚未公布最新財報，請輸出 pending 版 Markdown（說明預計財報日期與關注重點）。"
    )

    cmd = [
        CLAUDE_CLI,
        "-p", user_prompt,
        "--system-prompt", system_prompt,
        "--model", CLAUDE_MODEL,
        "--max-turns", str(CLAUDE_MAX_TURNS),
        "--allowedTools", ",".join(CLAUDE_ALLOWED_TOOLS),
        "--max-budget-usd", str(CLAUDE_MAX_BUDGET_USD),
        "--output-format", "text",
        "--permission-mode", "bypassPermissions",
    ]

    print(f"[runner] Claude CLI mode: {ticker} → {output_path}")
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        timeout=600,  # 10 min max
    )

    if result.returncode != 0:
        print(f"[runner] Claude CLI stderr: {result.stderr[:500]}", file=sys.stderr)
        # If CLI produced output despite error, still save it
        if result.stdout.strip():
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(result.stdout, encoding="utf-8")
            return True
        return False

    # The CLI agentic mode writes the file directly via Write tool.
    # But if it printed output instead, save it.
    if not output_path.exists() and result.stdout.strip():
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(result.stdout, encoding="utf-8")

    return output_path.exists()


def run_api_fallback(ticker: str, as_of: str, output_path: Path) -> bool:
    """Run via pure Anthropic API — no tools, single completion."""
    try:
        from anthropic import Anthropic
    except ImportError:
        print("[runner] anthropic SDK not installed", file=sys.stderr)
        return False

    system_prompt = load_workflow_prompt()
    user_prompt = (
        f"請以以下 ticker 與日期作為本輪分析對象，完全依照 workflow 承作完整財報分析：\n\n"
        f"- Ticker: {ticker}\n"
        f"- As-of date: {as_of}\n\n"
        f"如果在該日期前尚未公布最新財報，請輸出 pending 版 Markdown（說明預計財報日期與關注重點）。"
    )

    client = Anthropic()
    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=8000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_prompt}],
    )

    parts = []
    for block in resp.content:
        if block.type == "text":
            parts.append(block.text)
    text = "".join(parts)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(text, encoding="utf-8")
    print(f"[runner] API mode: {ticker} → {output_path} ({len(text)} chars)")
    return True


def main():
    parser = argparse.ArgumentParser(description="Run Claude earnings analysis.")
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--date", required=False, help="YYYY-MM-DD, default=today")
    parser.add_argument("--mode", choices=["claude_cli", "api", "auto"], default="auto",
                        help="Runner mode (default: auto-detect)")
    args = parser.parse_args()

    ticker = args.ticker.upper()
    as_of = args.date or datetime.today().strftime("%Y-%m-%d")

    output_dir = RUNNER_OUTPUT_BASE / ticker
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"{as_of}_{ticker}_earnings.md"

    mode = args.mode if args.mode != "auto" else RUNNER_MODE
    print(f"[runner] mode={mode}, ticker={ticker}, date={as_of}")

    if mode == "claude_cli":
        ok = run_claude_cli(ticker, as_of, output_path)
    else:
        ok = run_api_fallback(ticker, as_of, output_path)

    if ok:
        print(f"[runner] SUCCESS: {output_path}")
    else:
        print(f"[runner] FAILED: {ticker}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
