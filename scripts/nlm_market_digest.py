#!/usr/bin/env python3
"""
NLM Market Digest — 自動整合 market_news 爬蟲結果至 NotebookLM
讀取 daily_report_data.json → 建 NLM notebook → 餵入 3 個 source → 查詢整合摘要

可獨立執行，也可從 market_news_crawler.py 末端呼叫。
"""

import json
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

# ── 路徑 ──────────────────────────────────────────────────────────────
PRODUCTION_DIR = Path.home() / "jobs" / "production" / "market_news"
JSON_PATH = PRODUCTION_DIR / "daily_report_data.json"
NLM_BIN = Path.home() / ".local" / "bin" / "nlm"

# ── NLM 查詢 prompt ──────────────────────────────────────────────────
NLM_QUERY_PROMPT = """請整合今日股票、債券、財報三大市場資料，產出一份「每日市場整合摘要」，格式如下：

1. 【今日關鍵主題】（3 個 bullet points，每個 50 字內）
2. 【股債聯動分析】（股市與債市今日走勢如何互相影響？通膨/地緣政治對兩者的差異影響）
3. 【重點個股與財報前瞻】（哪些個股因今日事件值得關注？即將公布財報的公司與市場的關聯）
4. 【明日觀察重點】（投資人明天該注意什麼？）

語言：繁體中文，專業但簡潔，500 字以內。"""


def load_json(path: Path = JSON_PATH) -> dict:
    """讀取 daily_report_data.json"""
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def build_equity_source(data: dict) -> str:
    """組裝股票市場 text source"""
    date_display = data.get("date_display", "")
    eq = data["equity_market"]
    lines = [f"# 海外股票市場數據 {date_display}\n"]

    lines.append("## 主要指數")
    for idx in eq.get("indices", []):
        lines.append(f"- {idx['name']}: {idx['close']:,.2f} ({idx['change_pct']:+.2f}%)")

    lines.append("\n## 利率 / 商品 / 匯率")
    for rc in eq.get("rates_commodities", []):
        lines.append(f"- {rc['name']}: {rc['close']:,.3f} ({rc['change_pct']:+.2f}%)")

    lines.append("\n## 類股表現")
    for s in eq.get("sectors", []):
        lines.append(f"- {s['name']}: {s['close']:,.2f} ({s['change_pct']:+.2f}%)")

    lines.append("\n## AI 市場摘要")
    lines.append(data.get("equity_full_summary", data.get("equity_summary", "")))

    lines.append("\n## 重要新聞（依重要性排序）")
    for i, a in enumerate(data.get("equity_articles", []), 1):
        lines.append(f"\n### {i}. {a['title']} [重要性: {a.get('score', '?')}/10]")
        lines.append(f"來源：{a.get('source', 'N/A')}")
        lines.append(f"摘要：{a.get('summary_zh', '')}")

    return "\n".join(lines)


def build_bond_source(data: dict) -> str:
    """組裝債券市場 text source"""
    date_display = data.get("date_display", "")
    bd = data["bond_market"]
    lines = [f"# 海外債券市場數據 {date_display}\n"]

    lines.append("## 美國公債殖利率")
    for t in bd.get("us_treasury", []):
        lines.append(f"- {t['name']}: {t['close']:.3f}% ({t['change_pct']:+.1f} bp)")

    spread = bd.get("spread_10y3m", 0)
    lines.append(f"\n10Y-3M 利差: {spread:.1f} bp")

    lines.append("\n## 公司債 / 投資級 / 高收益")
    for c in bd.get("corporate", []):
        lines.append(f"- {c['name']}: {c['close']:.2f} ({c['change_pct']:+.2f}%)")

    lines.append("\n## 新興市場債")
    for e in bd.get("emerging", []):
        lines.append(f"- {e['name']}: {e['close']:.2f} ({e['change_pct']:+.2f}%)")

    lines.append("\n## AI 債券摘要")
    lines.append(data.get("bond_full_summary", data.get("bond_summary", "")))

    lines.append("\n## 重要新聞（依重要性排序）")
    for i, a in enumerate(data.get("bond_articles", []), 1):
        lines.append(f"\n### {i}. {a['title']} [重要性: {a.get('score', '?')}/10]")
        lines.append(f"來源：{a.get('source', 'N/A')}")
        lines.append(f"摘要：{a.get('summary_zh', '')}")

    return "\n".join(lines)


def build_earnings_source(data: dict) -> str:
    """組裝財報行事曆 text source"""
    date_display = data.get("date_display", "")
    lines = [f"# 財報行事曆 {date_display}\n"]

    lines.append("## 近期已公布財報")
    for e in data.get("earnings_reported", []):
        lines.append(f"- {e['company']}: {e.get('summary', '')}")

    lines.append("\n## 即將公布財報")
    for e in data.get("earnings_upcoming", []):
        lines.append(f"- {e['company']} ({e.get('date', '')}): {e.get('summary', '')}")

    lines.append("\n## AI 財報摘要")
    lines.append(data.get("earnings_summary", ""))

    return "\n".join(lines)


def nlm_cmd(*args, timeout: int = 120) -> str:
    """執行 nlm CLI 指令，回傳 stdout"""
    cmd = [str(NLM_BIN)] + list(args)
    result = subprocess.run(
        cmd, capture_output=True, text=True, timeout=timeout
    )
    if result.returncode != 0:
        raise RuntimeError(f"nlm failed: {result.stderr.strip()}")
    return result.stdout.strip()


def create_notebook(title: str) -> str:
    """建立 notebook，回傳 notebook_id"""
    import re
    output = nlm_cmd("notebook", "create", title)
    # 格式: "✓ Created notebook: ...\n  ID: <uuid>"
    match = re.search(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", output)
    if match:
        return match.group()
    raise RuntimeError(f"Cannot parse notebook ID from: {output}")


def add_source_file(notebook_id: str, file_path: str, title: str = "") -> None:
    """上傳檔案作為 source，等待處理完成"""
    args = ["source", "add", notebook_id, "--file", file_path, "--wait"]
    if title:
        args += ["--title", title]
    nlm_cmd(*args, timeout=180)


def query_notebook(notebook_id: str, question: str) -> str:
    """查詢 notebook，回傳純文字回答（去除 JSON/citations）"""
    raw = nlm_cmd("query", "notebook", notebook_id, question, timeout=180)
    # nlm CLI 輸出可能是 JSON（含 answer + citations），只取 answer
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            # 嘗試 parsed["value"]["answer"] 或 parsed["answer"]
            if "value" in parsed and "answer" in parsed["value"]:
                return parsed["value"]["answer"]
            if "answer" in parsed:
                return parsed["answer"]
    except (json.JSONDecodeError, KeyError):
        pass
    return raw


def run(json_path: Path = JSON_PATH) -> dict:
    """
    主流程：JSON → NLM notebook → 整合摘要

    Returns:
        dict with keys: notebook_id, notebook_url, digest
    """
    print("[NLM] 讀取 daily_report_data.json ...")
    data = load_json(json_path)
    date_str = data.get("date_str", datetime.now().strftime("%Y%m%d"))
    date_display = data.get("date_display", date_str)

    # 建立 text source 暫存檔
    sources = {
        "equity": ("海外股票市場", build_equity_source(data)),
        "bond": ("海外債券市場", build_bond_source(data)),
        "earnings": ("財報行事曆", build_earnings_source(data)),
    }

    tmpdir = tempfile.mkdtemp(prefix="nlm_market_")
    source_files = {}
    for key, (label, text) in sources.items():
        fpath = Path(tmpdir) / f"{key}_{date_str}.md"
        fpath.write_text(text, encoding="utf-8")
        source_files[key] = (label, str(fpath))
        print(f"[NLM] {label}: {len(text)} chars → {fpath.name}")

    # 建 notebook
    title = f"海外市場日報 {date_display}"
    print(f"[NLM] 建立 notebook: {title}")
    notebook_id = create_notebook(title)
    notebook_url = f"https://notebooklm.google.com/notebook/{notebook_id}"
    print(f"[NLM] notebook_id: {notebook_id}")

    # 上傳 sources
    for key, (label, fpath) in source_files.items():
        print(f"[NLM] 上傳 {label} ...")
        add_source_file(notebook_id, fpath, title=f"{label} {date_display}")

    # 等一下讓 NLM 處理完
    time.sleep(3)

    # 查詢整合摘要
    print("[NLM] 查詢整合摘要 ...")
    digest = query_notebook(notebook_id, NLM_QUERY_PROMPT)
    print(f"[NLM] 摘要完成 ({len(digest)} chars)")

    # 儲存摘要
    digest_path = PRODUCTION_DIR / f"nlm_digest_{date_str}.txt"
    digest_path.write_text(digest, encoding="utf-8")
    print(f"[NLM] 摘要已存至 {digest_path}")

    return {
        "notebook_id": notebook_id,
        "notebook_url": notebook_url,
        "digest": digest,
        "digest_path": str(digest_path),
    }


if __name__ == "__main__":
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else JSON_PATH
    result = run(path)
    print(f"\n{'='*60}")
    print(f"Notebook: {result['notebook_url']}")
    print(f"Digest:   {result['digest_path']}")
    print(f"{'='*60}")
    print(result["digest"])
