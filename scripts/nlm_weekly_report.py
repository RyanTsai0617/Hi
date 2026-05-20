#!/usr/bin/env python3
"""
NLM Weekly Report — 自動化週報 NotebookLM 投影片生成

流程：
  1. 讀取 YYYYMMDD/ 資料夾中的定稿 Word + 源 PDF
  2. 建 NLM notebook（總經 B 或 投資策略 C）
  3. 上傳 Word + PDF 作為 source
  4. 讀取 NLM 提示詞檔案
  5. 用 studio slides create 生成投影片
  6. 輪詢直到完成
  7. 下載 PDF

用法：
  python3 nlm_weekly_report.py 20260406 總經
  python3 nlm_weekly_report.py 20260406 投資策略
  python3 nlm_weekly_report.py 20260406 all        # 兩個都跑
"""

import json
import re
import subprocess
import sys
import time
from pathlib import Path

# ── 路徑設定 ──────────────────────────────────────────────────────────
import platform

_IS_WIN = platform.system() == "Windows"
NLM_BIN = "nlm" if _IS_WIN else str(Path.home() / ".local" / "bin" / "nlm")

# 週報工作資料夾的搜尋路徑（依優先順序，跨平台）
def _build_work_roots() -> list[Path]:
    home = Path.home()
    roots = []
    if _IS_WIN:
        # Windows OneDrive 常見路徑
        for onedrive in [home / "OneDrive" / "桌面" / "每周報告",
                         home / "OneDrive - 個人" / "桌面" / "每周報告"]:
            if onedrive.parent.exists():
                roots.append(onedrive)
    else:
        # macOS OneDrive
        roots.append(home / "Library" / "CloudStorage" / "OneDrive-個人" / "桌面" / "每周報告")
    roots += [home / "Desktop" / "每周報告", home / "Desktop", home / "Documents"]
    return roots

WORK_ROOTS = _build_work_roots()


def find_work_folder(date_str: str) -> Path:
    """尋找 YYYYMMDD 工作資料夾"""
    for root in WORK_ROOTS:
        # 直接在 root 下找
        candidate = root / date_str
        if candidate.is_dir():
            return candidate
        # 在子目錄找
        for sub in root.iterdir():
            if sub.is_dir():
                candidate = sub / date_str
                if candidate.is_dir():
                    return candidate
    raise FileNotFoundError(f"找不到工作資料夾 {date_str}，搜尋路徑: {[str(r) for r in WORK_ROOTS]}")


def nlm_cmd(*args, timeout: int = 300) -> str:
    """執行 nlm CLI"""
    cmd = [str(NLM_BIN)] + list(args)
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
    if result.returncode != 0:
        raise RuntimeError(f"nlm failed: {result.stderr.strip()}")
    return result.stdout.strip()


def parse_notebook_id(output: str) -> str:
    """從 nlm 輸出中提取 UUID"""
    match = re.search(r"[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}", output)
    if match:
        return match.group()
    raise RuntimeError(f"Cannot parse notebook ID from: {output}")


def wait_for_slides(notebook_id: str, max_wait: int = 600, interval: int = 15) -> dict:
    """輪詢等待投影片生成完成"""
    elapsed = 0
    while elapsed < max_wait:
        try:
            raw = nlm_cmd("status", "artifacts", notebook_id)
        except RuntimeError:
            time.sleep(interval)
            elapsed += interval
            continue
        try:
            data = json.loads(raw)
            artifacts = data if isinstance(data, list) else data.get("artifacts", [])
            for a in artifacts:
                if a.get("type") in ("slide_deck", "slides"):
                    st = a.get("status", "")
                    if st == "completed":
                        return a
                    elif st == "failed":
                        raise RuntimeError(f"投影片生成失敗: {a}")
        except json.JSONDecodeError:
            if "completed" in raw.lower():
                return {"status": "completed", "raw": raw}
        time.sleep(interval)
        elapsed += interval
        print(f"  [等待中 {elapsed}s / {max_wait}s]")
    raise TimeoutError(f"投影片生成超時 ({max_wait}s)")


def process_path(date_str: str, path_type: str, work_folder: Path):
    """
    處理單一路徑（總經 or 投資策略）

    Args:
        date_str: YYYYMMDD
        path_type: "總經" or "投資策略"
        work_folder: YYYYMMDD 工作資料夾
    """
    sub_folder = work_folder / path_type
    if not sub_folder.is_dir():
        raise FileNotFoundError(f"資料夾不存在: {sub_folder}")

    # 找 Word 定稿
    word_files = list(sub_folder.glob("*週報*.docx")) + list(sub_folder.glob("*周報*.docx"))
    word_files = [f for f in word_files if "~$" not in f.name]  # 排除暫存檔
    if not word_files:
        raise FileNotFoundError(f"找不到 Word 週報: {sub_folder}/*.docx")
    word_file = sorted(word_files, key=lambda f: f.stat().st_mtime, reverse=True)[0]

    # 找 NLM 提示詞
    nlm_type = "B" if path_type == "總經" else "C"
    prompt_files = list(sub_folder.glob(f"*NotebookLM*{nlm_type}*提示詞*.md"))
    if not prompt_files:
        prompt_files = list(sub_folder.glob(f"*NLM*{nlm_type}*.md"))
    focus_prompt = ""
    if prompt_files:
        focus_prompt = prompt_files[0].read_text(encoding="utf-8").strip()
        print(f"  [提示詞] {prompt_files[0].name} ({len(focus_prompt)} chars)")
    else:
        print(f"  [提示詞] 未找到，NLM 將自由發揮")

    # 找 PDF 來源（排除 NLM 產出、暫存、太小的檔案）
    pdf_files = []
    for pdf in sorted(sub_folder.glob("*.pdf")):
        name_lower = pdf.name.lower()
        # 排除 NLM 產出的投影片 PDF
        if "nlm_slide" in name_lower or "nlm_" in name_lower:
            continue
        # 排除太小的檔案 (< 100KB)
        if pdf.stat().st_size < 100_000:
            continue
        pdf_files.append(pdf)
    # 限制最多 19 個 PDF（避免稀釋 NLM 品質）
    if len(pdf_files) > 19:
        pdf_files = sorted(pdf_files, key=lambda f: f.stat().st_size, reverse=True)[:19]
    print(f"  [Word]  {word_file.name}")
    print(f"  [PDF]   {len(pdf_files)} 個檔案")

    # 建 NLM notebook
    title = f"{'總經' if path_type == '總經' else '投資策略'}週報 {date_str}"
    print(f"\n  [NLM] 建立 notebook: {title}")
    output = nlm_cmd("notebook", "create", title)
    notebook_id = parse_notebook_id(output)
    notebook_url = f"https://notebooklm.google.com/notebook/{notebook_id}"
    print(f"  [NLM] ID: {notebook_id}")

    # 上傳 Word
    print(f"  [NLM] 上傳 {word_file.name} ...")
    nlm_cmd("source", "add", notebook_id, "--file", str(word_file), "--wait")

    # 上傳 PDF
    for pdf in pdf_files:
        print(f"  [NLM] 上傳 {pdf.name} ...")
        nlm_cmd("source", "add", notebook_id, "--file", str(pdf), "--wait", "--wait-timeout", "180")

    # 等一下讓 NLM 處理
    time.sleep(5)

    # 生成投影片
    print(f"\n  [NLM] 生成投影片 ...")
    slide_args = [
        "slides", "create", notebook_id,
        "--language", "zh-TW",
        "--confirm",
    ]
    if focus_prompt:
        slide_args += ["--focus", focus_prompt]
    try:
        nlm_cmd(*slide_args, timeout=60)
    except RuntimeError:
        # slides create 可能在 stderr 輸出但仍然成功啟動
        pass

    # 等待完成
    print(f"  [NLM] 等待投影片生成 ...")
    artifact = wait_for_slides(notebook_id)
    print(f"  [NLM] 投影片生成完成！")

    # 下載 PDF
    output_path = sub_folder / f"NLM_{nlm_type}_投影片_{date_str}.pdf"
    print(f"  [NLM] 下載至 {output_path.name} ...")
    nlm_cmd(
        "download", "slide-deck", notebook_id,
        "--output", str(output_path),
        "--no-progress",
    )
    print(f"  [NLM] ✓ 下載完成: {output_path}")

    return {
        "notebook_id": notebook_id,
        "notebook_url": notebook_url,
        "output_path": str(output_path),
        "path_type": path_type,
    }


def run(date_str: str, targets: list[str]):
    """主流程"""
    print(f"{'='*60}")
    print(f"  NLM 週報自動化 — {date_str}")
    print(f"{'='*60}")

    work_folder = find_work_folder(date_str)
    print(f"工作資料夾: {work_folder}\n")

    results = []
    for target in targets:
        print(f"\n{'─'*40}")
        print(f"  路徑: {target}")
        print(f"{'─'*40}")
        try:
            result = process_path(date_str, target, work_folder)
            results.append(result)
        except Exception as e:
            print(f"  ✗ 失敗: {e}")
            results.append({"path_type": target, "error": str(e)})

    # 摘要
    print(f"\n{'='*60}")
    print(f"  結果摘要")
    print(f"{'='*60}")
    for r in results:
        if "error" in r:
            print(f"  ✗ {r['path_type']}: {r['error']}")
        else:
            print(f"  ✓ {r['path_type']}")
            print(f"    Notebook: {r['notebook_url']}")
            print(f"    PDF:      {r['output_path']}")

    return results


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python3 nlm_weekly_report.py <YYYYMMDD> [總經|投資策略|all]")
        print("範例: python3 nlm_weekly_report.py 20260406 總經")
        print("      python3 nlm_weekly_report.py 20260406 all")
        sys.exit(1)

    date_str = sys.argv[1]
    target = sys.argv[2] if len(sys.argv) > 2 else "all"

    if target == "all":
        targets = ["總經", "投資策略"]
    else:
        targets = [target]

    run(date_str, targets)
