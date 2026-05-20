#!/usr/bin/env python3
"""
Windows 移植安裝腳本 — 在 Windows Claude Code 環境中執行

用法：
  python setup_windows.py

會自動：
  1. 偵測 OneDrive 路徑
  2. 安裝 Python 依賴
  3. 安裝 npm 依賴
  4. 檢查系統工具（poppler, tesseract, nlm）
  5. 複製 Claude Code skills
  6. 匯入 Claude Code memory
  7. 驗證環境
"""

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

# ── 顏色輸出 ──────────────────────────────────────────────────────
def ok(msg): print(f"  ✓ {msg}")
def warn(msg): print(f"  ⚠ {msg}")
def fail(msg): print(f"  ✗ {msg}")
def info(msg): print(f"  → {msg}")


def check_platform():
    """確認在 Windows 上執行"""
    if platform.system() != "Windows":
        warn(f"此腳本為 Windows 設計，目前在 {platform.system()} 上執行")
        return False
    ok(f"Windows {platform.version()}")
    return True


def find_onedrive():
    """偵測 OneDrive 路徑"""
    home = Path.home()
    candidates = [
        home / "OneDrive" / "桌面" / "每周報告",
        home / "OneDrive - 個人" / "桌面" / "每周報告",
        home / "OneDrive" / "Desktop" / "每周報告",
    ]
    # 也檢查環境變數
    od_env = os.environ.get("OneDrive", "")
    if od_env:
        candidates.insert(0, Path(od_env) / "桌面" / "每周報告")

    for c in candidates:
        if c.exists():
            ok(f"OneDrive 週報: {c}")
            return c
        elif c.parent.exists():
            ok(f"OneDrive 路徑存在（週報資料夾待建）: {c.parent}")
            return c

    fail("找不到 OneDrive 路徑")
    info("請確認 OneDrive 已登入且同步")
    return None


def install_python_deps():
    """安裝 Python 依賴"""
    project_dir = Path(__file__).parent
    pyproject = project_dir / "pyproject.toml"

    if not pyproject.exists():
        fail("找不到 pyproject.toml")
        return False

    # 嘗試 uv，fallback 到 pip
    try:
        subprocess.run(["uv", "--version"], capture_output=True, check=True)
        info("使用 uv 安裝...")
        subprocess.run(["uv", "pip", "install", "-e", str(project_dir)], check=True)
        ok("Python 依賴安裝完成 (uv)")
        return True
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass

    try:
        info("使用 pip 安裝...")
        subprocess.run([sys.executable, "-m", "pip", "install", "-e", str(project_dir)], check=True)
        ok("Python 依賴安裝完成 (pip)")
        return True
    except subprocess.CalledProcessError as e:
        fail(f"安裝失敗: {e}")
        return False


def install_npm_deps():
    """安裝 npm 依賴"""
    project_dir = Path(__file__).parent
    pkg_json = project_dir / "package.json"

    if not pkg_json.exists():
        warn("沒有 package.json，跳過 npm")
        return True

    try:
        subprocess.run(["npm", "install"], cwd=str(project_dir), check=True)
        ok("npm 依賴安裝完成")
        return True
    except (FileNotFoundError, subprocess.CalledProcessError):
        fail("npm install 失敗，請確認 Node.js 已安裝")
        return False


def check_system_tools():
    """檢查系統工具"""
    tools = {
        "pdftoppm": "poppler (PDF 工具)",
        "tesseract": "tesseract OCR",
        "nlm": "NotebookLM CLI",
    }
    missing = []
    for cmd, desc in tools.items():
        if shutil.which(cmd):
            ok(f"{desc}: {shutil.which(cmd)}")
        else:
            fail(f"{desc} 未安裝")
            missing.append(cmd)

    if missing:
        print("\n  安裝指令：")
        if "pdftoppm" in missing:
            info("scoop install poppler  (或手動下載 poppler-windows)")
        if "tesseract" in missing:
            info("winget install UB-Mannheim.TesseractOCR  (安裝時勾選 Chinese Traditional)")
        if "nlm" in missing:
            info("uv tool install notebooklm-mcp && nlm login")
        return False
    return True


def check_tesseract_langs():
    """檢查 tesseract 中文語言包"""
    try:
        result = subprocess.run(
            ["tesseract", "--list-langs"],
            capture_output=True, text=True, timeout=10,
        )
        if "chi_tra" in result.stdout:
            ok("tesseract 繁體中文語言包已安裝")
            return True
        else:
            warn("tesseract 缺少 chi_tra 語言包")
            info("請重新安裝 tesseract 並勾選 Chinese Traditional")
            return False
    except (FileNotFoundError, subprocess.CalledProcessError):
        return False


def setup_skills():
    """複製 Claude Code skills"""
    project_dir = Path(__file__).parent
    skills_src = project_dir / ".claude" / "skills"

    if not skills_src.exists():
        fail(f"找不到 skills 目錄: {skills_src}")
        return False

    # Skills 已經在專案中，Claude Code 會自動讀取
    skill_count = sum(1 for f in skills_src.rglob("SKILL.md"))
    ok(f"Skills 已就緒: {skill_count} 個 skill")
    for skill_md in skills_src.rglob("SKILL.md"):
        info(f"  {skill_md.parent.name}/SKILL.md")
    return True


def setup_memory():
    """匯入 Claude Code memory"""
    project_dir = Path(__file__).parent
    memory_src = project_dir / "memory_export"

    if not memory_src.exists():
        warn("沒有 memory_export/ 資料夾，跳過 memory 匯入")
        info("請從 Mac 複製 memory 檔案到 memory_export/")
        info("或在 Claude Code 中手動建立 memory")
        return False

    # Claude Code memory 路徑（Windows）
    # 格式：%APPDATA%/Claude/projects/<project-hash>/memory/
    appdata = os.environ.get("APPDATA", "")
    if not appdata:
        fail("找不到 APPDATA 環境變數")
        return False

    claude_projects = Path(appdata) / "Claude" / "projects"
    if not claude_projects.exists():
        warn(f"Claude Code projects 目錄不存在: {claude_projects}")
        info("請先用 Claude Code 開啟此專案一次")
        return False

    # 找到對應的 project 目錄
    project_dirs = list(claude_projects.iterdir())
    if not project_dirs:
        warn("沒有找到 Claude Code project 目錄")
        return False

    # 用最新的 project 目錄
    target_dir = sorted(project_dirs, key=lambda d: d.stat().st_mtime, reverse=True)[0]
    memory_dir = target_dir / "memory"
    memory_dir.mkdir(exist_ok=True)

    copied = 0
    for md_file in memory_src.glob("*.md"):
        dst = memory_dir / md_file.name
        shutil.copy2(md_file, dst)
        copied += 1

    ok(f"Memory 匯入完成: {copied} 個檔案 → {memory_dir}")
    return True


def check_env_vars():
    """檢查環境變數"""
    required = {
        "ANTHROPIC_API_KEY": "Anthropic API 金鑰",
    }
    optional = {
        "GMAIL_APP_PASSWORD": "Gmail App Password（寄信用）",
    }

    all_ok = True
    for var, desc in required.items():
        if os.environ.get(var):
            ok(f"{desc}: 已設定")
        else:
            fail(f"{desc}: 未設定 ({var})")
            all_ok = False

    for var, desc in optional.items():
        if os.environ.get(var):
            ok(f"{desc}: 已設定")
        else:
            warn(f"{desc}: 未設定（選用）")

    return all_ok


def check_font():
    """檢查 Microsoft JhengHei 字型"""
    fonts_dir = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
    jhenghei = fonts_dir / "msjh.ttc"
    if jhenghei.exists():
        ok("Microsoft JhengHei 字型已安裝")
        return True
    else:
        warn("Microsoft JhengHei 字型未找到（應該預裝在 Windows 上）")
        return False


def main():
    print("=" * 55)
    print("  國泰週報系統 — Windows 移植安裝")
    print("=" * 55)

    sections = [
        ("平台檢查", check_platform),
        ("OneDrive 偵測", find_onedrive),
        ("環境變數", check_env_vars),
        ("Python 依賴", install_python_deps),
        ("npm 依賴", install_npm_deps),
        ("系統工具", check_system_tools),
        ("OCR 語言包", check_tesseract_langs),
        ("字型檢查", check_font),
        ("Claude Skills", setup_skills),
        ("Claude Memory", setup_memory),
    ]

    results = {}
    for name, func in sections:
        print(f"\n{'─'*40}")
        print(f"  {name}")
        print(f"{'─'*40}")
        try:
            results[name] = func()
        except Exception as e:
            fail(f"錯誤: {e}")
            results[name] = False

    # 摘要
    print(f"\n{'='*55}")
    print("  移植結果")
    print(f"{'='*55}")
    for name, passed in results.items():
        status = "✓" if passed else "✗"
        print(f"  {status} {name}")

    failed = [k for k, v in results.items() if not v]
    if failed:
        print(f"\n  有 {len(failed)} 項需要處理，請依上方指示修正")
    else:
        print(f"\n  全部通過！可以開始使用週報系統")


if __name__ == "__main__":
    main()
