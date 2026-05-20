# 國泰證券 — 共享報告系統

## 跨平台設定
- Mac: `~/Library/CloudStorage/OneDrive-個人/桌面/Coding/`
- Windows: `C:\Users\tsait\OneDrive\桌面\Coding\`
- 週報資料夾: `OneDrive/桌面/每周報告/YYYYMMDD/`
- 專題報告: `OneDrive/桌面/專題報告/{主題}/`

## 驗證規則
統一由 `verification-rules` skill 定義（三輪驗證 R1/R2/R3、來源分級 A/B/C、Gate `_draft` 規則、台灣金融用語、NLM 規範、爬蟲分級 L1/L2/L3）。
- 權威位置：`OneDrive/桌面/Coding/.claude/skills/verification-rules/SKILL.md`
- Windows junction：`%USERPROFILE%\Desktop\Py\.claude\skills\verification-rules`
- Mac symlink：`~/.claude/plugins/general-reports/skills/verification-rules`
- 被 weekly-report、daily-market-report、topic-report、financial-report、ppt-template-rules 引用

## Skills 清單
| Skill | 觸發詞 | 來源 |
|-------|--------|------|
| weekly-report | 做週報、做總經、做投資策略 | OneDrive（共用） |
| daily-market-report | 做日報、做早報 | OneDrive（共用） |
| topic-report | 專題報告、做專題 | OneDrive（共用） |
| financial-report | 做財報、財報分析 | OneDrive（共用） |
| verification-rules | (被其他 skill 引用) | OneDrive（共用） |
| ppt-template-rules | (被其他 skill 引用) | OneDrive（共用） |

## Scripts
位於 `OneDrive/桌面/Coding/scripts/`（跨 Mac/Windows 同步）

| 腳本 | 用途 |
|------|------|
| `nlm_weekly_report.py` | NLM 自動化（建 notebook → 上傳 → 生成投影片 → 下載） |
| `nlm_embed_slides.py` | 裁剪 NLM PDF + 嵌入 Weekly PPTX（動態偵測裁剪） |
| `assemble_pptx.py` | 裁剪 NLM PDF + 組裝 PPTX（固定比例 + JSON config） |
| `gen_pptx.py` | JSON → 可編輯 PPTX（bullet、雙欄、討論題、summary 卡片） |
| `nlm_market_digest.py` | market_news → NLM 整合摘要 |

## v4 PPTX 配色
| 角色 | 色碼 |
|---|---|
| 主色 Teal | `#0F766E` |
| 內頁背景 | `#FAFBFC` |
| 文字主色 | `#1E293B` |
| 漲 | `#15803D` |
| 跌 | `#B91C1C` |

## 台灣術語（必用）
超配→加碼、低配→減碼、標配→中立、對沖→避險、國債→公債、波動率→波動度、回報→報酬、宏觀→總體經濟

## Earnings Trigger 自動財報系統
位於 `earnings_trigger/` 目錄，自動監測美股財報並觸發 Claude 分析。

| 模組 | 用途 |
|------|------|
| `calendar_watcher.py` | yfinance 抓取 earnings calendar |
| `sec_trigger.py` | SEC EDGAR 8-K/10-Q/10-K 觸發 |
| `gmail_trigger.py` | Gmail 關鍵字觸發（Phase 2） |
| `job_runner.py` | Job queue → Claude API 分析 |
| `run_claude_earnings.py` | Anthropic SDK wrapper |

- Watchlist: `data/us_earnings_watchlist.csv`（NVDA, MSFT, AMZN, GOOGL, META）
- 輸出: `~/Desktop/財報/{ticker}/YYYY-MM-DD_{ticker}_earnings.md`
- venv: `earnings_trigger/.venv/`
- 跨平台: `config.py` 自動偵測 macOS / Windows

## NLM 整合
- 提示詞簡單為主（格式+語氣，不預填數據）
- 標題限單行 ≤15 中文字
- 裁剪：top 11%（標題）、bottom 3.5%（浮水印）
- Layout[5]「只有標題」：ph0=主標、ph1=副標、移除 ph2
