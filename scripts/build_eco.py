"""Build final economic calendar JSON for 2026-04-20 to 2026-04-24.
Selects top events per day, formats numbers, groups by time, adds JPM/GS notes.
"""
import json
import os
from datetime import datetime, date
from collections import defaultdict

SRC_RAW = r"C:\Users\tsait\OneDrive\桌面\Coding\.tmp\eco_rows.json"
OUT = r"C:\Users\tsait\OneDrive\桌面\04_週報日報\每周報告\20260420\經濟數據\economic_calendar.json"

with open(SRC_RAW, "r", encoding="utf-8") as f:
    rows = json.load(f)

# Major economies (Bloomberg country codes)
MAJOR = {"US", "CH", "JN", "EC", "GE", "UK", "TA", "SK", "CA"}
EXTRA = {"FR", "IT", "SP", "AU"}

CC_DISPLAY = {
    "US": "美國", "CH": "中國", "JN": "日本", "EC": "歐元區", "GE": "德國",
    "UK": "英國", "TA": "台灣", "SK": "韓國", "CA": "加拿大", "FR": "法國",
    "IT": "義大利", "SP": "西班牙", "AU": "澳洲",
}

WEEKDAY = {0: "週一", 1: "週二", 2: "週三", 3: "週四", 4: "週五", 5: "週六", 6: "週日"}

WEEK_START = date(2026, 4, 20)
WEEK_END = date(2026, 4, 24)

# JPM / GS notes by ticker or event name - curated from PDFs
JPM_NOTES = {
    # US week data from JPM p23
    ("US", "零售銷售(月比)"): "JPM +1.6% m/m（汽油暴漲拉抬）",
    ("US", "零售銷售(除汽車)(月比)"): "JPM +1.6% m/m（控制組 +0.5%）",
    ("US", "零售銷售(除汽車及汽油)"): "JPM +0.6% m/m",
    ("US", "零售銷售(控制組)"): "JPM 控制組 +0.5% m/m",
    ("US", "成屋待售銷售 (月比)"): "JPM +1.5% m/m（三月抵押利率升 40bp 為逆風）",
    ("US", "費城聯儲非製造業指數"): "JPM 預期回升、4 月 Empire 升至 11",
    ("US", "首次申請失業救濟的人數"): "JPM 215k（前值 207k），仍在兩年低檔",
    ("US", "續請失業救濟的人數"): "JPM 1820k（前值 1818k），續請回落暗示失業率已見頂",
    ("US", "標準普爾全球製造業PMI"): "JPM 53.0（前值 52.3），地區 Fed 調查強勢",
    ("US", "標準普爾全球服務業PMI"): "JPM 51.0（前值 49.8），關鍵能否回到 50 上",
    ("US", "標準普爾全球綜合PMI"): "JPM 綜合企穩 52 上方",
    ("US", "密西根大學信心指數"): "JPM 4 月終值 49.0（初值 47.6）",
    ("US", "密西根大學1年通膨預期"): "4.8%（短期通膨預期跳升 100bp）",
    ("US", "密西根大學5-10年通膨預期"): "3.4%（+20bp）",
    ("US", "MBA 貸款申請指數"): "JPM 3 月 +6.4%，但房貸利率 +40bp 壓抑後續",
    # Canada
    ("CA", "CPI(年比)"): "JPM 2.6% oya（前值 1.8%）、Core 2.1%",
    ("CA", "CPI不經季調(月比)"): "JPM 1.2% m/m、核心 0.3% m/m",
    # Japan
    ("JN", "貿易收支"): "JPM 實質出口 +14.1% m/m、進口 +9.1% m/m",
    ("JN", "出口(年比)"): "JPM 11.0%、搶運效應",
    ("JN", "進口(年比)"): "JPM 7.0%",
    ("JN", "標準普爾日本採購經理人指數-製造業"): "JPM 49.5 prelim（低於榮枯線）",
    ("JN", "標準普爾日本採購經理人指數-服務業"): "JPM 52.7 prelim",
    ("JN", "全國消費者物價指數 年比"): "JPM 核心 1.5% oya（Ex fresh food）、核心核心 2.3%",
    # Taiwan
    ("TA", "出口訂單(年比)"): "JPM 43.8% oya（AI 拉貨動能）",
    ("TA", "失業率"): "JPM 3.32% sa",
    ("TA", "工業生產(年比)"): "JPM 25.8% oya（AI 出口帶動）",
    # Korea
    ("SK", "PPI(年比)"): "JPM 4.5% oya（能源驅動）",
    ("SK", "1Q GDP"): "JPM 1Q 3.1% oya（年化 5.0%）",
    # China
    ("CH", "1年貸款基放利率"): "JPM 持穩 3.0%；2Q26 可能 -10bp",
    ("CH", "5年貸款基放利率"): "JPM 持穩 3.5%；Q1 GDP 飆 6.7% yoy 減緩降息壓力",
    # Euro area / Germany
    ("EC", "ZEW調查預期"): "JPM 受地緣政治改善帶動反彈",
    ("GE", "ZEW調查預期"): "JPM -7（前值 -0.5），能源衝擊壓抑",
    ("GE", "ZEW調查現況"): "JPM -70.7（前值 -62.9）",
    ("GE", "IFO企業景氣指數"): "JPM 85.5（前值 86.4）、景氣信心降至 4 月新低",
    ("GE", "IFO預期"): "JPM 85.0（前值 86.0）",
    ("GE", "IFO商業現況指數"): "JPM 86.2（前值 86.7）",
    ("EC", "消費者信心指數"): "JPM -17.2（前值 -16.3）",
    # UK
    ("UK", "CPI(年比)"): "JPM 3.3% oya（前值 3.0%）",
    ("UK", "核心CPI(年比)"): "JPM 3.3% oya（服務通膨 4.5%）",
    ("UK", "ILO失業率(3個月)"): "JPM 5.2%",
    ("UK", "零售銷售不含汽油(月比)"): "JPM -0.5% m/m",
}

def parse_date(dt_str):
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str).date()
    except Exception:
        pass
    if "-" in dt_str and "/" in dt_str:
        start = dt_str.split("-")[0]
        try:
            return datetime.strptime(start, "%m/%d/%Y").date()
        except Exception:
            return None
    return None

def in_range(d):
    return d is not None and WEEK_START <= d <= WEEK_END

# Filter
filtered = []
for r in rows:
    dt = r.get("Date Time")
    if not isinstance(dt, str) or "T" not in dt:
        continue  # skip range-date entries
    d = parse_date(dt)
    cc = (r.get("Country Code") or "").strip()
    if cc not in (MAJOR | EXTRA):
        continue
    if not in_range(d):
        continue
    r["_date"] = d.isoformat()
    filtered.append(r)

# Group by date
by_date = defaultdict(list)
for r in filtered:
    by_date[r["_date"]].append(r)

# Formatting helpers
def is_percent_event(name):
    if not name:
        return False
    markers = ["率", "月比", "年比", "季比", "CPI", "PPI", "PCE", "GDP",
               "利率", "FOMC", "Rate", "%", "通膨", "P.C.E", "Inflation"]
    return any(k in str(name) for k in markers)

def fmt_value(v, event_name):
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if not s or "--" in s:
            return None
        return s
    if isinstance(v, float):
        # Heuristic: if value < 1 and likely percent, multiply
        if abs(v) < 1 and is_percent_event(event_name):
            return f"{v*100:.1f}%"
        if abs(v) < 1 and "萬" not in str(event_name):
            # Ambiguous: treat as percent if looks fractional
            return f"{v*100:.1f}%"
        if abs(v) >= 1 and abs(v) < 10 and is_percent_event(event_name):
            return f"{v:.1f}%"
        if abs(v) >= 1000:
            return f"{int(v):,}"
        return f"{v:.1f}"
    if isinstance(v, int):
        if is_percent_event(event_name) and abs(v) < 20:
            return f"{v}%"
        return f"{v:,}"
    return str(v)

def fmt_time(dt_str):
    if not dt_str or "T" not in str(dt_str):
        return ""
    hhmm = dt_str.split("T")[1][:5]
    return f"{hhmm} TW"

# Heavy keywords for focus-day flag
HEAVY_KEYWORDS = [
    "零售銷售", "新屋銷售", "IFO", "密西根",
    "基放利率", "CPI", "採購經理人指數", "PMI",
    "ZEW", "耐久財", "失業率", "工業生產",
    "貿易收支", "出口訂單", "消費者物價",
]

# Build days
days = []
for d in sorted(by_date):
    dt_obj = date.fromisoformat(d)
    weekday = WEEKDAY[dt_obj.weekday()]
    items = by_date[d]
    items_sorted = sorted(items, key=lambda r: -(r.get("Relevance") or 0))

    selected = []
    seen = set()

    # Pass 1: include all US events with relevance >= 30
    for r in items_sorted:
        if r.get("Country Code") != "US":
            continue
        if (r.get("Relevance") or 0) < 30:
            continue
        key = (r["Country Code"], r["Event"])
        if key in seen:
            continue
        seen.add(key)
        selected.append(r)

    # Pass 2: TA and CH with relevance >= 40
    for r in items_sorted:
        if r.get("Country Code") not in ("TA", "CH", "SK", "JN"):
            continue
        if (r.get("Relevance") or 0) < 40:
            continue
        key = (r["Country Code"], r["Event"])
        if key in seen:
            continue
        seen.add(key)
        selected.append(r)

    # Pass 3: top by relevance others - up to 12 total
    for r in items_sorted:
        if len(selected) >= 12:
            break
        if (r.get("Relevance") or 0) < 60:
            continue
        key = (r["Country Code"], r["Event"])
        if key in seen:
            continue
        seen.add(key)
        selected.append(r)

    # Sort selected by time then relevance desc
    def _sk(r):
        dt = r.get("Date Time")
        try:
            return (datetime.fromisoformat(dt), -(r.get("Relevance") or 0))
        except Exception:
            return (datetime.min, 0)

    selected.sort(key=_sk)

    # Focus day?
    focus = False
    for r in selected:
        ev = r.get("Event") or ""
        cc = r.get("Country Code")
        if cc == "US" and any(k in ev for k in HEAVY_KEYWORDS):
            focus = True
            break
        if cc in ("GE", "UK", "CH", "TA") and any(k in ev for k in ["CPI", "IFO", "基放利率", "ZEW", "採購經理人指數", "出口訂單"]):
            focus = True
            break

    events_out = []
    time_groups = defaultdict(list)
    for r in selected:
        ev_name = r.get("Event") or ""
        cc = r.get("Country Code")
        tm = fmt_time(r.get("Date Time"))
        period = r.get("Period")
        if isinstance(period, str) and "T" in period:
            period = period.split("T")[0]
        entry = {
            "time": tm,
            "country": CC_DISPLAY.get(cc, cc),
            "event": ev_name,
            "period": period,
            "survey": fmt_value(r.get("Survey"), ev_name),
            "prior": fmt_value(r.get("Prior"), ev_name),
            "relevance": round(r.get("Relevance") or 0, 1),
        }
        note = JPM_NOTES.get((cc, ev_name))
        if note:
            entry["note"] = note
        events_out.append(entry)
        time_groups[tm].append(entry)

    tg_list = [{"time": tm, "events": time_groups[tm]} for tm in sorted(time_groups)]

    days.append({
        "date": d,
        "weekday": weekday,
        "focus": focus,
        "events": events_out,
        "time_groups": tg_list,
    })

# Highlights - curated narrative
highlights = [
    {
        "title": "美國 3 月零售銷售",
        "date": "2026-04-21",
        "analysis": (
            "Bloomberg 預估 headline MoM +1.4%（前值 +0.6%），除汽車 +1.4%；"
            "JPM 預測 +1.6% m/m、控制組 +0.5%。油價飆 25% 推升加油站銷售 +13% m/m、"
            "車輛銷售 +3.7% 亦偏強，惟實質消費支出因能源擠壓、低收入族群 Medicaid/SNAP 縮減"
            "而停滯。若控制組不如 +0.5% 預期則消費動能疑慮升溫。"
        ),
    },
    {
        "title": "4 月 Flash PMI",
        "date": "2026-04-23",
        "analysis": (
            "美國 Markit 製造業 PMI 預估 52.5（前值 52.3），JPM 看 53.0 小升；"
            "服務業預估 50.1（前值 49.8），JPM 看 51.0 企穩榮枯線上。Empire Fed 4 月"
            "大幅跳增至 11.0、Philly Fed Composite 56.2，顯示製造業仍具韌性。服務業連 "
            "9 個月下滑，3 月首破 50 是警訊；若 4 月未能回升將引發 Fed 關切。"
        ),
    },
    {
        "title": "德國 4 月 IFO 與 ZEW",
        "date": "2026-04-21 / 2026-04-24",
        "analysis": (
            "週二 ZEW 預期 -7.0（前值 -0.5）、現況 -70.7（前值 -62.9），反映中東戰爭能源"
            "衝擊壓抑信心；週五 IFO 企業景氣 85.5、預期 85.0、現況 86.2（皆較前值下滑 "
            "0.5-1 點）。JPM 將 ECB 升息由 4 月推遲至 6 月+9 月兩次，歐洲核心通膨接近 2% 目標。"
        ),
    },
    {
        "title": "英國 3 月 CPI",
        "date": "2026-04-22",
        "analysis": (
            "CPI 年比預估 3.3%（前值 3.0%）、月比 0.6%、核心 3.2%、服務通膨仍高於 4.5%，"
            "為 G4 通膨最具黏性者。勞動市場週二揭曉、ILO 失業率 5.2%（3 個月）。BoE 5 月 7 日"
            "開會，市場定價全年僅再降 25bp；若 CPI 續高，降息路徑將更為謹慎。"
        ),
    },
    {
        "title": "中國 LPR 與台灣出口訂單",
        "date": "2026-04-20 / 2026-04-21",
        "analysis": (
            "週一 PBoC 1Y / 5Y LPR 預期均持穩於 3.0% / 3.5%；上週公布 1Q GDP 飆 6.7% yoy "
            "（JPM tracking 7.0% ar、出口 +70%），減緩短線降息壓力。週二台灣 3 月出口訂單 "
            "預估 +43.8% yoy，AI 拉貨為主力；週四失業率 3.32%、工業生產 +25.8% yoy，"
            "台灣出口景氣續火熱。"
        ),
    },
    {
        "title": "加拿大 3 月 CPI 與 BoC Outlook",
        "date": "2026-04-20",
        "analysis": (
            "加拿大 CPI 年比預估 2.6%（前值 1.8%）、月比 1.2%、核心修剪平均 2.1%。"
            "同日公布 BoC 1Q Business Outlook Survey。JPM 預期 BoC 4/29 維持 2.25% "
            "overnight rate 不變，終點利率 2.25% 已到位。能源推升通膨但實質消費走弱，"
            "BoC 需觀察服務通膨再決定進一步行動。"
        ),
    },
]

out = {
    "date_range": "2026-04-20 至 2026-04-24",
    "days": days,
    "highlights": highlights,
}

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

total = sum(len(d["events"]) for d in days)
focus_days = sum(1 for d in days if d["focus"])
print(f"Days: {len(days)}, total events: {total}, focus days: {focus_days}")
for d in days:
    print(f"  {d['date']} {d['weekday']} focus={d['focus']} events={len(d['events'])}")
    for e in d["events"]:
        n = f" | note: {e['note']}" if e.get("note") else ""
        print(f"    {e['time']} {e['country']} | {e['event']} ({e['period']}) survey={e['survey']} prior={e['prior']}{n}")
print(f"\nWrote {OUT}")
