"""Filter Bloomberg calendar rows to US + major economies, 2026-04-20 to 2026-04-24."""
import json
from datetime import datetime, date

SRC = r"C:\Users\tsait\OneDrive\桌面\Coding\.tmp\eco_rows.json"
OUT = r"C:\Users\tsait\OneDrive\桌面\Coding\.tmp\eco_filtered.json"

with open(SRC, "r", encoding="utf-8") as f:
    rows = json.load(f)

MAJOR = {"US", "CH", "JN", "EC", "GE", "UK", "TA", "KS", "CA"}
# Bloomberg country code variants: China sometimes "CH", Taiwan sometimes "TA" or "TW",
# Korea south sometimes "KS" (Bloomberg code) or "KR"
EXTRA = {"FR", "IT", "SP"}  # other euro-area large ones, keep but lower priority

WEEK_START = date(2026, 4, 20)
WEEK_END = date(2026, 4, 24)

def parse_date(dt_str):
    if not dt_str:
        return None
    # ISO datetime
    try:
        return datetime.fromisoformat(dt_str).date()
    except Exception:
        pass
    # Range format '4/20/2026-4/24/2026' - take start
    if "-" in dt_str and "/" in dt_str:
        start = dt_str.split("-")[0]
        try:
            return datetime.strptime(start, "%m/%d/%Y").date()
        except Exception:
            try:
                return datetime.strptime(start, "%Y-%m-%d").date()
            except Exception:
                return None
    return None

def in_range(d):
    if d is None:
        return False
    return WEEK_START <= d <= WEEK_END

def is_range_overlap(dt_str):
    # For range entries like '4/20/2026-4/24/2026' check if any part overlaps week
    if not isinstance(dt_str, str) or "-" not in dt_str:
        return False
    parts = dt_str.split("-")
    if len(parts) != 2:
        return False
    try:
        start = datetime.strptime(parts[0], "%m/%d/%Y").date()
        end = datetime.strptime(parts[1], "%m/%d/%Y").date()
    except Exception:
        return False
    # overlap with WEEK_START..WEEK_END
    return not (end < WEEK_START or start > WEEK_END)

filtered = []
range_entries = []  # for week-range rows
for r in rows:
    dt = r.get("Date Time")
    d = parse_date(dt) if isinstance(dt, str) else None
    country = (r.get("Country Code") or "").strip()
    if country not in (MAJOR | EXTRA):
        continue
    if isinstance(dt, str) and "-" in dt and "/" in dt:
        # week-range entry
        if is_range_overlap(dt):
            range_entries.append(r)
        continue
    if in_range(d):
        r["_parsed_date"] = d.isoformat()
        filtered.append(r)

# Sort by datetime then relevance desc
def sort_key(r):
    try:
        dt = datetime.fromisoformat(r["Date Time"])
    except Exception:
        dt = datetime.min
    rel = r.get("Relevance") or 0
    return (dt, -rel)

filtered.sort(key=sort_key)
print(f"Filtered events in week: {len(filtered)}")
print(f"Week-range entries overlapping week: {len(range_entries)}")

# Group by date
by_date = {}
for r in filtered:
    d = r["_parsed_date"]
    by_date.setdefault(d, []).append(r)

print("\n=== Per-day counts ===")
for d in sorted(by_date):
    print(f"{d}: {len(by_date[d])} events")
    # print top 15 by relevance
    day_rows = sorted(by_date[d], key=lambda r: -(r.get('Relevance') or 0))
    for r in day_rows[:20]:
        dt = r.get('Date Time')
        tm = dt.split('T')[1][:5] if 'T' in str(dt) else ''
        print(f"  {tm} {r.get('Country Code')} rel={r.get('Relevance'):.0f} | {r.get('Event')} ({r.get('Period')}) survey={r.get('Survey')} prior={r.get('Prior')}")

with open(OUT, "w", encoding="utf-8") as f:
    json.dump({"filtered": filtered, "range_entries": range_entries, "by_date": by_date}, f, ensure_ascii=False, indent=2, default=str)
print(f"\nWrote {OUT}")
