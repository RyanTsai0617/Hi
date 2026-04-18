from datetime import datetime, timezone
import hashlib


def utc_now_iso():
    return datetime.now(timezone.utc).isoformat()


def make_dedupe_key(*parts):
    raw = "|".join("" if p is None else str(p) for p in parts)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def normalize_cik(cik: str | int) -> str:
    return str(cik).zfill(10)
