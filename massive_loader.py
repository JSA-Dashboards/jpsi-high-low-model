"""
massive_loader.py
-----------------
Download CBOT session-aggregate flat files from Massive's S3-compatible storage,
filter to corn (ZC*) and soybean (ZS*) futures, cache locally as Parquet, and
return a dict that matches the format already used by load_price_history():

    {
      "corn": { "ZCZ26": DataFrame(Open/High/Low/Close/OI/Volume, index=date), ... },
      "soy":  { "ZSX26": DataFrame(...), ... },
    }

Caching strategy
----------------
- One Parquet per commodity per year:  data/massive_cache/cbot_ZC_2024.parquet
- Past years (< current year) are cached permanently.
- Current year is always re-downloaded on refresh so today's data is included.
"""

import os
import io
import gzip
import logging
from pathlib import Path
from datetime import datetime, timedelta, date

import boto3
import pandas as pd
from botocore.config import Config
from botocore.exceptions import ClientError

log = logging.getLogger(__name__)

# ── Massive S3 config ────────────────────────────────────────────────────────
_ENDPOINT  = "https://files.massive.com"
_BUCKET    = "flatfiles"
_PREFIX    = "us_futures_cbot/session_aggs_v1"

# ── Local cache directory ─────────────────────────────────────────────────────
_CACHE_DIR = Path(__file__).parent / "data" / "massive_cache"

# ── Earliest date Massive has (Developer plan = 5 years back) ─────────────────
_MASSIVE_START = date(2021, 9, 1)   # ~5 yr back from mid-2026; adjust if needed


# ─────────────────────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────────────────────

def _get_cred(name: str) -> str | None:
    """Read a credential from env var, then Streamlit secrets as fallback."""
    val = os.environ.get(name)
    if val:
        return val
    try:
        import streamlit as st
        return st.secrets.get(name)
    except Exception:
        return None


def _s3():
    """Create an authenticated boto3 S3 client for Massive."""
    key_id = _get_cred("MASSIVE_ACCESS_KEY_ID")
    secret = _get_cred("MASSIVE_SECRET_ACCESS_KEY")
    if not key_id or not secret:
        raise RuntimeError(
            "MASSIVE_ACCESS_KEY_ID / MASSIVE_SECRET_ACCESS_KEY not set. "
            "Add them to your .env file (local) or Streamlit Cloud Secrets (cloud)."
        )
    return boto3.client(
        "s3",
        endpoint_url=_ENDPOINT,
        aws_access_key_id=key_id,
        aws_secret_access_key=secret,
        config=Config(signature_version="s3v4"),
    )


def _s3_key(d: date) -> str:
    return f"{_PREFIX}/{d.year:04d}/{d.month:02d}/{d.strftime('%Y-%m-%d')}.csv.gz"


def _cache_path(commodity: str, year: int) -> Path:
    return _CACHE_DIR / f"cbot_{commodity}_{year}.parquet"


def _download_day(client, d: date) -> pd.DataFrame | None:
    """Fetch one daily CSV.gz from S3; return filtered rows (ZC*/ZS*) or None."""
    key = _s3_key(d)
    try:
        obj = client.get_object(Bucket=_BUCKET, Key=key)
        raw = obj["Body"].read()
        with gzip.open(io.BytesIO(raw), "rb") as fh:
            df = pd.read_csv(fh)
        # keep only the tickers we care about
        df = df[df["ticker"].str.startswith(("ZC", "ZS"), na=False)].copy()
        return df if not df.empty else None
    except ClientError as e:
        if e.response["Error"]["Code"] in ("NoSuchKey", "404"):
            return None          # holiday / weekend — normal
        log.warning("S3 error on %s: %s", key, e)
        return None
    except Exception as e:
        log.warning("Error reading %s: %s", key, e)
        return None


def _trading_days(start: date, end: date):
    """Generate Mon–Fri dates in [start, end]."""
    d = start
    while d <= end:
        if d.weekday() < 5:   # 0=Mon … 4=Fri
            yield d
        d += timedelta(days=1)


def _download_year(client, year: int, prefix_filters=("ZC", "ZS")) -> pd.DataFrame:
    """Download all trading days for one year; return concatenated DataFrame."""
    start = max(date(year, 1, 1), _MASSIVE_START)
    end   = min(date(year, 12, 31), date.today() - timedelta(days=1))
    if start > end:
        return pd.DataFrame()

    frames = []
    for d in _trading_days(start, end):
        df = _download_day(client, d)
        if df is not None:
            frames.append(df)

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def _to_ohlcv(df: pd.DataFrame) -> pd.DataFrame:
    """Rename Massive columns → Open/High/Low/Close/OI/Volume indexed by date."""
    df = df.copy()
    df["date"] = pd.to_datetime(df["session_end_date"])
    df = df.sort_values("date")
    df = df.rename(columns={
        "open":  "Open",
        "high":  "High",
        "low":   "Low",
        "close": "Close",
        "volume": "Volume",
    })
    df["OI"] = float("nan")
    df = df.set_index("date")[["Open", "High", "Low", "Close", "OI", "Volume"]]
    return df


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def refresh_massive_cache() -> tuple[bool, str]:
    """
    (Re)download Massive data for all years from _MASSIVE_START to today.
    Past years are only re-downloaded if their cache file is missing.
    Current year is always refreshed.
    Returns (success, message).
    """
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)
    try:
        client = _s3()
    except RuntimeError as e:
        return False, str(e)

    today     = date.today()
    cur_year  = today.year
    messages  = []

    # Years to process
    start_year = _MASSIVE_START.year
    for year in range(start_year, cur_year + 1):
        zc_path = _cache_path("ZC", year)
        zs_path = _cache_path("ZS", year)

        # Skip past years that are already cached
        if year < cur_year and zc_path.exists() and zs_path.exists():
            messages.append(f"✅ {year}: using cached data")
            continue

        messages.append(f"⬇️  {year}: downloading from Massive…")
        year_df = _download_year(client, year)

        if year_df.empty:
            messages.append(f"   ⚠️  {year}: no data returned")
            continue

        for prefix, label, path in [
            ("ZC", "ZC", zc_path),
            ("ZS", "ZS", zs_path),
        ]:
            sub = year_df[year_df["ticker"].str.startswith(prefix)].copy()
            if not sub.empty:
                sub.to_parquet(path, index=False)
                n_tickers = sub["ticker"].nunique()
                messages.append(f"   ✅ {label} {year}: {n_tickers} contracts cached")
            else:
                messages.append(f"   ⚠️  {label} {year}: no contracts found")

    return True, "\n".join(messages)


def load_massive_history() -> tuple[dict | None, str | None]:
    """
    Load cached Massive data from disk and return (data_dict, error_str).

    data_dict = {
        "corn": { ticker: DataFrame(Open/High/Low/Close/OI/Volume, index=date) },
        "soy":  { ticker: DataFrame(...) },
    }
    Returns (None, error_str) if no cache exists yet.
    """
    _CACHE_DIR.mkdir(parents=True, exist_ok=True)

    zc_files = sorted(_CACHE_DIR.glob("cbot_ZC_*.parquet"))
    zs_files = sorted(_CACHE_DIR.glob("cbot_ZS_*.parquet"))

    if not zc_files and not zs_files:
        return None, (
            "No Massive cache found. Click '🔄 Refresh from Massive' "
            "to download price history."
        )

    def _load_group(files) -> dict:
        frames = []
        for f in files:
            try:
                frames.append(pd.read_parquet(f))
            except Exception as e:
                log.warning("Skipping %s: %s", f, e)
        if not frames:
            return {}
        big = pd.concat(frames, ignore_index=True)
        result = {}
        for ticker, grp in big.groupby("ticker"):
            result[ticker] = _to_ohlcv(grp)
        return result

    corn = _load_group(zc_files)
    soy  = _load_group(zs_files)

    if not corn and not soy:
        return None, "Massive cache is empty. Run refresh."

    return {"corn": corn, "soy": soy}, None
