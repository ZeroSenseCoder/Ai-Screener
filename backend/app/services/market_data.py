"""
Fetch OHLCV data from yfinance — supports all timeframes including intraday.
"""

import logging
import time
from datetime import datetime, timedelta, timezone

import pandas as pd
import yfinance as yf

logger = logging.getLogger(__name__)

# ── Timeframe → yfinance (interval, period) mapping ───────────────────────────
# For 4H: we fetch 1H and resample to 4-bar groups on the fly
TIMEFRAME_MAP: dict[str, dict] = {
    "1m":  {"interval": "1m",   "period": "1d"},
    "5m":  {"interval": "5m",   "period": "5d"},
    "15m": {"interval": "15m",  "period": "5d"},
    "1h":  {"interval": "60m",  "period": "1mo"},
    "4h":  {"interval": "60m",  "period": "3mo",  "resample": "4h"},
    "1D":  {"interval": "1d",   "period": "1y"},
    "1W":  {"interval": "1wk",  "period": "5y"},
    "1M":  {"interval": "1mo",  "period": "max"},
    "1Y":  {"interval": "1d",   "period": "1y"},
    "5Y":  {"interval": "1wk",  "period": "5y"},
    "MAX": {"interval": "1mo",  "period": "max"},
}

# Timeframes that carry datetime (not just date) timestamps
INTRADAY_TFS = {"1m", "5m", "15m", "1h", "4h"}


def _resample_4h(df: pd.DataFrame) -> pd.DataFrame:
    """Resample a 1H OHLCV DataFrame to 4-hour bars."""
    df = df.copy()
    df["ts"] = pd.to_datetime(df["date"])
    df = df.set_index("ts").sort_index()
    agg = df[["open", "high", "low", "close", "volume"]].resample("4h", closed="left", label="left").agg({
        "open":   "first",
        "high":   "max",
        "low":    "min",
        "close":  "last",
        "volume": "sum",
    }).dropna()
    agg = agg.reset_index()
    agg["date"] = agg["ts"].dt.strftime("%Y-%m-%dT%H:%M:%S")
    return agg[["date", "open", "high", "low", "close", "volume"]]


def fetch_ohlcv(yf_symbol: str, timeframe: str = "1D") -> pd.DataFrame:
    """
    Fetch OHLCV for a single symbol using the given timeframe key.
    Returns DataFrame with columns: date, open, high, low, close, volume
      - Intraday timeframes: date is ISO datetime string
      - Daily+  timeframes: date is YYYY-MM-DD string
    """
    cfg = TIMEFRAME_MAP.get(timeframe, TIMEFRAME_MAP["1D"])
    interval = cfg["interval"]
    period   = cfg["period"]
    resample = cfg.get("resample")

    try:
        ticker = yf.Ticker(yf_symbol)
        df = ticker.history(period=period, interval=interval, auto_adjust=True)
        if df.empty:
            return pd.DataFrame()

        df = df.reset_index()
        # Normalise column names
        df.columns = [c.lower() for c in df.columns]
        # The index column is called "datetime" for intraday, "date" for daily
        if "datetime" in df.columns:
            df = df.rename(columns={"datetime": "date"})

        # Format timestamp
        dt_col = pd.to_datetime(df["date"])
        if timeframe in INTRADAY_TFS and resample != "4h":
            df["date"] = dt_col.dt.strftime("%Y-%m-%dT%H:%M:%S")
        else:
            df["date"] = dt_col.dt.strftime("%Y-%m-%d")

        result = df[["date", "open", "high", "low", "close", "volume"]].copy()

        # Resample to 4h if needed
        if resample == "4h":
            # Re-parse with proper datetime for resampling
            result["date"] = pd.to_datetime(df["date"])
            result = result.set_index("date").resample("4h", closed="left", label="left").agg({
                "open": "first", "high": "max", "low": "min",
                "close": "last", "volume": "sum"
            }).dropna().reset_index()
            result["date"] = result["date"].dt.strftime("%Y-%m-%dT%H:%M:%S")

        # Drop rows with NaN prices
        result = result.dropna(subset=["close"])
        return result

    except Exception as e:
        logger.warning(f"fetch_ohlcv({yf_symbol}, {timeframe}): {e}")
        return pd.DataFrame()


def fetch_live_quote(yf_symbol: str) -> dict:
    """Real-time quote — price, change, 52w high/low, volume."""
    try:
        t = yf.Ticker(yf_symbol)
        fi = t.fast_info
        price    = getattr(fi, "last_price", None)
        prev     = getattr(fi, "previous_close", None)
        chg      = (price - prev) if price and prev else None
        chg_pct  = (chg / prev * 100) if chg and prev else None
        open_p = getattr(fi, "open", None)
        return {
            "last_price":   round(price, 2) if price else None,
            "open":         round(open_p, 2) if open_p else None,
            "prev_close":   round(prev, 2) if prev else None,
            "change":       round(chg, 2) if chg else None,
            "change_pct":   round(chg_pct, 2) if chg_pct else None,
            "day_high":     getattr(fi, "day_high", None),
            "day_low":      getattr(fi, "day_low", None),
            "year_high":    getattr(fi, "year_high", None),
            "year_low":     getattr(fi, "year_low", None),
            "52w_high":     getattr(fi, "year_high", None),
            "52w_low":      getattr(fi, "year_low", None),
            "volume":       getattr(fi, "last_volume", None),
            "avg_volume":   getattr(fi, "three_month_average_volume", None),
            "market_cap":   getattr(fi, "market_cap", None),
        }
    except Exception as e:
        logger.warning(f"fetch_live_quote({yf_symbol}): {e}")
        return {}


def fetch_fundamentals(yf_symbol: str) -> dict:
    """PE ratio, market cap, beta, sector, dividend yield, EPS, etc."""
    try:
        info = yf.Ticker(yf_symbol).info
        return {
            "sector":            info.get("sector") or "Unknown",
            "industry":          info.get("industry") or "",
            "market_cap":        info.get("marketCap"),
            "pe_ratio":          info.get("trailingPE"),
            "forward_pe":        info.get("forwardPE"),
            "eps":               info.get("trailingEps"),
            "beta":              info.get("beta"),
            "dividend_yield":    info.get("dividendYield"),
            "book_value":        info.get("bookValue"),
            "price_to_book":     info.get("priceToBook"),
            "return_on_equity":  info.get("returnOnEquity"),
            "return_on_assets":  info.get("returnOnAssets"),
            "debt_to_equity":    info.get("debtToEquity"),
            "revenue_growth":    info.get("revenueGrowth"),
            "earnings_growth":   info.get("earningsGrowth"),
            "profit_margins":    info.get("profitMargins"),
            "52w_high":          info.get("fiftyTwoWeekHigh"),
            "52w_low":           info.get("fiftyTwoWeekLow"),
            "avg_volume":        info.get("averageVolume"),
            "shares_outstanding": info.get("sharesOutstanding"),
            "float_shares":      info.get("floatShares"),
            "short_ratio":       info.get("shortRatio"),
        }
    except Exception:
        return {}


def fetch_yf_news(yf_symbol: str) -> list[dict]:
    """
    Fetch news from Yahoo Finance (completely free, no API key).
    yfinance.Ticker.news returns recent articles with title, link, publisher, publishedAt.
    """
    try:
        ticker = yf.Ticker(yf_symbol)
        raw = ticker.news or []
        articles = []
        for item in raw:
            content = item.get("content", {})
            # New yfinance >= 0.2.50 nests data under "content"
            title     = content.get("title") or item.get("title", "")
            url       = (content.get("canonicalUrl", {}) or {}).get("url") or item.get("link", "")
            publisher = (content.get("provider", {}) or {}).get("displayName") or item.get("publisher", "Yahoo Finance")
            pub_time  = content.get("pubDate") or item.get("providerPublishTime")
            summary   = content.get("summary") or ""
            thumb     = ((content.get("thumbnail") or {}).get("resolutions") or [{}])[0].get("url", "")

            if not title:
                continue

            # Convert unix timestamp to ISO string if needed
            if isinstance(pub_time, (int, float)):
                pub_time = datetime.fromtimestamp(pub_time, tz=timezone.utc).isoformat()

            articles.append({
                "title":        title,
                "url":          url,
                "source":       publisher,
                "published_at": pub_time or "",
                "summary":      summary,
                "thumbnail":    thumb,
            })
        return articles
    except Exception as e:
        logger.warning(f"fetch_yf_news({yf_symbol}): {e}")
        return []


def fetch_ohlcv_batch(symbols: list[str], period: str = "1y") -> dict[str, pd.DataFrame]:
    """Batch fetch daily OHLCV for a list of symbols."""
    results: dict[str, pd.DataFrame] = {}
    batch_size = 50
    for i in range(0, len(symbols), batch_size):
        batch = symbols[i: i + batch_size]
        for sym in batch:
            results[sym] = fetch_ohlcv(sym, "1D")
            time.sleep(0.2)
        logger.info(f"Batch {i // batch_size + 1}/{(len(symbols) + batch_size - 1) // batch_size}")
        time.sleep(1.0)
    return results
