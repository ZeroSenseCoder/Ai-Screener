"""
/api/v1/stocks — per-stock detail: live quote, OHLCV (all timeframes), indicators, OI, news.
"""

import math
from fastapi import APIRouter, HTTPException, Query
from app.core.state import app_state
from app.services.market_data import (
    fetch_ohlcv, fetch_live_quote, fetch_fundamentals, fetch_yf_news, TIMEFRAME_MAP
)
from app.services.indicators import compute_all_indicators
from app.services.news_service import get_stock_news
from app.services.oi_service import get_oi
from app.config import get_settings

router = APIRouter(prefix="/stocks", tags=["Stocks"])
settings = get_settings()

VALID_TIMEFRAMES = list(TIMEFRAME_MAP.keys())


def _clean(obj):
    """Recursively replace NaN/Inf with None for JSON safety."""
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    return obj


def _resolve(symbol: str) -> tuple[str, str, str]:
    """Return (yf_symbol, base_symbol, company_name)."""
    universe_df = app_state.universe_df
    if universe_df is None:
        raise HTTPException(503, "Universe not loaded yet")

    upper = symbol.upper()
    row = universe_df[universe_df["yf_symbol"] == upper]
    if row.empty:
        row = universe_df[universe_df["symbol"] == upper]
    if row.empty:
        row = universe_df[universe_df["symbol"].str.contains(upper, na=False)]
    if row.empty:
        raise HTTPException(404, f"Symbol '{symbol}' not found")

    r = row.iloc[0]
    return r["yf_symbol"], r["symbol"], str(r.get("company_name", symbol))


# ── Search ─────────────────────────────────────────────────────────────────────

@router.get("/search")
async def search_stocks(q: str = Query(..., min_length=1), limit: int = 10):
    universe_df = app_state.universe_df
    if universe_df is None:
        return []
    q_lower = q.lower()
    mask = (
        universe_df["symbol"].str.lower().str.startswith(q_lower)
        | universe_df["company_name"].str.lower().str.contains(q_lower, na=False)
    )
    results = universe_df[mask].head(limit)
    drop = ["bse_code", "_priority", "industry_raw"]
    results = results.drop(columns=[c for c in drop if c in results.columns])
    return _clean(results.where(results.notna(), other=None).to_dict(orient="records"))


# ── Live Quote ─────────────────────────────────────────────────────────────────

@router.get("/{symbol}/quote")
async def get_quote(symbol: str):
    yf_sym, _, company_name = _resolve(symbol)
    quote = fetch_live_quote(yf_sym)
    return _clean({"symbol": symbol, "yf_symbol": yf_sym, "company_name": company_name, **quote})


# ── OHLCV / Chart ──────────────────────────────────────────────────────────────

@router.get("/{symbol}/ohlcv")
async def get_ohlcv(
    symbol: str,
    timeframe: str = Query("1D", description=f"One of: {', '.join(VALID_TIMEFRAMES)}"),
):
    """
    Candlestick data for the given timeframe.
    Supported: 1m 5m 15m 1h 4h 1D 1W 1M 1Y 5Y MAX
    """
    if timeframe not in TIMEFRAME_MAP:
        raise HTTPException(400, f"Invalid timeframe. Use one of: {VALID_TIMEFRAMES}")
    yf_sym, _, _ = _resolve(symbol)
    df = fetch_ohlcv(yf_sym, timeframe)
    if df.empty:
        return {"symbol": symbol, "timeframe": timeframe, "candles": []}
    candles = _clean(df.where(df.notna(), other=None).to_dict(orient="records"))
    return {"symbol": symbol, "yf_symbol": yf_sym, "timeframe": timeframe, "candles": candles}


# ── Indicators ─────────────────────────────────────────────────────────────────

@router.get("/{symbol}/indicators")
async def get_indicators(symbol: str):
    yf_sym, _, _ = _resolve(symbol)
    merged = app_state.merged_df
    if merged is not None and not merged.empty:
        row = merged[merged["yf_symbol"] == yf_sym]
        if not row.empty:
            ind_cols = [
                "last_price", "sma_20", "sma_50", "sma_200",
                "ema_20", "ema_50", "ema_200", "rsi_14",
                "macd", "macd_signal", "macd_hist", "beta",
                "max_drawdown_52w", "daily_return", "return_5d",
                "return_1m", "return_3m", "return_1y", "avg_volume_20d",
            ]
            available = [c for c in ind_cols if c in row.columns]
            data = row[available].iloc[0].where(row[available].iloc[0].notna(), None).to_dict()
            return _clean({"symbol": symbol, "indicators": data})
    # Fallback: compute on the fly from 2Y daily data
    df = fetch_ohlcv(yf_sym, "1Y")
    indicators = compute_all_indicators(df) if not df.empty else {}
    return _clean({"symbol": symbol, "indicators": indicators})


# ── Fundamentals ───────────────────────────────────────────────────────────────

@router.get("/{symbol}/fundamentals")
async def get_fundamentals(symbol: str):
    yf_sym, _, _ = _resolve(symbol)
    data = fetch_fundamentals(yf_sym)
    return _clean({"symbol": symbol, "yf_symbol": yf_sym, **data})


# ── Open Interest ───────────────────────────────────────────────────────────────

@router.get("/{symbol}/oi")
async def get_open_interest(symbol: str):
    """F&O open interest from NSE (or yfinance options as fallback)."""
    yf_sym, base_sym, _ = _resolve(symbol)
    oi_data = await get_oi(base_sym, yf_sym)
    return _clean({"symbol": symbol, **oi_data})


# ── News ────────────────────────────────────────────────────────────────────────

@router.get("/{symbol}/news")
async def get_news(
    symbol: str,
    days: int = Query(7, ge=1, le=30),
):
    """
    News from Yahoo Finance (free, no key) + RSS feeds + optional NewsAPI.
    Returns articles sorted by importance (high → medium → low).
    """
    yf_sym, base_sym, company_name = _resolve(symbol)

    # Fetch Yahoo Finance news (primary — always free)
    yf_news = fetch_yf_news(yf_sym)

    # Merge with RSS + optional NewsAPI
    articles = await get_stock_news(
        symbol=base_sym,
        company_name=company_name,
        yf_news=yf_news,
        api_key=settings.news_api_key,
        days=days,
    )
    return _clean({
        "symbol":       symbol,
        "company_name": company_name,
        "days":         days,
        "count":        len(articles),
        "articles":     articles,
    })


# ── Full Summary ────────────────────────────────────────────────────────────────

@router.get("/{symbol}/summary")
async def get_summary(symbol: str):
    """Live quote + fundamentals + indicators in one call."""
    yf_sym, base_sym, company_name = _resolve(symbol)

    # Check cached merged data first
    merged = app_state.merged_df
    cached: dict = {}
    if merged is not None and not merged.empty:
        row = merged[merged["yf_symbol"] == yf_sym]
        if not row.empty:
            r = row.iloc[0]
            cached = {
                "exchange":       r.get("exchange"),
                "sector":         r.get("sector"),
                "industry":       r.get("industry_raw", r.get("industry", "")),
                "pe_ratio":       r.get("pe_ratio"),
                "beta":           r.get("beta"),
                "market_cap":     r.get("market_cap"),
                "rsi_14":         r.get("rsi_14"),
                "sma_20":         r.get("sma_20"),
                "sma_50":         r.get("sma_50"),
                "sma_200":        r.get("sma_200"),
                "ema_20":         r.get("ema_20"),
                "ema_50":         r.get("ema_50"),
                "ema_200":        r.get("ema_200"),
                "macd":           r.get("macd"),
                "macd_signal":    r.get("macd_signal"),
                "macd_hist":      r.get("macd_hist"),
                "max_drawdown_52w": r.get("max_drawdown_52w"),
                "daily_return":   r.get("daily_return"),
                "return_5d":      r.get("return_5d"),
                "return_1m":      r.get("return_1m"),
                "return_3m":      r.get("return_3m"),
                "return_6m":      r.get("return_6m"),
                "return_1y":      r.get("return_1y"),
                "avg_volume_20d": r.get("avg_volume_20d"),
            }

    # If technicals are missing from the enrichment cache, compute them on-the-fly now
    if cached.get("rsi_14") is None and cached.get("sma_20") is None:
        try:
            ohlcv_df = fetch_ohlcv(yf_sym, "1Y")
            if not ohlcv_df.empty and len(ohlcv_df) >= 20:
                computed = compute_all_indicators(ohlcv_df)
                for k, v in computed.items():
                    if cached.get(k) is None:
                        cached[k] = v
        except Exception:
            pass

    # Always fetch live quote (real-time)
    quote = fetch_live_quote(yf_sym)

    # Always fetch fundamentals for complete data (pe, eps, book value, 52w high/low, etc.)
    fundamentals = fetch_fundamentals(yf_sym)
    # Merge: fundamentals fill in any missing fields, but don't overwrite non-null cached values
    for k, v in fundamentals.items():
        if cached.get(k) is None:
            cached[k] = v

    result = {
        "symbol":       base_sym,
        "yf_symbol":    yf_sym,
        "company_name": company_name,
        **cached,
        **quote,   # live quote overrides cached price
    }
    return _clean(result)
