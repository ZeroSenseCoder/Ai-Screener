"""
Global macro and economic data (all free sources):
  - Global indices via yfinance
  - FII/DII data via NSE India API
  - USD/INR forex via yfinance
  - India VIX via yfinance
"""

import logging
from datetime import datetime

import httpx
import yfinance as yf

logger = logging.getLogger(__name__)

GLOBAL_INDICES = {
    "NIFTY 50": "^NSEI",
    "SENSEX": "^BSESN",
    "S&P 500": "^GSPC",
    "NASDAQ": "^IXIC",
    "FTSE 100": "^FTSE",
    "NIKKEI 225": "^N225",
    "HANG SENG": "^HSI",
    "DAX": "^GDAXI",
    "SHANGHAI": "000001.SS",
}

FOREX_PAIRS = {
    "USD/INR": "USDINR=X",
    "EUR/INR": "EURINR=X",
    "GBP/INR": "GBPINR=X",
    "JPY/INR": "JPYINR=X",
}

COMMODITIES = {
    "Crude Oil (WTI)": "CL=F",
    "Crude Oil (Brent)": "BZ=F",
    "Gold": "GC=F",
    "Silver": "SI=F",
    "Natural Gas": "NG=F",
}


def _get_quote(yf_symbol: str) -> dict:
    try:
        t = yf.Ticker(yf_symbol)
        fi = t.fast_info
        price = getattr(fi, "last_price", None) or getattr(fi, "regularMarketPrice", None)
        prev = getattr(fi, "previous_close", None) or getattr(fi, "regularMarketPreviousClose", None)
        change_pct = ((price - prev) / prev * 100) if price and prev else None
        return {
            "price": round(price, 2) if price else None,
            "change_pct": round(change_pct, 2) if change_pct else None,
        }
    except Exception:
        return {"price": None, "change_pct": None}


def fetch_global_indices() -> list[dict]:
    results = []
    for name, symbol in GLOBAL_INDICES.items():
        quote = _get_quote(symbol)
        results.append({"name": name, "symbol": symbol, **quote})
    return results


def fetch_forex() -> list[dict]:
    results = []
    for name, symbol in FOREX_PAIRS.items():
        quote = _get_quote(symbol)
        results.append({"name": name, "symbol": symbol, **quote})
    return results


def fetch_commodities() -> list[dict]:
    results = []
    for name, symbol in COMMODITIES.items():
        quote = _get_quote(symbol)
        results.append({"name": name, "symbol": symbol, **quote})
    return results


def fetch_india_vix() -> dict:
    return {"name": "India VIX", **_get_quote("^INDIAVIX")}


async def fetch_fii_dii() -> list[dict]:
    """
    Fetch FII/DII net flows from NSE India (free, no auth required).
    Returns last available day's data.
    """
    try:
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Referer": "https://www.nseindia.com/",
        }
        async with httpx.AsyncClient(timeout=15, headers=headers) as client:
            # Get session cookie first
            await client.get("https://www.nseindia.com")
            resp = await client.get(
                "https://www.nseindia.com/api/fiidiiTradeReact"
            )
            data = resp.json()
            results = []
            for row in data[:5]:  # last 5 trading days
                results.append({
                    "date": row.get("date", ""),
                    "fii_net": row.get("fiiNet", row.get("fii_net", 0)),
                    "dii_net": row.get("diiNet", row.get("dii_net", 0)),
                    "fii_buy": row.get("fiiBuy", 0),
                    "fii_sell": row.get("fiiSell", 0),
                    "dii_buy": row.get("diiBuy", 0),
                    "dii_sell": row.get("diiSell", 0),
                })
            return results
    except Exception as e:
        logger.warning(f"FII/DII fetch failed: {e}")
        return []


async def get_macro_overview() -> dict:
    """Return all macro data in one payload."""
    return {
        "global_indices": fetch_global_indices(),
        "forex": fetch_forex(),
        "commodities": fetch_commodities(),
        "india_vix": fetch_india_vix(),
        "fii_dii": await fetch_fii_dii(),
        "timestamp": datetime.utcnow().isoformat(),
    }
