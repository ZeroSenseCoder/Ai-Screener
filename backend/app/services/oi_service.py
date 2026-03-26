"""
Open Interest (OI) data for Indian F&O stocks.

Sources:
  1. NSE derivative quote API  → futures + options OI per expiry (free, no auth)
  2. yfinance .options         → options chain OI (works for some NSE stocks)
"""

import logging
from datetime import datetime

import httpx
import yfinance as yf

logger = logging.getLogger(__name__)

NSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer":    "https://www.nseindia.com/",
    "Accept":     "application/json",
}


async def fetch_nse_oi(symbol: str) -> dict:
    """
    Fetch futures + options OI from NSE for a given symbol.
    Returns structured OI data: total futures OI, PCR, and per-expiry breakdown.
    """
    base_sym = symbol.upper().replace(".NS", "").replace(".BO", "")
    url = f"https://www.nseindia.com/api/quote-derivative?symbol={base_sym}"
    try:
        async with httpx.AsyncClient(timeout=15, headers=NSE_HEADERS) as client:
            # warm up session cookie
            await client.get("https://www.nseindia.com")
            resp = await client.get(url)
            if resp.status_code != 200:
                return {"available": False, "reason": f"NSE returned {resp.status_code}"}
            data = resp.json()
    except Exception as e:
        logger.warning(f"NSE OI fetch failed for {base_sym}: {e}")
        return {"available": False, "reason": str(e)}

    stocks = data.get("stocks", [])
    if not stocks:
        return {"available": False, "reason": "Not an F&O stock"}

    futures = []
    options_by_expiry: dict[str, dict] = {}
    total_call_oi = 0
    total_put_oi  = 0

    for item in stocks:
        md = item.get("marketDeptOrderBook", {})
        meta = item.get("metadata", {})
        inst = meta.get("instrumentType", "")
        expiry = meta.get("expiryDate", "")
        oi    = md.get("tradeInfo", {}).get("openInterest", 0) or 0
        chg_oi = md.get("tradeInfo", {}).get("changeinOpenInterest", 0) or 0
        ltp   = meta.get("lastPrice", 0)

        if "FUT" in inst:
            futures.append({
                "expiry":     expiry,
                "oi":         oi,
                "change_oi":  chg_oi,
                "ltp":        ltp,
            })

        elif "CE" in inst:  # Call option
            total_call_oi += oi
            strike = meta.get("strikePrice", 0)
            if expiry not in options_by_expiry:
                options_by_expiry[expiry] = {"calls": [], "puts": []}
            options_by_expiry[expiry]["calls"].append({
                "strike": strike, "oi": oi, "change_oi": chg_oi, "ltp": ltp,
            })

        elif "PE" in inst:  # Put option
            total_put_oi += oi
            strike = meta.get("strikePrice", 0)
            if expiry not in options_by_expiry:
                options_by_expiry[expiry] = {"calls": [], "puts": []}
            options_by_expiry[expiry]["puts"].append({
                "strike": strike, "oi": oi, "change_oi": chg_oi, "ltp": ltp,
            })

    pcr = round(total_put_oi / total_call_oi, 3) if total_call_oi else None

    # Near-month expiry breakdown (first expiry)
    near_expiry = sorted(options_by_expiry.keys())[0] if options_by_expiry else None
    near_data = options_by_expiry.get(near_expiry, {}) if near_expiry else {}

    # Top 5 strikes by OI for near expiry
    top_calls = sorted(near_data.get("calls", []), key=lambda x: x["oi"], reverse=True)[:5]
    top_puts  = sorted(near_data.get("puts",  []), key=lambda x: x["oi"], reverse=True)[:5]

    return {
        "available":     True,
        "symbol":        base_sym,
        "futures":       futures[:3],          # near + mid + far month
        "total_call_oi": total_call_oi,
        "total_put_oi":  total_put_oi,
        "pcr":           pcr,                   # Put-Call Ratio
        "near_expiry":   near_expiry,
        "top_call_strikes": top_calls,          # highest OI call strikes
        "top_put_strikes":  top_puts,           # highest OI put strikes
    }


def fetch_yf_options_oi(yf_symbol: str) -> dict:
    """
    Fallback: get options OI from yfinance (works for some NSE F&O stocks).
    """
    try:
        t = yf.Ticker(yf_symbol)
        expirations = t.options
        if not expirations:
            return {"available": False}

        nearest = expirations[0]
        chain = t.option_chain(nearest)
        calls = chain.calls[["strike", "openInterest", "lastPrice", "impliedVolatility"]].head(10).to_dict(orient="records")
        puts  = chain.puts[["strike", "openInterest", "lastPrice", "impliedVolatility"]].head(10).to_dict(orient="records")

        call_oi = int(chain.calls["openInterest"].sum())
        put_oi  = int(chain.puts["openInterest"].sum())
        pcr = round(put_oi / call_oi, 3) if call_oi else None

        return {
            "available":        True,
            "source":           "yfinance",
            "near_expiry":      nearest,
            "total_call_oi":    call_oi,
            "total_put_oi":     put_oi,
            "pcr":              pcr,
            "top_call_strikes": sorted(calls, key=lambda x: x.get("openInterest", 0), reverse=True)[:5],
            "top_put_strikes":  sorted(puts,  key=lambda x: x.get("openInterest", 0), reverse=True)[:5],
        }
    except Exception as e:
        logger.warning(f"yfinance OI fetch failed for {yf_symbol}: {e}")
        return {"available": False}


async def get_oi(symbol: str, yf_symbol: str) -> dict:
    """
    Try NSE first, fall back to yfinance options.
    """
    nse = await fetch_nse_oi(symbol)
    if nse.get("available"):
        nse["source"] = "NSE"
        return nse
    yf_oi = fetch_yf_options_oi(yf_symbol)
    return yf_oi
