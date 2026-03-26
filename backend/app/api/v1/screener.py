"""
/api/v1/screener — filter stocks by technical and fundamental criteria.
"""

import math
from typing import List, Optional
from fastapi import APIRouter
from pydantic import BaseModel

from app.core.state import app_state
from app.services.screener_engine import ScreenerFilters, apply_filters, get_sector_summary


def _clean(obj):
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    return obj

router = APIRouter(prefix="/screener", tags=["Screener"])


class ScreenerRequest(BaseModel):
    exchanges: List[str] = ["NSE", "BSE"]
    sectors: List[str] = []
    industries: List[str] = []
    min_price: Optional[float] = None
    max_price: Optional[float] = None
    min_market_cap: Optional[float] = None
    max_market_cap: Optional[float] = None
    min_pe: Optional[float] = None
    max_pe: Optional[float] = None
    min_beta: Optional[float] = None
    max_beta: Optional[float] = None
    min_rsi: Optional[float] = None
    max_rsi: Optional[float] = None
    price_above_sma20: bool = False
    price_above_sma50: bool = False
    price_above_sma200: bool = False
    sma50_above_sma200: bool = False
    price_below_sma200: bool = False
    macd_bullish: bool = False
    macd_bearish: bool = False
    min_daily_return: Optional[float] = None
    max_daily_return: Optional[float] = None
    min_return_1m: Optional[float] = None
    max_return_1m: Optional[float] = None
    min_return_3m: Optional[float] = None
    max_return_3m: Optional[float] = None
    min_return_1y: Optional[float] = None
    max_return_1y: Optional[float] = None
    min_avg_volume: Optional[int] = None
    max_drawdown_threshold: Optional[float] = None
    sort_by: str = "market_cap"
    sort_asc: bool = False
    page: int = 1
    page_size: int = 50


@router.post("/screen")
async def screen_stocks(req: ScreenerRequest):
    """Apply filters and return paginated stock results."""
    merged_df = app_state.merged_df
    if merged_df is None or merged_df.empty:
        return {"results": [], "total": 0, "page": 1, "page_size": req.page_size}

    filters = ScreenerFilters(**req.model_dump())
    result_df, total = apply_filters(merged_df, filters)

    columns = [
        "symbol", "company_name", "exchange", "sector",
        "last_price", "daily_return", "return_1m", "return_3m", "return_6m", "return_1y",
        "market_cap", "pe_ratio", "beta", "rsi_14",
        "sma_20", "sma_50", "sma_200",
        "macd", "macd_signal", "avg_volume_20d", "max_drawdown_52w",
    ]
    available = [c for c in columns if c in result_df.columns]
    records = _clean(result_df[available].to_dict(orient="records"))

    return {
        "results": records,
        "total": total,
        "page": req.page,
        "page_size": req.page_size,
        "pages": (total + req.page_size - 1) // req.page_size,
    }


@router.get("/sector-summary")
async def sector_summary():
    """Sector-level aggregates for the filter panel sidebar."""
    merged_df = app_state.merged_df
    if merged_df is None or merged_df.empty:
        return []
    return _clean(get_sector_summary(merged_df))


@router.get("/meta")
async def screener_meta():
    """Return available filter options (sectors, industries, exchanges)."""
    universe_df = app_state.universe_df
    if universe_df is None or universe_df.empty:
        return {"sectors": [], "industries": [], "exchanges": ["NSE", "BSE"]}

    sectors = sorted(universe_df["sector"].dropna().unique().tolist())
    industries = sorted(universe_df["industry"].dropna().unique().tolist())
    industries = [i for i in industries if i]  # remove empty strings

    return {
        "sectors": sectors,
        "industries": industries,
        "exchanges": ["NSE", "BSE"],
    }
