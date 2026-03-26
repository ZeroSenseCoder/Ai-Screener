"""
/api/v1/pro-screener — Bloomberg-style weighted screener.
"""

import math
from typing import Optional
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel, Field
from app.core.state import app_state
from app.services.pro_screener_engine import (
    FILTER_CATALOG,
    FilterCondition,
    apply_pro_screen,
    get_catalog_by_category,
)

router = APIRouter(prefix="/pro-screener", tags=["Pro Screener"])


def _clean(obj):
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    return obj


# ── Request / Response models ──────────────────────────────────────────────────

class ConditionIn(BaseModel):
    filter_id: str
    min_val: Optional[float] = None
    max_val: Optional[float] = None
    weight: float = Field(default=0.0, ge=0.0, le=1.0)  # decimal; all weights sum to 1
    required: bool = True


class ProScreenRequest(BaseModel):
    sector: Optional[str] = None
    conditions: list[ConditionIn] = []
    score_mode: bool = False
    sort_by: str = "score"
    sort_asc: bool = False
    page: int = Field(default=1, ge=1)
    page_size: int = Field(default=50, ge=1, le=200)


# ── Endpoints ──────────────────────────────────────────────────────────────────

@router.get("/catalog")
def get_filter_catalog():
    """Return all available filters grouped by category."""
    return get_catalog_by_category()


@router.post("/screen")
def pro_screen(req: ProScreenRequest):
    """
    Apply weighted filters to the stock universe.
    Returns scored + ranked results.
    """
    merged = app_state.merged_df
    if merged is None or merged.empty:
        raise HTTPException(503, "Universe not loaded yet")

    # Validate filter IDs
    bad = [c.filter_id for c in req.conditions if c.filter_id not in FILTER_CATALOG]
    if bad:
        raise HTTPException(400, f"Unknown filter IDs: {bad}")

    conditions = [
        FilterCondition(
            filter_id=c.filter_id,
            min_val=c.min_val,
            max_val=c.max_val,
            weight=c.weight,
            required=c.required,
        )
        for c in req.conditions
    ]

    result_df, total = apply_pro_screen(
        merged_df=merged,
        sector=req.sector,
        conditions=conditions,
        score_mode=req.score_mode,
        sort_by=req.sort_by,
        sort_asc=req.sort_asc,
        page=req.page,
        page_size=req.page_size,
    )

    # Select columns for output
    output_cols = [
        "symbol", "company_name", "exchange", "sector", "score",
        "last_price", "market_cap", "pe_ratio", "beta",
        "rsi_14", "macd_hist", "max_drawdown_52w",
        "daily_return", "return_5d", "return_1m", "return_3m", "return_6m", "return_1y",
        "avg_volume_20d", "sma_20", "sma_50", "sma_200",
        "ema_20", "ema_50", "ema_200", "macd", "macd_signal",
        # Valuation
        "forward_pe", "price_to_book", "price_to_sales", "ev_ebitda", "ev_sales",
        "dividend_yield", "fcf_yield",
        # Quality
        "roe", "roa", "roic", "gross_margins", "operating_margins", "profit_margins",
        "revenue_growth", "earnings_growth", "ebitda_growth", "forecast_growth",
        # Leverage
        "debt_to_equity", "net_debt_ebitda", "total_debt", "interest_coverage", "volatility_52w",
        # Cash flow
        "free_cashflow", "operating_cashflow",
        # Analyst
        "target_upside", "analyst_buy_pct", "analyst_mean_rating",
        "earnings_revisions", "earnings_surprise",
        # ESG
        "esg_score", "esg_controversies",
        # Other
        "relative_strength",
    ]
    available_cols = [c for c in output_cols if c in result_df.columns]
    page_df = result_df[available_cols]

    records = page_df.where(page_df.notna(), other=None).to_dict(orient="records")

    return _clean({
        "total":      total,
        "page":       req.page,
        "page_size":  req.page_size,
        "sector":     req.sector,
        "score_mode": req.score_mode,
        "conditions": len(conditions),
        "results":    records,
    })


@router.get("/sectors")
def list_sectors():
    """Return all sectors with stock counts."""
    merged = app_state.merged_df
    if merged is None:
        return []
    counts = (
        merged.groupby("sector")["symbol"]
        .count()
        .reset_index()
        .rename(columns={"symbol": "count"})
        .sort_values("count", ascending=False)
    )
    return _clean(counts.to_dict(orient="records"))
