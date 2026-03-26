"""
/api/v1/universe — endpoints to get all stocks and sector groupings.
"""

import math
from fastapi import APIRouter, BackgroundTasks, Query
from fastapi.responses import JSONResponse

from app.core.state import app_state


def _clean(obj):
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else obj
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    return obj

router = APIRouter(prefix="/universe", tags=["Universe"])


@router.get("/sectors")
async def get_sectors():
    """Return all sectors with their stock counts."""
    by_sector = app_state.stocks_by_sector
    return {
        sector: {
            "count": len(stocks),
            "stocks": stocks,
        }
        for sector, stocks in by_sector.items()
    }


@router.get("/sectors/list")
async def list_sectors():
    """Return just sector names and counts (lightweight)."""
    by_sector = app_state.stocks_by_sector
    return [
        {"sector": sector, "count": len(stocks)}
        for sector, stocks in sorted(by_sector.items())
    ]


@router.get("/sectors/{sector_name}")
async def get_sector_stocks(sector_name: str):
    """Return all stocks in a specific sector."""
    by_sector = app_state.stocks_by_sector
    stocks = by_sector.get(sector_name)
    if stocks is None:
        return JSONResponse(
            status_code=404,
            content={"error": f"Sector '{sector_name}' not found"},
        )
    return _clean({"sector": sector_name, "count": len(stocks), "stocks": stocks})


@router.get("/stocks")
async def list_all_stocks(
    exchange: str = Query(None, description="Filter by NSE or BSE"),
    sector: str = Query(None, description="Filter by sector name"),
    search: str = Query(None, description="Search by symbol or company name"),
    page: int = Query(1, ge=1),
    page_size: int = Query(100, ge=1, le=500),
):
    """List all stocks with optional filtering and pagination."""
    universe_df = app_state.universe_df
    if universe_df is None or universe_df.empty:
        return {"stocks": [], "total": 0}

    df = universe_df.copy()

    if exchange:
        df = df[df["exchange"].str.upper() == exchange.upper()]
    if sector:
        df = df[df["sector"].str.lower() == sector.lower()]
    if search:
        mask = (
            df["symbol"].str.contains(search, case=False, na=False)
            | df["company_name"].str.contains(search, case=False, na=False)
        )
        df = df[mask]

    total = len(df)
    start = (page - 1) * page_size
    page_df = df.iloc[start : start + page_size]

    # Drop internal columns and replace NaN with None for JSON safety
    drop_cols = ["bse_code", "_priority", "industry_raw"]
    page_df = page_df.drop(columns=[c for c in drop_cols if c in page_df.columns])
    records = page_df.where(page_df.notna(), other=None).to_dict(orient="records")

    return {
        "total": total,
        "page": page,
        "page_size": page_size,
        "stocks": records,
    }


@router.get("/stats")
async def get_universe_stats():
    """Return high-level stats about the stock universe."""
    universe_df = app_state.universe_df
    by_sector = app_state.stocks_by_sector

    if universe_df is None or universe_df.empty:
        return {"status": "not_loaded", "total": 0}

    nse_count = int((universe_df["exchange"] == "NSE").sum())
    bse_count = int((universe_df["exchange"] == "BSE").sum())

    return {
        "status": "loaded",
        "total": len(universe_df),
        "nse_count": nse_count,
        "bse_count": bse_count,
        "sector_count": len(by_sector),
        "sectors": [
            {"sector": s, "count": len(stocks)}
            for s, stocks in sorted(by_sector.items())
        ],
        "last_updated": app_state.last_updated,
    }


@router.post("/refresh")
async def refresh_universe(background_tasks: BackgroundTasks):
    """Trigger a background refresh of the stock universe."""
    from app.core.state import load_universe
    background_tasks.add_task(load_universe)
    return {"status": "refresh_started"}
