"""
Screener engine: applies user filters to the in-memory universe + indicators data.
All filtering is done in pandas for speed (no DB round-trip on every filter change).
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)


@dataclass
class ScreenerFilters:
    # Exchange
    exchanges: list[str] = field(default_factory=lambda: ["NSE", "BSE"])

    # Sector / Industry
    sectors: list[str] = field(default_factory=list)
    industries: list[str] = field(default_factory=list)

    # Price
    min_price: Optional[float] = None
    max_price: Optional[float] = None

    # Market Cap (in Crores)
    min_market_cap: Optional[float] = None
    max_market_cap: Optional[float] = None

    # PE Ratio
    min_pe: Optional[float] = None
    max_pe: Optional[float] = None

    # Beta
    min_beta: Optional[float] = None
    max_beta: Optional[float] = None

    # RSI
    min_rsi: Optional[float] = None
    max_rsi: Optional[float] = None

    # Moving Average signals (boolean toggles)
    price_above_sma20: bool = False
    price_above_sma50: bool = False
    price_above_sma200: bool = False
    sma50_above_sma200: bool = False   # Golden cross zone
    price_below_sma200: bool = False   # Potential value/distressed

    # MACD
    macd_bullish: bool = False   # MACD > signal line
    macd_bearish: bool = False

    # Returns (%)
    min_daily_return: Optional[float] = None
    max_daily_return: Optional[float] = None
    min_return_1m: Optional[float] = None
    max_return_1m: Optional[float] = None
    min_return_3m: Optional[float] = None
    max_return_3m: Optional[float] = None
    min_return_1y: Optional[float] = None
    max_return_1y: Optional[float] = None

    # Volume
    min_avg_volume: Optional[int] = None

    # Max Drawdown (%)  — more negative = bigger drawdown
    max_drawdown_threshold: Optional[float] = None  # e.g. -30 means drawdown worse than -30%

    # Sorting
    sort_by: str = "market_cap"
    sort_asc: bool = False

    # Pagination
    page: int = 1
    page_size: int = 50


def apply_filters(merged_df: pd.DataFrame, filters: ScreenerFilters) -> pd.DataFrame:
    """
    merged_df: DataFrame with all columns from stock_meta + indicators joined.
    Returns filtered, sorted, paginated result.
    """
    df = merged_df.copy()

    # Exchange
    if filters.exchanges:
        df = df[df["exchange"].isin(filters.exchanges)]

    # Sector
    if filters.sectors:
        df = df[df["sector"].isin(filters.sectors)]

    # Industry (handle both column name variants)
    if filters.industries:
        ind_col = "industry" if "industry" in df.columns else "industry_raw"
        if ind_col in df.columns:
            df = df[df[ind_col].isin(filters.industries)]

    # Price
    if filters.min_price is not None:
        df = df[df["last_price"] >= filters.min_price]
    if filters.max_price is not None:
        df = df[df["last_price"] <= filters.max_price]

    # Market Cap (convert from absolute to Crores: 1 Cr = 10_000_000)
    if filters.min_market_cap is not None:
        df = df[df["market_cap"] >= filters.min_market_cap * 1e7]
    if filters.max_market_cap is not None:
        df = df[df["market_cap"] <= filters.max_market_cap * 1e7]

    # PE
    if filters.min_pe is not None:
        df = df[df["pe_ratio"] >= filters.min_pe]
    if filters.max_pe is not None:
        df = df[df["pe_ratio"] <= filters.max_pe]

    # Beta
    if filters.min_beta is not None:
        df = df[df["beta"] >= filters.min_beta]
    if filters.max_beta is not None:
        df = df[df["beta"] <= filters.max_beta]

    # RSI
    if filters.min_rsi is not None:
        df = df[df["rsi_14"] >= filters.min_rsi]
    if filters.max_rsi is not None:
        df = df[df["rsi_14"] <= filters.max_rsi]

    # MA signals
    if filters.price_above_sma20:
        df = df[df["last_price"] > df["sma_20"]]
    if filters.price_above_sma50:
        df = df[df["last_price"] > df["sma_50"]]
    if filters.price_above_sma200:
        df = df[df["last_price"] > df["sma_200"]]
    if filters.price_below_sma200:
        df = df[df["last_price"] < df["sma_200"]]
    if filters.sma50_above_sma200:
        df = df[df["sma_50"] > df["sma_200"]]

    # MACD
    if filters.macd_bullish:
        df = df[df["macd"] > df["macd_signal"]]
    if filters.macd_bearish:
        df = df[df["macd"] < df["macd_signal"]]

    # Returns
    if filters.min_daily_return is not None:
        df = df[df["daily_return"] >= filters.min_daily_return]
    if filters.max_daily_return is not None:
        df = df[df["daily_return"] <= filters.max_daily_return]
    if filters.min_return_1m is not None:
        df = df[df["return_1m"] >= filters.min_return_1m]
    if filters.max_return_1m is not None:
        df = df[df["return_1m"] <= filters.max_return_1m]
    if filters.min_return_3m is not None:
        df = df[df["return_3m"] >= filters.min_return_3m]
    if filters.max_return_3m is not None:
        df = df[df["return_3m"] <= filters.max_return_3m]
    if filters.min_return_1y is not None:
        df = df[df["return_1y"] >= filters.min_return_1y]
    if filters.max_return_1y is not None:
        df = df[df["return_1y"] <= filters.max_return_1y]

    # Volume
    if filters.min_avg_volume is not None:
        df = df[df["avg_volume_20d"] >= filters.min_avg_volume]

    # Max Drawdown
    if filters.max_drawdown_threshold is not None:
        df = df[df["max_drawdown_52w"] >= filters.max_drawdown_threshold]

    # Sort
    sort_col = filters.sort_by if filters.sort_by in df.columns else "market_cap"
    df = df.sort_values(sort_col, ascending=filters.sort_asc, na_position="last")

    # Paginate
    total = len(df)
    start = (filters.page - 1) * filters.page_size
    end = start + filters.page_size
    page_df = df.iloc[start:end]

    return page_df, total


def get_sector_summary(merged_df: pd.DataFrame) -> list[dict]:
    """Return stock count and avg PE per sector — for sector filter sidebar."""
    summary = (
        merged_df.groupby("sector")
        .agg(
            count=("symbol", "count"),
            avg_pe=("pe_ratio", "mean"),
            avg_return_1m=("return_1m", "mean"),
        )
        .reset_index()
        .sort_values("count", ascending=False)
    )
    # Replace NaN/Inf with None before serialising
    import math
    records = []
    for row in summary.to_dict(orient="records"):
        clean = {}
        for k, v in row.items():
            if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                clean[k] = None
            else:
                clean[k] = v
        records.append(clean)
    return records
