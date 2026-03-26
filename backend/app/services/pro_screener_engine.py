"""
Pro Screener Engine — Bloomberg-style weighted screener.

Two modes:
  1. Hard-filter mode: ALL required conditions must pass (classic screener)
  2. Score mode:       Conditions contribute weighted scores; rank by composite score

Scoring: score = (sum of weights where condition satisfied / total weight) × 100
"""

from __future__ import annotations
from typing import Optional
import pandas as pd

# ── Filter Catalog ─────────────────────────────────────────────────────────────
# type: "range" | "bool"
# available: False = column not in merged_df (show in UI but disabled)
# input_scale: user enters value in display units; multiply by scale for df comparison
# display_scale: df value * display_scale = display value (for rendering)

FILTER_CATALOG: dict[str, dict] = {
    # ── A. Market Data ──
    "price":          {"col": "last_price",      "label": "Price (₹)",               "category": "A", "cat_label": "Market Data",      "type": "range", "unit": "₹"},
    "market_cap":     {"col": "market_cap",       "label": "Market Cap (₹ Cr)",       "category": "A", "cat_label": "Market Data",      "type": "range", "unit": "₹Cr",     "input_scale": 1e7},
    "volume_20d":     {"col": "avg_volume_20d",   "label": "Avg Volume 20D (K)",      "category": "A", "cat_label": "Market Data",      "type": "range", "unit": "K",       "input_scale": 1000},
    "beta":           {"col": "beta",             "label": "Beta",                    "category": "A", "cat_label": "Market Data",      "type": "range", "unit": ""},
    "short_ratio":    {"col": "short_ratio",      "label": "Short Ratio",             "category": "A", "cat_label": "Market Data",      "type": "range", "unit": "",        "available": False},
    "bid_ask_spread": {"col": "bid_ask_spread",  "label": "Bid-Ask Spread (%)",      "category": "A", "cat_label": "Market Data",      "type": "range", "unit": "%",       "available": False},

    # ── C. Valuation ──
    "pe_ratio":       {"col": "pe_ratio",         "label": "P/E Ratio",               "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "x"},
    "forward_pe":     {"col": "forward_pe",       "label": "Forward P/E",             "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "x",       "available": False},
    "pb_ratio":       {"col": "price_to_book",    "label": "Price-to-Book",           "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "x",       "available": False},
    "ps_ratio":       {"col": "price_to_sales",   "label": "Price-to-Sales",          "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "x",       "available": False},
    "ev_ebitda":      {"col": "ev_ebitda",        "label": "EV/EBITDA",               "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "x",       "available": False},
    "ev_sales":       {"col": "ev_sales",         "label": "EV/Sales",                "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "x",       "available": False},
    "div_yield":      {"col": "dividend_yield",   "label": "Dividend Yield (%)",      "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "%",       "available": False},
    "fcf_yield_val":  {"col": "fcf_yield",        "label": "FCF Yield (%)",           "category": "C", "cat_label": "Valuation",        "type": "range", "unit": "%",       "available": False},

    # ── D. Profitability ──
    "roe":            {"col": "roe",              "label": "Return on Equity (%)",    "category": "D", "cat_label": "Profitability",    "type": "range", "unit": "%",       "available": False},
    "roa":            {"col": "roa",              "label": "Return on Assets (%)",    "category": "D", "cat_label": "Profitability",    "type": "range", "unit": "%",       "available": False},
    "roic":           {"col": "roic",             "label": "ROIC (%)",                "category": "D", "cat_label": "Profitability",    "type": "range", "unit": "%",       "available": False},
    "gross_margin":   {"col": "gross_margins",    "label": "Gross Margin (%)",        "category": "D", "cat_label": "Profitability",    "type": "range", "unit": "%",       "available": False},
    "op_margin":      {"col": "operating_margins","label": "Operating Margin (%)",    "category": "D", "cat_label": "Profitability",    "type": "range", "unit": "%",       "available": False},
    "net_margin":     {"col": "profit_margins",   "label": "Net Profit Margin (%)",   "category": "D", "cat_label": "Profitability",    "type": "range", "unit": "%",       "available": False},

    # ── E. Growth & Returns ──
    "return_1d":      {"col": "daily_return",     "label": "Return 1D (%)",           "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%"},
    "return_5d":      {"col": "return_5d",        "label": "Return 5D (%)",           "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%"},
    "return_1m":      {"col": "return_1m",        "label": "Return 1M (%)",           "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%"},
    "return_3m":      {"col": "return_3m",        "label": "Return 3M (%)",           "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%"},
    "return_1y":      {"col": "return_1y",        "label": "Return 1Y (%)",           "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%"},
    "return_6m":      {"col": "return_6m",        "label": "Return 6M (%)",           "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%",       "available": False},
    "rev_growth":     {"col": "revenue_growth",   "label": "Revenue Growth YoY (%)",  "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%",       "available": False},
    "eps_growth":     {"col": "earnings_growth",  "label": "EPS Growth (%)",          "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%",       "available": False},
    "ebitda_growth":  {"col": "ebitda_growth",    "label": "EBITDA Growth YoY (%)",   "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%",       "available": False},
    "forecast_growth":{"col": "forecast_growth",  "label": "Forecast EPS Growth (%)", "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "%",       "available": False},
    "eps":            {"col": "eps",              "label": "EPS (₹)",                 "category": "E", "cat_label": "Growth & Returns", "type": "range", "unit": "₹",       "available": False},

    # ── F. Leverage & Credit ──
    "debt_equity":    {"col": "debt_to_equity",   "label": "Debt / Equity",           "category": "F", "cat_label": "Leverage & Credit","type": "range", "unit": "x",       "available": False},
    "net_debt_ebitda":{"col": "net_debt_ebitda",  "label": "Net Debt / EBITDA",       "category": "F", "cat_label": "Leverage & Credit","type": "range", "unit": "x",       "available": False},
    "total_debt":     {"col": "total_debt",       "label": "Total Debt (₹ Cr)",       "category": "F", "cat_label": "Leverage & Credit","type": "range", "unit": "₹Cr",     "available": False, "input_scale": 1e7},
    "int_coverage":   {"col": "interest_coverage","label": "Interest Coverage",       "category": "F", "cat_label": "Leverage & Credit","type": "range", "unit": "x",       "available": False},
    "volatility":     {"col": "volatility_52w",   "label": "Volatility 52W (%)",      "category": "F", "cat_label": "Leverage & Credit","type": "range", "unit": "%",       "available": False},

    # ── G. Cash Flow ──
    "fcf_yield":      {"col": "fcf_yield",        "label": "FCF Yield (%)",           "category": "G", "cat_label": "Cash Flow",        "type": "range", "unit": "%",       "available": False},
    "free_cashflow":  {"col": "free_cashflow",    "label": "Free Cash Flow (₹ Cr)",   "category": "G", "cat_label": "Cash Flow",        "type": "range", "unit": "₹Cr",     "available": False, "input_scale": 1e7},
    "op_cashflow":    {"col": "operating_cashflow","label": "Operating Cash Flow (Cr)","category": "G", "cat_label": "Cash Flow",       "type": "range", "unit": "₹Cr",     "available": False},
    "capex":          {"col": "capex",            "label": "Capex (₹ Cr)",            "category": "G", "cat_label": "Cash Flow",        "type": "range", "unit": "₹Cr",     "available": False},

    # ── H. Analyst & Sentiment ──
    "target_upside":  {"col": "target_upside",    "label": "Target Price Upside (%)", "category": "H", "cat_label": "Analyst & Sentiment","type":"range","unit": "%",       "available": False},
    "analyst_buy_pct":{"col": "analyst_buy_pct",  "label": "Analyst Buy % (≥ x%)",   "category": "H", "cat_label": "Analyst & Sentiment","type":"range","unit": "%",       "available": False},
    "analyst_rating": {"col": "analyst_mean_rating","label": "Analyst Rating (1=Strong Buy → 5=Sell)","category": "H", "cat_label": "Analyst & Sentiment","type":"range","unit": "",       "available": False},
    "earnings_rev":   {"col": "earnings_revisions","label": "Earnings Revisions (%)", "category": "H", "cat_label": "Analyst & Sentiment","type":"range","unit": "%",       "available": False},
    "earnings_surp":  {"col": "earnings_surprise", "label": "Earnings Surprise (%)",  "category": "H", "cat_label": "Analyst & Sentiment","type":"range","unit": "%",       "available": False},

    # ── I. Ownership ──
    "insider_pct":    {"col": "insider_pct",      "label": "Insider Ownership (%)",   "category": "I", "cat_label": "Ownership",        "type": "range", "unit": "%",       "available": False},
    "inst_pct":       {"col": "inst_pct",         "label": "Institutional Holding (%)","category": "I","cat_label": "Ownership",        "type": "range", "unit": "%",       "available": False},

    # ── K. Technical — range ──
    "rsi":            {"col": "rsi_14",           "label": "RSI (14)",                "category": "K", "cat_label": "Technical",        "type": "range", "unit": ""},
    "macd_hist_val":  {"col": "macd_hist",        "label": "MACD Histogram",          "category": "K", "cat_label": "Technical",        "type": "range", "unit": ""},
    "drawdown":       {"col": "max_drawdown_52w", "label": "Max Drawdown 52W (%)",    "category": "K", "cat_label": "Technical",        "type": "range", "unit": "%"},

    # ── K. Technical — boolean signals ──
    "above_sma20":    {"col": None, "label": "Price above SMA 20",                    "category": "K", "cat_label": "Technical",        "type": "bool"},
    "above_sma50":    {"col": None, "label": "Price above SMA 50",                    "category": "K", "cat_label": "Technical",        "type": "bool"},
    "above_sma200":   {"col": None, "label": "Price above SMA 200",                   "category": "K", "cat_label": "Technical",        "type": "bool"},
    "above_ema20":    {"col": None, "label": "Price above EMA 20",                    "category": "K", "cat_label": "Technical",        "type": "bool"},
    "above_ema50":    {"col": None, "label": "Price above EMA 50",                    "category": "K", "cat_label": "Technical",        "type": "bool"},
    "above_ema200":   {"col": None, "label": "Price above EMA 200",                   "category": "K", "cat_label": "Technical",        "type": "bool"},
    "golden_cross":   {"col": None, "label": "Golden Cross (SMA50 > SMA200)",         "category": "K", "cat_label": "Technical",        "type": "bool"},
    "death_cross":    {"col": None, "label": "Death Cross (SMA50 < SMA200)",          "category": "K", "cat_label": "Technical",        "type": "bool"},
    "macd_bullish":   {"col": None, "label": "MACD Bullish (MACD > Signal)",          "category": "K", "cat_label": "Technical",        "type": "bool"},
    "macd_bearish":   {"col": None, "label": "MACD Bearish (MACD < Signal)",          "category": "K", "cat_label": "Technical",        "type": "bool"},
    "rsi_oversold":   {"col": None, "label": "RSI Oversold (< 30)",                   "category": "K", "cat_label": "Technical",        "type": "bool"},
    "rsi_overbought": {"col": None, "label": "RSI Overbought (> 70)",                 "category": "K", "cat_label": "Technical",        "type": "bool"},
    "momentum_up":    {"col": None, "label": "Momentum (Return 1M > 3M > 0)",         "category": "K", "cat_label": "Technical",        "type": "bool"},
    "rel_strength":   {"col": "relative_strength","label": "Relative Strength (vs Index)","category": "K", "cat_label": "Technical",   "type": "range", "unit": "",        "available": False},

    # ── J. ESG ──
    "esg_score":      {"col": "esg_score",        "label": "ESG Score",               "category": "J", "cat_label": "ESG",              "type": "range", "unit": "",        "available": False},
    "esg_controversies":{"col": "esg_controversies","label": "ESG Controversies (lower is better)","category": "J", "cat_label": "ESG","type": "range", "unit": "",        "available": False},
}

# ── Boolean filter evaluation functions ────────────────────────────────────────
def _safe(df: pd.DataFrame, a: str, b: str, gt: bool = True) -> pd.Series:
    """Compare two columns safely, returning False where either is NaN."""
    if a not in df.columns or b not in df.columns:
        return pd.Series(False, index=df.index)
    ca, cb = pd.to_numeric(df[a], errors="coerce"), pd.to_numeric(df[b], errors="coerce")
    mask = ca > cb if gt else ca < cb
    return mask.fillna(False)

def _safe_val(df: pd.DataFrame, col: str, op, val: float) -> pd.Series:
    if col not in df.columns:
        return pd.Series(False, index=df.index)
    c = pd.to_numeric(df[col], errors="coerce")
    return op(c, val).fillna(False)

BOOL_FUNCS: dict[str, callable] = {
    "above_sma20":   lambda df: _safe(df, "last_price", "sma_20"),
    "above_sma50":   lambda df: _safe(df, "last_price", "sma_50"),
    "above_sma200":  lambda df: _safe(df, "last_price", "sma_200"),
    "above_ema20":   lambda df: _safe(df, "last_price", "ema_20"),
    "above_ema50":   lambda df: _safe(df, "last_price", "ema_50"),
    "above_ema200":  lambda df: _safe(df, "last_price", "ema_200"),
    "golden_cross":  lambda df: _safe(df, "sma_50", "sma_200"),
    "death_cross":   lambda df: _safe(df, "sma_50", "sma_200", gt=False),
    "macd_bullish":  lambda df: _safe(df, "macd", "macd_signal"),
    "macd_bearish":  lambda df: _safe(df, "macd", "macd_signal", gt=False),
    "rsi_oversold":  lambda df: _safe_val(df, "rsi_14", lambda c, v: c < v, 30),
    "rsi_overbought":lambda df: _safe_val(df, "rsi_14", lambda c, v: c > v, 70),
    "momentum_up":   lambda df: (
        _safe(df, "return_1m", "return_3m") &
        _safe_val(df, "return_3m", lambda c, v: c > v, 0)
    ),
}


# ── Condition dataclass ────────────────────────────────────────────────────────
class FilterCondition:
    def __init__(
        self,
        filter_id: str,
        min_val: Optional[float] = None,
        max_val: Optional[float] = None,
        weight: float = 0.0,   # decimal 0–1; caller ensures sum == 1
        required: bool = True,
    ):
        self.filter_id = filter_id
        self.min_val = min_val
        self.max_val = max_val
        self.weight = max(0.0, min(1.0, float(weight)))
        self.required = required


def _evaluate(df: pd.DataFrame, cond: FilterCondition) -> pd.Series:
    """Boolean Series: True if stock satisfies condition."""
    meta = FILTER_CATALOG.get(cond.filter_id)
    if not meta:
        return pd.Series(True, index=df.index)

    if meta["type"] == "bool":
        fn = BOOL_FUNCS.get(cond.filter_id)
        if fn is None:
            return pd.Series(True, index=df.index)
        return fn(df)

    # range filter
    col = meta["col"]
    if not col or col not in df.columns:
        return pd.Series(True, index=df.index)

    scale = meta.get("input_scale", 1)
    series = pd.to_numeric(df[col], errors="coerce")
    mask = pd.Series(True, index=df.index)
    if cond.min_val is not None:
        mask &= series >= cond.min_val * scale
    if cond.max_val is not None:
        mask &= series <= cond.max_val * scale
    return mask.fillna(False)


def apply_pro_screen(
    merged_df: pd.DataFrame,
    sector: Optional[str],
    conditions: list[FilterCondition],
    score_mode: bool = False,
    sort_by: str = "score",
    sort_asc: bool = False,
    page: int = 1,
    page_size: int = 50,
) -> tuple[pd.DataFrame, int]:
    """
    Apply sector filter + all conditions with weighted scoring.

    Hard-filter mode (score_mode=False):
        Stocks must pass ALL required conditions. Non-required conditions
        still contribute to the score used for ranking.

    Score mode (score_mode=True):
        No stock is excluded. All stocks are ranked by composite score.
    """
    df = merged_df.copy()

    # Sector scope
    if sector and sector != "all":
        df = df[df["sector"].str.lower() == sector.lower()]

    if df.empty or not conditions:
        total = len(df)
        out = df.iloc[(page - 1) * page_size: page * page_size]
        out = out.assign(score=0)
        return out, total

    # ── Evaluate every condition ──────────────────────────────────────────────
    masks: dict[str, pd.Series] = {}
    for cond in conditions:
        masks[cond.filter_id] = _evaluate(df, cond)

    # ── Hard-filter: exclude failing required conditions ──────────────────────
    if not score_mode:
        combined = pd.Series(True, index=df.index)
        for cond in conditions:
            if cond.required:
                combined &= masks[cond.filter_id]
        df = df[combined]

    # ── Compute weighted score ────────────────────────────────────────────────
    # Weights sum to 1 → score = sum(weight_i × passes_i) × 100  (0–100 range)
    score = pd.Series(0.0, index=df.index)
    for cond in conditions:
        if cond.filter_id in masks:
            m = masks[cond.filter_id].reindex(df.index).fillna(False)
            score += m.astype(float) * cond.weight
    score = (score * 100).round(1)

    df = df.copy()
    df["score"] = score

    # ── Sort ──────────────────────────────────────────────────────────────────
    if sort_by == "score" or sort_by not in df.columns:
        df = df.sort_values("score", ascending=sort_asc, na_position="last")
    else:
        df = df.sort_values(sort_by, ascending=sort_asc, na_position="last")

    total = len(df)
    start = (page - 1) * page_size
    page_df = df.iloc[start: start + page_size]

    return page_df, total


def get_catalog_by_category() -> list[dict]:
    """Return catalog grouped by category, with availability status."""
    from collections import OrderedDict
    groups: OrderedDict[str, dict] = OrderedDict()
    for fid, meta in FILTER_CATALOG.items():
        cat = meta["category"]
        cat_label = meta["cat_label"]
        if cat not in groups:
            groups[cat] = {"category": cat, "label": cat_label, "filters": []}
        groups[cat]["filters"].append({
            "id":        fid,
            "label":     meta["label"],
            "type":      meta["type"],
            "unit":      meta.get("unit", ""),
            "available": meta.get("available", True),
        })
    return list(groups.values())
