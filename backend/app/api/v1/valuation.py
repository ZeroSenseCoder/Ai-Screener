"""
/api/v1/valuation — multi-model equity valuation engine.
Financial inputs sourced from:
  1. screener.in  (BSE/NSE mandatory XBRL filings — P&L, BS, CF)
  2. yfinance     (market data, beta, dividend, growth estimates)
"""

import math
import logging
from typing import Optional, List
from fastapi import APIRouter, HTTPException
from fastapi.responses import Response
from pydantic import BaseModel, Field
import yfinance as yf

from app.core.state import app_state
from app.services import valuation_engine as ve
from app.services.bse_filings import fetch_bse_filings
from app.services.dcf_excel_generator import generate_dcf_excel
from app.services.generators_dispatcher import generate_for_model

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/valuation", tags=["Valuation"])


def _clean(obj):
    if isinstance(obj, float):
        return None if (math.isnan(obj) or math.isinf(obj)) else round(obj, 4)
    if isinstance(obj, dict):
        return {k: _clean(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_clean(v) for v in obj]
    return obj


def _get_merged_row(yf_symbol: str) -> dict:
    """Pull stock data from in-memory merged_df."""
    merged = app_state.merged_df
    if merged is None:
        return {}
    row = merged[merged["yf_symbol"] == yf_symbol]
    if row.empty:
        return {}
    r = row.iloc[0]
    def g(col):
        try:
            v = r.get(col)
            if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
                return None
            return v
        except Exception:
            return None
    return {
        "last_price": g("last_price"),
        "market_cap": g("market_cap"),
        "pe_ratio": g("pe_ratio"),
        "forward_pe": g("forward_pe"),
        "beta": g("beta"),
        "dividend_yield": g("dividend_yield"),
        "price_to_book": g("price_to_book"),
        "debt_to_equity": g("debt_to_equity"),
        "revenue_growth": g("revenue_growth"),
        "earnings_growth": g("earnings_growth"),
        "profit_margins": g("profit_margins"),
        "roe": g("roe"),
        "roa": g("roa"),
    }


def _fetch_financials(yf_symbol: str, base_symbol: str = "", isin: str = "") -> dict:
    """
    Fetch financial data from two sources and merge:
      1. screener.in  → official BSE/NSE filing data (P&L, BS, CF)
      2. yfinance     → market data, beta, dividends, growth rates, volatility
    Filing data takes priority for income statement / balance sheet items.
    yfinance fills market-level fields and growth estimates.
    """
    merged_row = _get_merged_row(yf_symbol)

    # ── Source 1: BSE/NSE filings via screener.in ─────────────────────────────
    filing: dict = {}
    if base_symbol:
        try:
            filing = fetch_bse_filings(base_symbol, isin=isin)
        except Exception as e:
            logger.debug(f"bse_filings failed for {base_symbol}: {e}")

    # ── Source 2: yfinance ─────────────────────────────────────────────────────
    try:
        t = yf.Ticker(yf_symbol)
        info = t.info or {}

        cf_df = None
        bs_df = None
        try:
            cf_df = t.cashflow
        except Exception:
            pass
        try:
            bs_df = t.balance_sheet
        except Exception:
            pass

        def ci(key, default=None):
            v = info.get(key)
            if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
                return default
            return v

        def from_df(df, *labels):
            if df is None or df.empty:
                return None
            for label in labels:
                for idx in df.index:
                    if label.lower() in str(idx).lower():
                        try:
                            v = float(df.loc[idx].iloc[0])
                            if not math.isnan(v):
                                return v
                        except Exception:
                            pass
            return None

        yf_shares    = ci("sharesOutstanding")
        yf_price     = ci("currentPrice") or ci("regularMarketPrice") or merged_row.get("last_price")
        yf_fcf       = ci("freeCashflow")
        yf_op_cf     = ci("operatingCashflow")
        yf_revenue   = ci("totalRevenue")
        yf_ebitda    = ci("ebitda")
        yf_net_inc   = ci("netIncomeToCommon")
        yf_debt      = ci("totalDebt") or 0
        yf_cash      = ci("totalCash") or 0
        yf_tot_ass   = ci("totalAssets")
        yf_tot_liab  = ci("totalLiab")
        yf_book_ps   = ci("bookValue")
        yf_dps       = ci("lastDividendValue") or (ci("dividendRate") or 0)
        yf_eps       = ci("trailingEps")
        yf_roe       = ci("returnOnEquity") or merged_row.get("roe")
        yf_roa       = ci("returnOnAssets")
        yf_rev_g     = ci("revenueGrowth")    or merged_row.get("revenue_growth") or 0
        yf_earn_g    = ci("earningsGrowth")   or merged_row.get("earnings_growth") or 0
        yf_mkt_cap   = ci("marketCap")        or merged_row.get("market_cap")
        yf_ev        = ci("enterpriseValue")
        yf_sector    = info.get("sector", "")

        # Volatility from 1Y price history
        volatility = None
        try:
            hist = t.history(period="1y", interval="1d", auto_adjust=True)
            if not hist.empty and len(hist) > 30:
                import numpy as np
                rets = hist["Close"].pct_change().dropna()
                volatility = float(rets.std() * (252 ** 0.5))
        except Exception:
            pass

    except Exception as e:
        logger.warning(f"_fetch_financials yf ({yf_symbol}): {e}")
        yf_shares = yf_price = yf_fcf = yf_op_cf = yf_revenue = None
        yf_ebitda = yf_net_inc = yf_dps = yf_eps = yf_roe = yf_roa = None
        yf_rev_g = yf_earn_g = 0
        yf_debt = yf_cash = 0
        yf_tot_ass = yf_tot_liab = yf_book_ps = yf_mkt_cap = yf_ev = None
        yf_sector = ""
        volatility = None

    def _pick(filing_val, yf_val, default=None):
        """Use filing value if available, else yfinance, else default."""
        if filing_val is not None:
            return filing_val
        if yf_val is not None:
            return yf_val
        return default

    shares  = _pick(filing.get("shares"),  yf_shares)
    price   = yf_price  # always use live market price from yfinance

    # Income statement: prefer yfinance (consolidated) over screener.in (standalone).
    # screener.in returns parent-only data for large conglomerates (e.g. Reliance),
    # while the market prices the consolidated entity — using standalone inputs
    # against a consolidated market cap produces grossly wrong DCF/PEG/EVA values.
    revenue = _pick(yf_revenue, filing.get("revenue"))
    ebitda  = _pick(yf_ebitda,  filing.get("ebitda"))
    net_inc = _pick(yf_net_inc, filing.get("net_income"))
    op_cf   = _pick(yf_op_cf,   filing.get("operating_cf"))
    fcf     = _pick(yf_fcf,     filing.get("fcf"))
    eps     = _pick(yf_eps,     filing.get("eps"))

    # Balance sheet: screener.in filing priority (more reliable for Indian companies)
    debt    = _pick(filing.get("total_debt"),  yf_debt, 0)
    cash    = _pick(filing.get("cash"),        yf_cash, 0)
    tot_ass = _pick(filing.get("total_assets"), yf_tot_ass)
    tot_liab= _pick(filing.get("total_liabilities"), yf_tot_liab)
    equity  = filing.get("total_equity")
    bvps    = _pick(filing.get("book_value_per_share"), yf_book_ps)
    book_tot= _pick(filing.get("book_value_total"), (yf_book_ps * shares if yf_book_ps and shares else None))
    dps     = yf_dps   # dividends from yfinance are more reliable
    roe     = _pick(filing.get("roe"), yf_roe)
    roa     = yf_roa
    nopat   = _pick(yf_net_inc, filing.get("nopat")) or (net_inc or 0)
    invested_cap = _pick(filing.get("invested_capital"), (book_tot or 0) + debt)
    noi     = filing.get("noi") or (ebitda * 0.9 if ebitda else None)

    # Revenue growth: prefer yfinance (forward-looking) or compute from filing
    rev_growth  = yf_rev_g or 0
    # earn_growth: yfinance earningsGrowth can be unreliable (e.g. 0.6% for Reliance).
    # Use the higher of earnings growth and revenue growth as a floor.
    earn_growth = max(yf_earn_g or 0, yf_rev_g or 0) if (yf_earn_g or yf_rev_g) else 0

    rev_ps = (revenue / shares) if revenue and shares and shares > 0 else None

    filing_source = filing.get("_source", "yfinance only")

    return {
        "price":              price,
        "shares":             shares,
        "market_cap":         yf_mkt_cap,
        "enterprise_value":   yf_ev,
        "fcf":                fcf,
        "fcfe":               fcf,
        "operating_cf":       op_cf,
        "revenue":            revenue,
        "ebitda":             ebitda,
        "net_income":         net_inc,
        "nopat":              nopat,
        "total_debt":         debt,
        "cash":               cash,
        "total_assets":       tot_ass,
        "total_liabilities":  tot_liab,
        "total_equity":       equity,
        "book_value_total":   book_tot,
        "book_value_per_share": bvps,
        "invested_capital":   invested_cap,
        "dps":                dps,
        "eps":                eps,
        "roe":                roe,
        "roa":                roa,
        "revenue_growth":     rev_growth,
        "earnings_growth":    earn_growth,
        "revenue_per_share":  rev_ps,
        "noi":                noi,
        "volatility_annual":  volatility,
        "sector":             yf_sector,
        "data_source":        filing_source,
    }


def _suggest_wacc(beta: Optional[float], debt_ratio: float = 0.3) -> float:
    """Estimate WACC for an Indian stock."""
    rf     = 0.071        # India 10Y G-sec ~7.1%
    erp    = 0.055        # Equity risk premium
    b      = beta or 1.0
    ke     = rf + b * erp
    kd     = 0.09         # average Indian debt cost
    tax    = 0.25
    wacc   = ke * (1 - debt_ratio) + kd * (1 - tax) * debt_ratio
    return round(wacc, 4)


def _suggest_ke(beta: Optional[float]) -> float:
    rf  = 0.071
    erp = 0.055
    b   = beta or 1.0
    return round(rf + b * erp, 4)


# ── Request model ─────────────────────────────────────────────────────────────

class RunValuationRequest(BaseModel):
    model: str
    params: dict = Field(default_factory=dict)


# ── Endpoints ─────────────────────────────────────────────────────────────────

@router.get("/{symbol}/inputs")
def get_valuation_inputs(symbol: str):
    """
    Return all financial data needed for valuation models + suggested parameters.
    """
    merged = app_state.merged_df
    if merged is None:
        raise HTTPException(503, "Universe not loaded")

    upper = symbol.upper()
    row = merged[merged["symbol"] == upper]
    if row.empty:
        row = merged[merged["yf_symbol"] == upper]
    if row.empty:
        raise HTTPException(404, f"Symbol '{symbol}' not found")

    r = row.iloc[0]
    yf_sym    = r["yf_symbol"]
    base_sym  = str(r.get("symbol", upper))
    isin_val  = str(r.get("isin", ""))
    price     = float(r.get("last_price") or 0)
    beta      = r.get("beta")

    fin = _fetch_financials(yf_sym, base_symbol=base_sym, isin=isin_val)
    price  = fin.get("price") or price

    wacc_suggested = _suggest_wacc(beta)
    ke_suggested   = _suggest_ke(beta)

    # Which models are applicable?
    models_available = {
        "dcf_fcff":       fin.get("fcf") is not None,
        "dcf_fcfe":       fin.get("fcfe") is not None,
        "dcf_multistage": fin.get("fcf") is not None,
        "gordon_growth":  (fin.get("dps") or 0) > 0,
        "ddm_multistage": (fin.get("dps") or 0) > 0,
        "residual_income":fin.get("book_value_per_share") is not None and fin.get("roe") is not None,
        "capitalized_earnings": (fin.get("eps") or 0) > 0,
        "nav":            fin.get("total_assets") is not None,
        "liquidation":    fin.get("total_assets") is not None,
        "eva":            fin.get("nopat") is not None and fin.get("invested_capital", 0) > 0,
        "peg":            (fin.get("eps") or 0) > 0 and (fin.get("earnings_growth") or 0) > 0,
        "trading_comps":  True,
        "lbo":            fin.get("ebitda") is not None,
        "black_scholes":  fin.get("volatility_annual") is not None,
        "pb_banks":       fin.get("book_value_per_share") is not None and fin.get("roe") is not None,
        "cap_rate":       fin.get("noi") is not None,
        "sum_of_parts":   True,
        "vc_method":      True,
        "precedent_transactions": fin.get("ebitda") is not None,
        "replacement_cost":       fin.get("total_assets") is not None,
        "excess_earnings":        fin.get("net_income") is not None and fin.get("total_assets") is not None,
        "cfroi":                  fin.get("operating_cf") is not None,
        "revenue_multiple":       fin.get("revenue") is not None,
        "user_based":             True,
    }

    return _clean({
        "symbol": upper,
        "yf_symbol": yf_sym,
        "current_price": price,
        "sector": fin.get("sector", r.get("sector", "")),
        "financials": fin,
        "suggested": {
            "wacc": wacc_suggested,
            "cost_of_equity": ke_suggested,
            "risk_free_rate": 0.071,
            "terminal_growth": 0.05,
            "growth_stage1": max(0.05, (fin.get("revenue_growth") or 0.10)),
            "years": 5,
        },
        "models_available": models_available,
    })


@router.post("/{symbol}/run")
def run_valuation(symbol: str, req: RunValuationRequest):
    """
    Run a specific valuation model with provided parameters.
    Returns intrinsic value, upside %, and model-specific details.
    """
    merged = app_state.merged_df
    if merged is None:
        raise HTTPException(503, "Universe not loaded")

    upper = symbol.upper()
    row = merged[merged["symbol"] == upper]
    if row.empty:
        row = merged[merged["yf_symbol"] == upper]
    if row.empty:
        raise HTTPException(404, f"Symbol '{symbol}' not found")

    r        = row.iloc[0]
    yf_sym   = r["yf_symbol"]
    base_sym = str(r.get("symbol", upper))
    isin_val = str(r.get("isin", ""))
    price    = float(r.get("last_price") or 0)
    beta     = r.get("beta")

    fin   = _fetch_financials(yf_sym, base_symbol=base_sym, isin=isin_val)
    price = fin.get("price") or price

    p = req.params  # user-supplied parameters, fall back to sensible defaults

    wacc   = p.get("wacc",   _suggest_wacc(beta))
    ke     = p.get("cost_of_equity", _suggest_ke(beta))
    tg     = p.get("terminal_growth", 0.05)
    g1     = p.get("growth_stage1", max(0.05, (fin.get("revenue_growth") or 0.10)))
    g2     = p.get("growth_stage2", max(0.04, g1 * 0.6))
    years  = int(p.get("years", 5))
    y1     = int(p.get("stage1_years", 5))
    y2     = int(p.get("stage2_years", 5))

    fcf    = p.get("fcf",   fin.get("fcf") or 0)
    fcfe   = p.get("fcfe",  fin.get("fcfe") or fcf)
    debt   = p.get("total_debt", fin.get("total_debt") or 0)
    cash   = p.get("cash",       fin.get("cash") or 0)
    shares = p.get("shares", fin.get("shares") or 0)
    eps    = p.get("eps",    fin.get("eps") or 0)
    dps    = p.get("dps",    fin.get("dps") or 0)
    bvps   = p.get("book_value_per_share", fin.get("book_value_per_share") or 0)
    roe    = p.get("roe",    fin.get("roe") or 0.12)
    ebitda = p.get("ebitda", fin.get("ebitda") or 0)
    nopat  = p.get("nopat",  fin.get("nopat") or 0)
    ic     = p.get("invested_capital", fin.get("invested_capital") or 0)
    tot_assets = p.get("total_assets", fin.get("total_assets") or 0)
    tot_liab   = p.get("total_liabilities", fin.get("total_liabilities") or 0)
    earn_g_pct = p.get("earnings_growth_pct", (fin.get("earnings_growth") or 0.10) * 100)
    vol    = p.get("volatility", fin.get("volatility_annual") or 0.30)
    rev_ps = fin.get("revenue_per_share") or (fin.get("revenue", 0) / shares if shares else 0)

    model = req.model
    result = {}

    if model == "dcf_fcff":
        result = ve.dcf_fcff(fcf, g1, tg, wacc, years, debt, cash, shares, price)

    elif model == "dcf_fcfe":
        result = ve.dcf_fcfe(fcfe, g1, tg, ke, years, shares, price)

    elif model == "dcf_multistage":
        result = ve.dcf_multistage(fcf, g1, g2, tg, wacc, y1, y2, debt, cash, shares, price)

    elif model == "gordon_growth":
        result = ve.gordon_growth(dps, ke, tg, price)

    elif model == "ddm_multistage":
        result = ve.ddm_multistage(dps, g1, tg, ke, years, price)

    elif model == "residual_income":
        bv_growth = p.get("bv_growth", roe * (1 - (p.get("payout_ratio", 0.3))))
        result = ve.residual_income(bvps, roe, ke, bv_growth, years, tg, price)

    elif model == "capitalized_earnings":
        req_ret = p.get("required_return", ke)
        eg_norm = p.get("eps_growth", fin.get("earnings_growth") or 0.05)
        result = ve.capitalized_earnings(eps, req_ret, eg_norm, price)

    elif model == "nav":
        goodwill = p.get("goodwill", 0)
        result = ve.nav_model(tot_assets, tot_liab, shares, price, goodwill)

    elif model == "liquidation":
        assets = fin.get("total_assets") or 0
        liq_split = p.get("asset_split", {})
        result = ve.liquidation_value(
            cash=liq_split.get("cash", cash),
            receivables=liq_split.get("receivables", assets * 0.15),
            inventory=liq_split.get("inventory", assets * 0.10),
            ppe=liq_split.get("ppe", assets * 0.40),
            other_assets=liq_split.get("other", assets * 0.20 - cash),
            total_liabilities=tot_liab,
            shares=shares,
            current_price=price,
            cash_rate=p.get("cash_rate", 1.0),
            receivables_rate=p.get("receivables_rate", 0.85),
            inventory_rate=p.get("inventory_rate", 0.50),
            ppe_rate=p.get("ppe_rate", 0.60),
            other_rate=p.get("other_rate", 0.25),
        )

    elif model == "eva":
        result = ve.eva_model(nopat, ic, wacc, g1, years, tg, debt, cash, shares, price)

    elif model == "peg":
        result = ve.peg_valuation(eps, earn_g_pct, p.get("target_peg", 1.0), price)

    elif model == "trading_comps":
        # Find sector peers from merged_df (only select columns that actually exist)
        sector = str(r.get("sector", ""))
        peers_df = merged[
            (merged["sector"] == sector) &
            (merged["yf_symbol"] != yf_sym) &
            (merged["pe_ratio"].notna())
        ].head(30)
        _peer_cols = [c for c in ["pe_ratio", "price_to_book", "price_to_sales", "ev_ebitda"]
                      if c in peers_df.columns]
        if _peer_cols and not peers_df.empty:
            peers = peers_df[_peer_cols].where(peers_df[_peer_cols].notna(), None).to_dict(orient="records")
        else:
            peers = []
        stock_data = {
            "eps": eps, "book_value_per_share": bvps, "revenue_per_share": rev_ps,
            "pe_ratio": r.get("pe_ratio"), "price_to_book": r.get("price_to_book"),
        }
        result = ve.trading_comps(stock_data, peers, price)

    elif model == "lbo":
        result = ve.lbo_model(
            ebitda=ebitda,
            entry_ev_multiple=p.get("entry_multiple", 8.0),
            exit_ev_multiple=p.get("exit_multiple", 10.0),
            debt_ratio=p.get("debt_ratio", 0.60),
            interest_rate=p.get("interest_rate", 0.09),
            ebitda_growth=p.get("ebitda_growth", fin.get("revenue_growth") or 0.08),
            hold_years=int(p.get("hold_years", 5)),
            shares=shares,
            cash=cash,
            current_price=price,
        )

    elif model == "black_scholes":
        result = ve.black_scholes(
            spot=price,
            strike=p.get("strike", price),
            volatility=vol,
            risk_free=p.get("risk_free", 0.071),
            time_years=p.get("time_years", 1.0),
            option_type=p.get("option_type", "call"),
            dividend_yield=fin.get("dividend_yield") or 0.0,
        )

    elif model == "pb_banks":
        result = ve.pb_banks(roe, ke, tg, bvps, price)

    elif model == "cap_rate":
        noi = p.get("noi", fin.get("noi") or ebitda * 0.9 if ebitda else 0)
        result = ve.cap_rate_model(noi, p.get("cap_rate", 0.07), debt, cash, shares, price)

    elif model == "sum_of_parts":
        segments = p.get("segments", [{"name": "Core Business", "ebitda": ebitda, "multiple": 8.0}])
        result = ve.sum_of_parts(segments, cash, debt, shares, price)

    elif model == "vc_method":
        result = ve.vc_method(
            projected_revenue=p.get("projected_revenue", (fin.get("revenue") or 0) * (1.2 ** int(p.get("years", 5)))),
            terminal_revenue_multiple=p.get("terminal_revenue_multiple", 3.0),
            target_return=p.get("target_return", 0.25),
            investment=p.get("investment", price * shares * 0.10 if shares else 0),
            years=int(p.get("years", 5)),
            shares=shares,
            current_price=price,
        )

    elif model == "precedent_transactions":
        result = ve.precedent_transactions(
            revenue=fin.get("revenue") or 0,
            ebitda=ebitda,
            ev_ebitda_multiple=p.get("ev_ebitda_multiple", 10.0),
            deal_premium=p.get("deal_premium", 0.20),
            synergies_pct=p.get("synergies_pct", 0.05),
            shares=shares, debt=debt, cash=cash, price=price,
        )

    elif model == "replacement_cost":
        result = ve.replacement_cost(
            fixed_assets=fin.get("fixed_assets") or 0,
            rebuild_multiplier=p.get("rebuild_multiplier", 1.2),
            depreciation_adj=p.get("depreciation_adj", 0.3),
            total_assets=tot_assets or 0,
            total_liabilities=tot_liab or 0,
            shares=shares, price=price,
        )

    elif model == "excess_earnings":
        result = ve.excess_earnings(
            net_income=net_inc or 0,
            total_assets=tot_assets or 0,
            book_value=fin.get("book_value_total") or 0,
            fair_return_rate=p.get("fair_return_rate", 0.08),
            discount_rate=p.get("discount_rate", ke),
            shares=shares, price=price,
        )

    elif model == "cfroi":
        result = ve.cfroi_model(
            operating_cf=fin.get("operating_cf") or 0,
            asset_base=tot_assets or 0,
            asset_life=int(p.get("asset_life", 10)),
            required_return=p.get("required_return", wacc),
            shares=shares, debt=debt, cash=cash, price=price,
            invested_capital=ic,
        )

    elif model == "revenue_multiple":
        result = ve.revenue_multiple(
            revenue=fin.get("revenue") or 0,
            ev_revenue_multiple=p.get("ev_revenue_multiple", 3.0),
            shares=shares, debt=debt, cash=cash, price=price,
        )

    elif model == "user_based":
        result = ve.user_based_valuation(
            users=p.get("users", 0),
            revenue_per_user=p.get("revenue_per_user", 0),
            user_growth=p.get("user_growth", 0.15),
            churn_rate=p.get("churn_rate", 0.05),
            discount_rate=p.get("discount_rate", ke),
            years=int(p.get("years", 5)),
            shares=shares, price=price,
        )

    else:
        raise HTTPException(400, f"Unknown model: {model}")

    return _clean({
        "symbol": upper,
        "model": model,
        "current_price": price,
        **result,
    })


@router.get("/{symbol}/dcf-excel")
def download_dcf_excel(symbol: str):
    """
    Generate and download a professional DCF valuation Excel workbook
    for an Indian stock, matching the NVIDIA DCF Model template format.
    5 sheets: Cover, Assumptions, Income Statement, DCF Valuation, Sensitivity.
    """
    merged = app_state.merged_df
    if merged is None:
        raise HTTPException(503, "Universe not loaded")

    upper = symbol.upper()
    row = merged[merged["symbol"] == upper]
    if row.empty:
        row = merged[merged["yf_symbol"] == upper]
    if row.empty:
        raise HTTPException(404, f"Symbol '{symbol}' not found")

    r        = row.iloc[0]
    yf_sym   = r["yf_symbol"]
    base_sym = str(r.get("symbol", upper))
    isin_val = str(r.get("isin", ""))
    company_name = str(r.get("company_name") or r.get("name") or upper)

    fin = _fetch_financials(yf_sym, base_symbol=base_sym, isin=isin_val)

    try:
        xlsx_bytes = generate_dcf_excel(fin, symbol=upper, company_name=company_name)
    except Exception as e:
        logger.exception(f"DCF Excel generation failed for {upper}: {e}")
        raise HTTPException(500, f"Excel generation failed: {e}")

    filename = f"{upper}_DCF_Model.xlsx"
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.post("/{symbol}/model-excel")
def download_model_excel(symbol: str, req: RunValuationRequest):
    """
    Generate and download a professional Excel workbook for ANY valuation model.
    4 sheets: Cover, Inputs & Assumptions, Analysis, Results & Sensitivity.
    Accepts the same body as /run: { model: str, params: dict }.
    """
    merged = app_state.merged_df
    if merged is None:
        raise HTTPException(503, "Universe not loaded")

    upper = symbol.upper()
    row = merged[merged["symbol"] == upper]
    if row.empty:
        row = merged[merged["yf_symbol"] == upper]
    if row.empty:
        raise HTTPException(404, f"Symbol '{symbol}' not found")

    r        = row.iloc[0]
    yf_sym   = r["yf_symbol"]
    base_sym = str(r.get("symbol", upper))
    isin_val = str(r.get("isin", ""))
    price    = float(r.get("last_price") or 0)
    beta     = r.get("beta")
    company_name = str(r.get("company_name") or r.get("name") or upper)

    fin   = _fetch_financials(yf_sym, base_symbol=base_sym, isin=isin_val)
    price = fin.get("price") or price

    # ── Run the valuation model (same logic as /run) ──────────────────────
    p = req.params
    wacc   = p.get("wacc",   _suggest_wacc(beta))
    ke     = p.get("cost_of_equity", _suggest_ke(beta))
    tg     = p.get("terminal_growth", 0.05)
    g1     = p.get("growth_stage1", max(0.05, (fin.get("revenue_growth") or 0.10)))
    g2     = p.get("growth_stage2", max(0.04, g1 * 0.6))
    years  = int(p.get("years", 5))
    y1     = int(p.get("stage1_years", 5))
    y2     = int(p.get("stage2_years", 5))

    fcf    = p.get("fcf",   fin.get("fcf") or 0)
    fcfe   = p.get("fcfe",  fin.get("fcfe") or fcf)
    debt   = p.get("total_debt", fin.get("total_debt") or 0)
    cash   = p.get("cash",       fin.get("cash") or 0)
    shares = p.get("shares", fin.get("shares") or 0)
    eps    = p.get("eps",    fin.get("eps") or 0)
    dps    = p.get("dps",    fin.get("dps") or 0)
    bvps   = p.get("book_value_per_share", fin.get("book_value_per_share") or 0)
    roe    = p.get("roe",    fin.get("roe") or 0.12)
    ebitda = p.get("ebitda", fin.get("ebitda") or 0)
    nopat  = p.get("nopat",  fin.get("nopat") or 0)
    ic     = p.get("invested_capital", fin.get("invested_capital") or 0)
    tot_assets = p.get("total_assets", fin.get("total_assets") or 0)
    tot_liab   = p.get("total_liabilities", fin.get("total_liabilities") or 0)
    earn_g_pct = p.get("earnings_growth_pct", (fin.get("earnings_growth") or 0.10) * 100)
    vol    = p.get("volatility", fin.get("volatility_annual") or 0.30)
    rev_ps = fin.get("revenue_per_share") or (fin.get("revenue", 0) / shares if shares else 0)
    net_inc = fin.get("net_income") or 0

    model = req.model
    result = {}

    if model == "dcf_fcff":
        # For DCF FCFF, redirect to the dedicated 5-sheet generator
        try:
            xlsx_bytes = generate_dcf_excel(fin, symbol=upper, company_name=company_name)
        except Exception as e:
            logger.exception(f"DCF Excel generation failed for {upper}: {e}")
            raise HTTPException(500, f"Excel generation failed: {e}")
        filename = f"{upper}_DCF_Model.xlsx"
        return Response(
            content=xlsx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{filename}"'},
        )

    elif model == "dcf_fcfe":
        result = ve.dcf_fcfe(fcfe, g1, tg, ke, years, shares, price)
    elif model == "dcf_multistage":
        result = ve.dcf_multistage(fcf, g1, g2, tg, wacc, y1, y2, debt, cash, shares, price)
    elif model == "gordon_growth":
        result = ve.gordon_growth(dps, ke, tg, price)
    elif model == "ddm_multistage":
        result = ve.ddm_multistage(dps, g1, tg, ke, years, price)
    elif model == "residual_income":
        bv_growth = p.get("bv_growth", roe * (1 - (p.get("payout_ratio", 0.3))))
        result = ve.residual_income(bvps, roe, ke, bv_growth, years, tg, price)
    elif model == "capitalized_earnings":
        req_ret = p.get("required_return", ke)
        eg_norm = p.get("eps_growth", fin.get("earnings_growth") or 0.05)
        result = ve.capitalized_earnings(eps, req_ret, eg_norm, price)
    elif model == "nav":
        goodwill = p.get("goodwill", 0)
        result = ve.nav_model(tot_assets, tot_liab, shares, price, goodwill)
    elif model == "liquidation":
        assets = fin.get("total_assets") or 0
        liq_split = p.get("asset_split", {})
        result = ve.liquidation_value(
            cash=liq_split.get("cash", cash),
            receivables=liq_split.get("receivables", assets * 0.15),
            inventory=liq_split.get("inventory", assets * 0.10),
            ppe=liq_split.get("ppe", assets * 0.40),
            other_assets=liq_split.get("other", assets * 0.20 - cash),
            total_liabilities=tot_liab,
            shares=shares,
            current_price=price,
            cash_rate=p.get("cash_rate", 1.0),
            receivables_rate=p.get("receivables_rate", 0.85),
            inventory_rate=p.get("inventory_rate", 0.50),
            ppe_rate=p.get("ppe_rate", 0.60),
            other_rate=p.get("other_rate", 0.25),
        )
    elif model == "eva":
        result = ve.eva_model(nopat, ic, wacc, g1, years, tg, debt, cash, shares, price)
    elif model == "peg":
        result = ve.peg_valuation(eps, earn_g_pct, p.get("target_peg", 1.0), price)
    elif model == "trading_comps":
        # Excel generator uses its own built-in peer tables — just pass fin directly
        result = {"intrinsic_value": None, "upside_pct": None}
    elif model == "lbo":
        result = ve.lbo_model(
            ebitda=ebitda, entry_ev_multiple=p.get("entry_multiple", 8.0),
            exit_ev_multiple=p.get("exit_multiple", 10.0),
            debt_ratio=p.get("debt_ratio", 0.60), interest_rate=p.get("interest_rate", 0.09),
            ebitda_growth=p.get("ebitda_growth", fin.get("revenue_growth") or 0.08),
            hold_years=int(p.get("hold_years", 5)),
            shares=shares, cash=cash, current_price=price,
        )
    elif model == "black_scholes":
        result = ve.black_scholes(
            spot=price, strike=p.get("strike", price), volatility=vol,
            risk_free=p.get("risk_free", 0.071), time_years=p.get("time_years", 1.0),
            option_type=p.get("option_type", "call"),
            dividend_yield=fin.get("dividend_yield") or 0.0,
        )
    elif model == "pb_banks":
        result = ve.pb_banks(roe, ke, tg, bvps, price)
    elif model == "cap_rate":
        noi = p.get("noi", fin.get("noi") or ebitda * 0.9 if ebitda else 0)
        result = ve.cap_rate_model(noi, p.get("cap_rate", 0.07), debt, cash, shares, price)
    elif model == "sum_of_parts":
        segments = p.get("segments", [{"name": "Core Business", "ebitda": ebitda, "multiple": 8.0}])
        result = ve.sum_of_parts(segments, cash, debt, shares, price)
    elif model == "vc_method":
        result = ve.vc_method(
            projected_revenue=p.get("projected_revenue", (fin.get("revenue") or 0) * (1.2 ** int(p.get("years", 5)))),
            terminal_revenue_multiple=p.get("terminal_revenue_multiple", 3.0),
            target_return=p.get("target_return", 0.25),
            investment=p.get("investment", price * shares * 0.10 if shares else 0),
            years=int(p.get("years", 5)), shares=shares, current_price=price,
        )
    elif model == "precedent_transactions":
        result = ve.precedent_transactions(
            revenue=fin.get("revenue") or 0, ebitda=ebitda,
            ev_ebitda_multiple=p.get("ev_ebitda_multiple", 10.0),
            deal_premium=p.get("deal_premium", 0.20),
            synergies_pct=p.get("synergies_pct", 0.05),
            shares=shares, debt=debt, cash=cash, price=price,
        )
    elif model == "replacement_cost":
        result = ve.replacement_cost(
            fixed_assets=fin.get("fixed_assets") or 0,
            rebuild_multiplier=p.get("rebuild_multiplier", 1.2),
            depreciation_adj=p.get("depreciation_adj", 0.3),
            total_assets=tot_assets or 0, total_liabilities=tot_liab or 0,
            shares=shares, price=price,
        )
    elif model == "excess_earnings":
        result = ve.excess_earnings(
            net_income=net_inc or 0, total_assets=tot_assets or 0,
            book_value=fin.get("book_value_total") or 0,
            fair_return_rate=p.get("fair_return_rate", 0.08),
            discount_rate=p.get("discount_rate", ke),
            shares=shares, price=price,
        )
    elif model == "cfroi":
        result = ve.cfroi_model(
            operating_cf=fin.get("operating_cf") or 0, asset_base=tot_assets or 0,
            asset_life=int(p.get("asset_life", 10)),
            required_return=p.get("required_return", wacc),
            shares=shares, debt=debt, cash=cash, price=price,
            invested_capital=ic,
        )
    elif model == "revenue_multiple":
        result = ve.revenue_multiple(
            revenue=fin.get("revenue") or 0,
            ev_revenue_multiple=p.get("ev_revenue_multiple", 3.0),
            shares=shares, debt=debt, cash=cash, price=price,
        )
    elif model == "user_based":
        result = ve.user_based_valuation(
            users=p.get("users", 0), revenue_per_user=p.get("revenue_per_user", 0),
            user_growth=p.get("user_growth", 0.15), churn_rate=p.get("churn_rate", 0.05),
            discount_rate=p.get("discount_rate", ke),
            years=int(p.get("years", 5)), shares=shares, price=price,
        )
    else:
        raise HTTPException(400, f"Unknown model: {model}")

    if result.get("error"):
        raise HTTPException(400, result["error"])

    # Add current_price to result for the Excel generator
    result["current_price"] = price

    # ── Generate professional Excel workbook (per-model generator) ────────
    try:
        xlsx_bytes = generate_for_model(
            model_id=model,
            fin=fin,
            symbol=upper,
            company_name=company_name,
        )
    except Exception as e:
        logger.exception(f"Model Excel generation failed for {upper}/{model}: {e}")
        raise HTTPException(500, f"Excel generation failed: {e}")

    model_label = model.replace('_', '-').title()
    filename = f"{upper}_{model_label}_Model.xlsx"
    return Response(
        content=xlsx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
