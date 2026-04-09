"""
Valuation Engine — 18 equity valuation models.
All models accept pre-fetched financial inputs and return a dict.
"""
import math
from typing import Optional


# ── helpers ───────────────────────────────────────────────────────────────────

def _safe(v):
    """Return None for NaN/Inf/None."""
    try:
        if v is None or math.isnan(v) or math.isinf(v):
            return None
        return v
    except Exception:
        return None

def _norm_cdf(x: float) -> float:
    """Standard normal CDF via complementary error function."""
    return 0.5 * math.erfc(-x / math.sqrt(2))

def _upside(iv, price):
    if iv and price and price > 0:
        return round((iv - price) / price * 100, 2)
    return None

def _sens_dcf(fcf, growth, terminal_growth_range, wacc_range, years, debt, cash, shares):
    """5×5 sensitivity: rows=WACC, cols=terminal_growth."""
    table = []
    for wacc in wacc_range:
        row = []
        for tg in terminal_growth_range:
            try:
                if wacc <= tg:
                    row.append(None)
                    continue
                pv = 0
                cf = fcf
                for y in range(1, years + 1):
                    cf *= (1 + growth)
                    pv += cf / (1 + wacc) ** y
                tv = cf * (1 + tg) / (wacc - tg)
                pv_tv = tv / (1 + wacc) ** years
                ev = pv + pv_tv
                eq = ev - debt + cash
                val = round(eq / shares, 2) if shares > 0 else None
                row.append(val)
            except Exception:
                row.append(None)
        table.append(row)
    return table


# ── Model 1: DCF – Free Cash Flow to Firm ─────────────────────────────────────

def dcf_fcff(
    fcf: float,
    growth_rate: float,          # stage-1 FCF growth (decimal)
    terminal_growth: float,      # long-run growth (decimal)
    wacc: float,                 # discount rate (decimal)
    years: int,                  # stage-1 explicit years
    total_debt: float,
    cash: float,
    shares: float,
    current_price: float,
) -> dict:
    if wacc <= terminal_growth:
        return {"error": "WACC must exceed terminal growth rate"}
    if shares <= 0:
        return {"error": "Shares outstanding required"}

    pv_stage1 = 0
    year_rows = []
    cf = fcf
    for y in range(1, years + 1):
        cf *= (1 + growth_rate)
        pv = cf / (1 + wacc) ** y
        pv_stage1 += pv
        year_rows.append({"year": y, "fcf": round(cf), "pv_fcf": round(pv)})

    terminal_value = cf * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal = terminal_value / (1 + wacc) ** years
    enterprise_value = pv_stage1 + pv_terminal
    equity_value = enterprise_value - total_debt + cash
    iv = _safe(equity_value / shares)

    wacc_range  = [round(wacc - 0.04 + i * 0.02, 4) for i in range(5)]
    tg_range    = [round(terminal_growth - 0.02 + i * 0.01, 4) for i in range(5)]
    sensitivity = _sens_dcf(fcf, growth_rate, tg_range, wacc_range, years, total_debt, cash, shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "enterprise_value": round(enterprise_value),
        "equity_value": round(equity_value),
        "pv_stage1": round(pv_stage1),
        "pv_terminal": round(pv_terminal),
        "terminal_value": round(terminal_value),
        "year_details": year_rows,
        "sensitivity": {"rows": wacc_range, "cols": tg_range, "values": sensitivity,
                        "row_label": "WACC", "col_label": "Terminal Growth"},
    }


# ── Model 2: DCF – Free Cash Flow to Equity ───────────────────────────────────

def dcf_fcfe(
    fcfe: float,
    growth_rate: float,
    terminal_growth: float,
    cost_of_equity: float,
    years: int,
    shares: float,
    current_price: float,
) -> dict:
    if cost_of_equity <= terminal_growth:
        return {"error": "Cost of equity must exceed terminal growth rate"}
    if shares <= 0:
        return {"error": "Shares outstanding required"}

    pv = 0
    year_rows = []
    cf = fcfe
    for y in range(1, years + 1):
        cf *= (1 + growth_rate)
        pv_y = cf / (1 + cost_of_equity) ** y
        pv += pv_y
        year_rows.append({"year": y, "fcfe": round(cf), "pv_fcfe": round(pv_y)})

    terminal_value = cf * (1 + terminal_growth) / (cost_of_equity - terminal_growth)
    pv_terminal = terminal_value / (1 + cost_of_equity) ** years
    equity_value = pv + pv_terminal
    iv = _safe(equity_value / shares)

    ke_range = [round(cost_of_equity - 0.04 + i * 0.02, 4) for i in range(5)]
    tg_range = [round(terminal_growth - 0.02 + i * 0.01, 4) for i in range(5)]
    sensitivity = _sens_dcf(fcfe, growth_rate, tg_range, ke_range, years, 0, 0, shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "equity_value": round(equity_value),
        "pv_stage1": round(pv),
        "pv_terminal": round(pv_terminal),
        "year_details": year_rows,
        "sensitivity": {"rows": ke_range, "cols": tg_range, "values": sensitivity,
                        "row_label": "Cost of Equity", "col_label": "Terminal Growth"},
    }


# ── Model 3: Multi-stage DCF (3-stage) ────────────────────────────────────────

def dcf_multistage(
    fcf: float,
    growth_stage1: float,    # high growth (years 1-stage1_years)
    growth_stage2: float,    # transition growth (next stage2_years)
    terminal_growth: float,  # perpetuity growth
    wacc: float,
    stage1_years: int,
    stage2_years: int,
    total_debt: float,
    cash: float,
    shares: float,
    current_price: float,
) -> dict:
    if wacc <= terminal_growth:
        return {"error": "WACC must exceed terminal growth rate"}

    pv_total = 0
    year_rows = []
    cf = fcf
    y = 0

    for i in range(stage1_years):
        y += 1
        cf *= (1 + growth_stage1)
        pv_y = cf / (1 + wacc) ** y
        pv_total += pv_y
        year_rows.append({"year": y, "stage": 1, "growth": growth_stage1, "fcf": round(cf), "pv": round(pv_y)})

    for i in range(stage2_years):
        y += 1
        cf *= (1 + growth_stage2)
        pv_y = cf / (1 + wacc) ** y
        pv_total += pv_y
        year_rows.append({"year": y, "stage": 2, "growth": growth_stage2, "fcf": round(cf), "pv": round(pv_y)})

    terminal_value = cf * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal = terminal_value / (1 + wacc) ** y
    pv_total += pv_terminal

    equity_value = pv_total - total_debt + cash
    iv = _safe(equity_value / shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "enterprise_value": round(pv_total),
        "equity_value": round(equity_value),
        "terminal_value": round(terminal_value),
        "pv_terminal": round(pv_terminal),
        "year_details": year_rows,
    }


# ── Model 4: Gordon Growth Model (DDM) ────────────────────────────────────────

def gordon_growth(
    dps: float,             # annual dividends per share
    cost_of_equity: float,  # ke
    growth_rate: float,     # g (perpetual)
    current_price: float,
) -> dict:
    if cost_of_equity <= growth_rate:
        return {"error": "Cost of equity must exceed growth rate"}
    if dps <= 0:
        return {"error": "No dividends — Gordon Growth not applicable"}

    d1 = dps * (1 + growth_rate)
    iv = _safe(d1 / (cost_of_equity - growth_rate))

    ke_range = [round(cost_of_equity - 0.04 + i * 0.02, 4) for i in range(5)]
    g_range  = [round(growth_rate - 0.02 + i * 0.01, 4)   for i in range(5)]
    sens = []
    for ke in ke_range:
        row = []
        for g in g_range:
            if ke <= g or ke <= 0:
                row.append(None)
            else:
                row.append(round(dps * (1 + g) / (ke - g), 2))
        sens.append(row)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "d1": round(d1, 4),
        "sensitivity": {"rows": ke_range, "cols": g_range, "values": sens,
                        "row_label": "Cost of Equity", "col_label": "Growth Rate"},
    }


# ── Model 5: Multi-stage DDM ───────────────────────────────────────────────────

def ddm_multistage(
    dps: float,
    growth_stage1: float,
    terminal_growth: float,
    cost_of_equity: float,
    stage1_years: int,
    current_price: float,
) -> dict:
    if cost_of_equity <= terminal_growth:
        return {"error": "Cost of equity must exceed terminal growth rate"}
    if dps <= 0:
        return {"error": "No dividends — DDM not applicable"}

    pv = 0
    year_rows = []
    d = dps
    for y in range(1, stage1_years + 1):
        d *= (1 + growth_stage1)
        pv_y = d / (1 + cost_of_equity) ** y
        pv += pv_y
        year_rows.append({"year": y, "dps": round(d, 4), "pv": round(pv_y, 4)})

    terminal_value = d * (1 + terminal_growth) / (cost_of_equity - terminal_growth)
    pv_terminal = terminal_value / (1 + cost_of_equity) ** stage1_years
    iv = _safe(pv + pv_terminal)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "pv_stage1_dividends": round(pv, 4),
        "pv_terminal": round(pv_terminal, 4),
        "terminal_value": round(terminal_value, 4),
        "year_details": year_rows,
    }


# ── Model 6: Residual Income Model (RIM) ──────────────────────────────────────

def residual_income(
    book_value_per_share: float,
    roe: float,
    cost_of_equity: float,
    growth_rate: float,    # book value growth per year
    years: int,
    terminal_growth: float,
    current_price: float,
) -> dict:
    if cost_of_equity <= terminal_growth:
        return {"error": "Cost of equity must exceed terminal growth rate"}

    pv = 0
    year_rows = []
    bv = book_value_per_share
    for y in range(1, years + 1):
        eps_y = bv * roe
        charge = bv * cost_of_equity
        ri = eps_y - charge
        pv_y = ri / (1 + cost_of_equity) ** y
        pv += pv_y
        bv_next = bv * (1 + growth_rate)
        year_rows.append({"year": y, "bv": round(bv, 2), "eps": round(eps_y, 2),
                          "equity_charge": round(charge, 2), "ri": round(ri, 2), "pv_ri": round(pv_y, 2)})
        bv = bv_next

    # Terminal RI (continuing residual income)
    ri_t = bv * (roe - cost_of_equity)
    pv_terminal_ri = ri_t / ((cost_of_equity - terminal_growth) * (1 + cost_of_equity) ** years)

    iv = _safe(book_value_per_share + pv + pv_terminal_ri)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "book_value_per_share": book_value_per_share,
        "pv_residual_incomes": round(pv, 2),
        "pv_terminal_ri": round(pv_terminal_ri, 2),
        "year_details": year_rows,
    }


# ── Model 7: Capitalized Earnings ─────────────────────────────────────────────

def capitalized_earnings(
    eps: float,
    required_return: float,
    eps_growth: float,       # normalized EPS growth
    current_price: float,
) -> dict:
    if required_return <= 0:
        return {"error": "Required return must be positive"}
    if eps <= 0:
        return {"error": "Negative EPS — model not applicable"}

    normalized_eps = eps * (1 + eps_growth)
    iv = _safe(normalized_eps / required_return)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "normalized_eps": round(normalized_eps, 2),
        "capitalization_rate": required_return,
        "implied_pe": round(1 / required_return, 2),
    }


# ── Model 8: NAV / Book Value ──────────────────────────────────────────────────

def nav_model(
    total_assets: float,
    total_liabilities: float,
    shares: float,
    current_price: float,
    goodwill: float = 0,       # intangible assets to exclude
    off_balance_items: float = 0,
) -> dict:
    if shares <= 0:
        return {"error": "Shares outstanding required"}

    adjusted_assets = total_assets - goodwill + off_balance_items
    nav = adjusted_assets - total_liabilities
    iv = _safe(nav / shares)

    pb_ratio = current_price / iv if iv and iv > 0 else None

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "total_nav": round(nav),
        "adjusted_assets": round(adjusted_assets),
        "total_liabilities": round(total_liabilities),
        "current_pb": round(pb_ratio, 2) if pb_ratio else None,
    }


# ── Model 9: Liquidation Value ────────────────────────────────────────────────

def liquidation_value(
    cash: float,
    receivables: float,
    inventory: float,
    ppe: float,              # property, plant & equipment
    other_assets: float,
    total_liabilities: float,
    shares: float,
    current_price: float,
    # Recovery rate assumptions
    cash_rate: float = 1.00,
    receivables_rate: float = 0.85,
    inventory_rate: float = 0.50,
    ppe_rate: float = 0.60,
    other_rate: float = 0.25,
) -> dict:
    if shares <= 0:
        return {"error": "Shares outstanding required"}

    recoverable = (
        cash * cash_rate +
        receivables * receivables_rate +
        inventory * inventory_rate +
        ppe * ppe_rate +
        other_assets * other_rate
    )
    liq_nav = recoverable - total_liabilities
    iv = _safe(liq_nav / shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "recoverable_assets": round(recoverable),
        "total_liabilities": round(total_liabilities),
        "liquidation_nav": round(liq_nav),
        "recovery_breakdown": {
            "cash": round(cash * cash_rate),
            "receivables": round(receivables * receivables_rate),
            "inventory": round(inventory * inventory_rate),
            "ppe": round(ppe * ppe_rate),
            "other": round(other_assets * other_rate),
        },
    }


# ── Model 10: Economic Value Added (EVA) ──────────────────────────────────────

def eva_model(
    nopat: float,           # Net operating profit after tax
    invested_capital: float,
    wacc: float,
    growth_rate: float,     # EVA growth rate
    years: int,
    terminal_growth: float,
    total_debt: float,
    cash: float,
    shares: float,
    current_price: float,
) -> dict:
    if wacc <= terminal_growth:
        return {"error": "WACC must exceed terminal growth rate"}

    eva_0 = nopat - wacc * invested_capital
    pv_eva = 0
    year_rows = []
    eva = eva_0
    for y in range(1, years + 1):
        eva *= (1 + growth_rate)
        pv_y = eva / (1 + wacc) ** y
        pv_eva += pv_y
        year_rows.append({"year": y, "eva": round(eva), "pv_eva": round(pv_y)})

    # Terminal MVA
    terminal_mva = eva * (1 + terminal_growth) / (wacc - terminal_growth)
    pv_terminal_mva = terminal_mva / (1 + wacc) ** years

    total_mva = pv_eva + pv_terminal_mva
    enterprise_value = invested_capital + total_mva
    equity_value = enterprise_value - total_debt + cash
    iv = _safe(equity_value / shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "current_eva": round(eva_0),
        "total_mva": round(total_mva),
        "enterprise_value": round(enterprise_value),
        "equity_value": round(equity_value),
        "roic": round(nopat / invested_capital * 100, 2) if invested_capital > 0 else None,
        "wacc_pct": round(wacc * 100, 2),
        "year_details": year_rows,
    }


# ── Model 11: PEG Ratio Valuation ─────────────────────────────────────────────

def peg_valuation(
    eps: float,
    earnings_growth_pct: float,  # in percent e.g. 15 for 15%
    target_peg: float,           # typically 1.0
    current_price: float,
) -> dict:
    if eps <= 0:
        return {"error": "Negative EPS — PEG not applicable"}
    if earnings_growth_pct <= 0:
        return {"error": "Growth must be positive for PEG"}

    fair_pe = target_peg * earnings_growth_pct
    iv = _safe(fair_pe * eps)
    current_pe = current_price / eps if eps > 0 else None
    current_peg = (current_pe / earnings_growth_pct) if current_pe else None

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "fair_pe": round(fair_pe, 2),
        "current_pe": round(current_pe, 2) if current_pe else None,
        "current_peg": round(current_peg, 2) if current_peg else None,
        "target_peg": target_peg,
        "earnings_growth_pct": earnings_growth_pct,
    }


# ── Model 12: Trading Comparables ─────────────────────────────────────────────

def trading_comps(
    stock: dict,      # {pe, ev_ebitda, pb, ps, ev_sales, ebitda, revenue, bv_per_share, eps, shares, debt, cash}
    peers: list,      # list of peer dicts with same fields
    current_price: float,
) -> dict:
    if not peers:
        return {"error": "No sector peers found for comparison"}

    def median(vals):
        clean = sorted(v for v in vals if v and v > 0)
        if not clean:
            return None
        n = len(clean)
        return clean[n // 2] if n % 2 else (clean[n // 2 - 1] + clean[n // 2]) / 2

    peer_pe     = median([p.get("pe_ratio") for p in peers])
    peer_pb     = median([p.get("price_to_book") for p in peers])
    peer_ps     = median([p.get("price_to_sales") for p in peers])
    peer_ev_ebt = median([p.get("ev_ebitda") for p in peers])

    # Implied values per share
    ivs = {}
    if peer_pe and stock.get("eps") and stock["eps"] > 0:
        ivs["pe"] = round(peer_pe * stock["eps"], 2)
    if peer_pb and stock.get("book_value_per_share") and stock["book_value_per_share"] > 0:
        ivs["pb"] = round(peer_pb * stock["book_value_per_share"], 2)
    if peer_ps and stock.get("revenue_per_share") and stock["revenue_per_share"] > 0:
        ivs["ps"] = round(peer_ps * stock["revenue_per_share"], 2)

    all_ivs = [v for v in ivs.values() if v]
    avg_iv = round(sum(all_ivs) / len(all_ivs), 2) if all_ivs else None

    return {
        "intrinsic_value": avg_iv,
        "upside_pct": _upside(avg_iv, current_price),
        "peer_count": len(peers),
        "peer_medians": {
            "pe": round(peer_pe, 2) if peer_pe else None,
            "pb": round(peer_pb, 2) if peer_pb else None,
            "ps": round(peer_ps, 2) if peer_ps else None,
            "ev_ebitda": round(peer_ev_ebt, 2) if peer_ev_ebt else None,
        },
        "implied_values": ivs,
        "stock_multiples": {
            "pe": round(current_price / stock["eps"], 2) if stock.get("eps") and stock["eps"] > 0 else None,
            "pb": round(current_price / stock["book_value_per_share"], 2) if stock.get("book_value_per_share") and stock["book_value_per_share"] > 0 else None,
        },
    }


# ── Model 13: LBO (Simplified) ────────────────────────────────────────────────

def lbo_model(
    ebitda: float,
    entry_ev_multiple: float,   # EV/EBITDA entry
    exit_ev_multiple: float,    # EV/EBITDA exit
    debt_ratio: float,          # debt as % of entry EV
    interest_rate: float,       # annual interest on debt
    ebitda_growth: float,       # annual EBITDA growth
    hold_years: int,
    shares: float,
    cash: float,
    current_price: float,
) -> dict:
    if shares <= 0:
        return {"error": "Shares outstanding required"}

    entry_ev = ebitda * entry_ev_multiple
    entry_debt = entry_ev * debt_ratio
    entry_equity = entry_ev * (1 - debt_ratio)

    # Project EBITDA
    ebitda_exit = ebitda * (1 + ebitda_growth) ** hold_years
    exit_ev = ebitda_exit * exit_ev_multiple

    # Debt paydown (simple: use EBITDA minus interest to pay debt)
    remaining_debt = entry_debt
    for _ in range(hold_years):
        interest = remaining_debt * interest_rate
        ebitda_cur = ebitda * (1 + ebitda_growth) ** (_ + 1)
        free_cf = ebitda_cur * 0.5 - interest   # assume 50% EBITDA conversion to FCF
        remaining_debt = max(0, remaining_debt - free_cf)

    exit_equity = exit_ev - remaining_debt + cash
    iv = _safe(exit_equity / shares)

    # IRR approximation: (exit_equity / entry_equity)^(1/years) - 1
    irr = _safe((exit_equity / entry_equity) ** (1 / hold_years) - 1) if entry_equity > 0 else None
    moic = _safe(exit_equity / entry_equity) if entry_equity > 0 else None

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "entry_ev": round(entry_ev),
        "exit_ev": round(exit_ev),
        "entry_equity": round(entry_equity),
        "exit_equity": round(exit_equity),
        "remaining_debt": round(remaining_debt),
        "irr_pct": round(irr * 100, 1) if irr else None,
        "moic": round(moic, 2) if moic else None,
    }


# ── Model 14: Black-Scholes (Real Options / Warrant Pricing) ──────────────────

def black_scholes(
    spot: float,              # current stock price (S)
    strike: float,            # exercise price (K)
    volatility: float,        # annualized vol (σ)
    risk_free: float,         # risk-free rate (r)
    time_years: float,        # time to expiry in years (T)
    option_type: str = "call",
    dividend_yield: float = 0.0,  # continuous dividend yield (q) — Merton model
) -> dict:
    """Black-Scholes-Merton European option pricing with continuous dividend yield.
    For non-dividend stocks pass dividend_yield=0 (reduces to plain BSM).
    """
    if time_years <= 0 or volatility <= 0:
        return {"error": "Invalid time or volatility"}

    q = dividend_yield or 0.0
    sqrt_T = math.sqrt(time_years)
    disc_q = math.exp(-q * time_years)          # e^(-qT)
    disc_r = math.exp(-risk_free * time_years)  # e^(-rT)
    npdf   = lambda x: math.exp(-0.5 * x * x) / math.sqrt(2 * math.pi)

    # Merton d1 includes (r − q + σ²/2)
    d1 = (math.log(spot / strike) + (risk_free - q + 0.5 * volatility ** 2) * time_years) / (volatility * sqrt_T)
    d2 = d1 - volatility * sqrt_T

    Nd1  = _norm_cdf(d1);  Nd2  = _norm_cdf(d2)
    Nnd1 = _norm_cdf(-d1); Nnd2 = _norm_cdf(-d2)
    nd1  = npdf(d1)

    if option_type == "call":
        price = spot * disc_q * Nd1 - strike * disc_r * Nd2
    else:
        price = strike * disc_r * Nnd2 - spot * disc_q * Nnd1

    # Greeks (Merton-adjusted)
    delta = disc_q * Nd1 if option_type == "call" else disc_q * (Nd1 - 1)
    gamma = disc_q * nd1 / (spot * volatility * sqrt_T)
    # Theta (per calendar day)
    if option_type == "call":
        theta = (
            -spot * disc_q * nd1 * volatility / (2 * sqrt_T)
            - risk_free * strike * disc_r * Nd2
            + q * spot * disc_q * Nd1
        ) / 365
    else:
        theta = (
            -spot * disc_q * nd1 * volatility / (2 * sqrt_T)
            + risk_free * strike * disc_r * Nnd2
            - q * spot * disc_q * Nnd1
        ) / 365
    vega     = spot * disc_q * nd1 * sqrt_T / 100
    rho_call = strike * time_years * disc_r * (Nd2 if option_type == "call" else -Nnd2) / 100

    return {
        "option_price":    round(price, 4),
        "intrinsic_value": round(price, 4),
        "upside_pct":      None,
        "d1":              round(d1, 4),
        "d2":              round(d2, 4),
        "delta":           round(delta, 4),
        "gamma":           round(gamma, 6),
        "theta_daily":     round(theta, 4),
        "vega":            round(vega, 4),
        "rho":             round(rho_call, 4),
        "n_d1":            round(Nd1, 4),
        "n_d2":            round(Nd2, 4),
    }


# ── Model 15: Price-to-Book for Banks ─────────────────────────────────────────

def pb_banks(
    roe: float,
    cost_of_equity: float,
    growth_rate: float,
    book_value_per_share: float,
    current_price: float,
) -> dict:
    if cost_of_equity <= growth_rate:
        return {"error": "Cost of equity must exceed growth rate"}
    if book_value_per_share <= 0:
        return {"error": "Book value required"}

    # Gordon-Growth P/B: P/B = (ROE - g) / (ke - g)
    fair_pb = (roe - growth_rate) / (cost_of_equity - growth_rate)
    iv = _safe(fair_pb * book_value_per_share)
    current_pb = current_price / book_value_per_share if book_value_per_share > 0 else None

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "fair_pb": round(fair_pb, 2),
        "current_pb": round(current_pb, 2) if current_pb else None,
        "book_value_per_share": round(book_value_per_share, 2),
        "roe_pct": round(roe * 100, 2),
        "cost_of_equity_pct": round(cost_of_equity * 100, 2),
    }


# ── Model 16: Cap Rate Valuation (Real Estate) ────────────────────────────────

def cap_rate_model(
    noi: float,              # Net Operating Income
    cap_rate: float,         # market cap rate (e.g. 0.07 for 7%)
    total_debt: float,
    cash: float,
    shares: float,
    current_price: float,
) -> dict:
    if cap_rate <= 0:
        return {"error": "Cap rate must be positive"}
    if shares <= 0:
        return {"error": "Shares outstanding required"}

    property_value = noi / cap_rate
    equity_value = property_value - total_debt + cash
    iv = _safe(equity_value / shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "property_value": round(property_value),
        "equity_value": round(equity_value),
        "implied_cap_rate": round(noi / (current_price * shares + total_debt - cash) * 100, 2) if shares > 0 else None,
    }


# ── Model 17: Sum of the Parts ─────────────────────────────────────────────────

def sum_of_parts(
    segments: list,   # [{"name": str, "ebitda": float, "multiple": float, "debt": float}]
    cash: float,
    total_debt: float,
    shares: float,
    current_price: float,
) -> dict:
    if shares <= 0:
        return {"error": "Shares outstanding required"}
    if not segments:
        return {"error": "At least one business segment required"}

    segment_values = []
    total_segment_ev = 0
    for seg in segments:
        seg_ev = seg.get("ebitda", 0) * seg.get("multiple", 0)
        total_segment_ev += seg_ev
        segment_values.append({
            "name": seg["name"],
            "ebitda": seg.get("ebitda", 0),
            "multiple": seg.get("multiple", 0),
            "value": round(seg_ev),
        })

    equity_value = total_segment_ev - total_debt + cash
    iv = _safe(equity_value / shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "total_segment_ev": round(total_segment_ev),
        "equity_value": round(equity_value),
        "segment_values": segment_values,
    }


# ── Model 18: VC Method ────────────────────────────────────────────────────────

def vc_method(
    projected_revenue: float,   # terminal year revenue
    terminal_revenue_multiple: float,
    target_return: float,       # e.g. 0.25 for 25% IRR
    investment: float,          # capital invested
    years: int,
    shares: float,
    current_price: float,
) -> dict:
    terminal_value = projected_revenue * terminal_revenue_multiple
    post_money_value = terminal_value / (1 + target_return) ** years
    pre_money_value = post_money_value - investment
    ownership_required = investment / post_money_value if post_money_value > 0 else None
    iv = _safe(pre_money_value / shares)

    return {
        "intrinsic_value": iv,
        "upside_pct": _upside(iv, current_price),
        "terminal_value": round(terminal_value),
        "post_money_value": round(post_money_value),
        "pre_money_value": round(pre_money_value),
        "ownership_required_pct": round(ownership_required * 100, 2) if ownership_required else None,
    }


# ── Model 19: Precedent Transactions ──────────────────────────────────────────

def precedent_transactions(revenue, ebitda, ev_ebitda_multiple, deal_premium, synergies_pct, shares, debt, cash, price):
    """
    Precedent transaction analysis — acquisition-premium approach.

    Uses market price as base to avoid standalone vs consolidated EBITDA mismatch
    (screener.in filing data is standalone; yfinance EV is consolidated — applying
    a deal multiple to standalone EBITDA produces grossly understated values).

    Formula:
        Acquisition Value (equity) = price × shares × (1 + deal_premium)
        Synergy EV                 = ebitda × synergies_pct × ev_ebitda_multiple
        Total Offer Equity         = Acquisition Value + Synergy EV
        Implied Price              = Total Offer Equity / Shares
    """
    if not price or price <= 0:
        return {"error": "Current price required"}
    if not shares or shares <= 0:
        return {"error": "Shares outstanding required"}

    mktcap       = price * shares
    acq_equity   = mktcap * (1.0 + deal_premium)
    synergy_ev   = (ebitda or 0) * synergies_pct * ev_ebitda_multiple
    total_equity = acq_equity + synergy_ev
    iv           = total_equity / shares

    return {
        "intrinsic_value":  iv,
        "upside_pct":       (iv - price) / price * 100 if price else None,
        "acquisition_value": acq_equity,
        "synergy_ev":        synergy_ev,
        "total_offer_equity": total_equity,
        "deal_premium_applied": deal_premium,
    }


# ── Model 20: Replacement Cost ─────────────────────────────────────────────────

def replacement_cost(fixed_assets, rebuild_multiplier, depreciation_adj, total_assets, total_liabilities, shares, price):
    """Replacement cost valuation - cost to rebuild the asset base."""
    fa = fixed_assets or (total_assets * 0.4 if total_assets else 0)
    replacement_gross = fa * rebuild_multiplier
    depreciation_deduction = replacement_gross * depreciation_adj
    net_replacement = replacement_gross - depreciation_deduction
    net_other = (total_assets - fa) if total_assets else 0
    total_val = net_replacement + net_other
    equity_value = total_val - (total_liabilities or 0)
    if not shares or shares <= 0:
        return {"error": "Shares outstanding required"}
    iv = equity_value / shares
    return {
        "intrinsic_value": iv,
        "upside_pct": (iv - price) / price * 100 if price else None,
        "replacement_gross": replacement_gross,
        "depreciation_deduction": depreciation_deduction,
        "net_replacement_value": net_replacement,
        "equity_value": equity_value,
    }


# ── Model 21: Excess Earnings ──────────────────────────────────────────────────

def excess_earnings(net_income, total_assets, book_value, fair_return_rate, discount_rate, shares, price):
    """Excess Earnings Model: value = tangible assets + PV of excess earnings."""
    fair_return = (total_assets or 0) * fair_return_rate
    excess = (net_income or 0) - fair_return
    pv_excess = excess / discount_rate if discount_rate > 0 else 0
    total_value = (book_value or total_assets or 0) + pv_excess
    if not shares or shares <= 0:
        return {"error": "Shares outstanding required"}
    iv = total_value / shares
    return {
        "intrinsic_value": iv,
        "upside_pct": (iv - price) / price * 100 if price else None,
        "fair_return_on_assets": fair_return,
        "excess_earnings_annual": excess,
        "pv_of_excess_earnings": pv_excess,
        "total_equity_value": total_value,
    }


# ── Model 22: CFROI ────────────────────────────────────────────────────────────

def cfroi_model(operating_cf, asset_base, asset_life, required_return, shares, debt, cash, price,
                invested_capital=0):
    """CFROI = Operating Cash Flow / Capital Employed.
    Capital Employed = Invested Capital (Equity + Financial Debt) when provided,
    otherwise falls back to Total Assets (asset_base).
    EV = Capital Employed + PV of excess economic returns over asset life.
    """
    capital_employed = (invested_capital or 0) if (invested_capital or 0) > 0 else (asset_base or 0)
    if not capital_employed or capital_employed <= 0:
        return {"error": "Capital employed / asset base required"}
    cfroi_rate = (operating_cf or 0) / capital_employed
    spread = cfroi_rate - required_return
    annuity = (1 - (1 + required_return) ** (-asset_life)) / required_return if required_return > 0 else asset_life
    value_from_spread = capital_employed * spread * annuity
    enterprise_value = capital_employed + value_from_spread
    equity_value = enterprise_value - (debt or 0) + (cash or 0)
    if not shares or shares <= 0:
        return {"error": "Shares outstanding required"}
    iv = equity_value / shares
    return {
        "intrinsic_value": iv,
        "upside_pct": (iv - price) / price * 100 if price else None,
        "cfroi_rate": round(cfroi_rate * 100, 2),
        "capital_employed": round(capital_employed),
        "required_return_pct": round(required_return * 100, 2),
        "spread_pct": round(spread * 100, 2),
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
    }


# ── Model 23: Revenue Multiple ─────────────────────────────────────────────────

def revenue_multiple(revenue, ev_revenue_multiple, shares, debt, cash, price):
    """
    EV/Revenue Multiple valuation.

    EV/Revenue is an ENTERPRISE VALUE multiple (not P/S).
    Equity bridge: Equity = EV − Net Debt.

    Formula:
        Implied EV    = Revenue × EV/Revenue_Multiple
        Net Debt      = Total Debt − Cash
        Equity Value  = Implied EV − Net Debt
        Implied Price = Equity Value / Shares
    """
    if not revenue:
        return {"error": "Revenue required"}
    if not shares or shares <= 0:
        return {"error": "Shares outstanding required"}

    net_debt     = (debt or 0) - (cash or 0)
    ev           = revenue * ev_revenue_multiple
    equity_value = ev - net_debt
    iv           = equity_value / shares

    return {
        "intrinsic_value":    iv,
        "upside_pct":         (iv - price) / price * 100 if price else None,
        "enterprise_value":   ev,
        "net_debt":           net_debt,
        "equity_value":       equity_value,
        "multiple_used":      ev_revenue_multiple,
    }


# ── Model 24: User-Based Valuation ────────────────────────────────────────────

def user_based_valuation(users, revenue_per_user, user_growth, churn_rate, discount_rate, years, shares, price):
    """User-based valuation for platforms/SaaS companies."""
    if not users or users <= 0:
        return {"error": "User count required (from annual report)"}
    current_users = users
    total_pv = 0
    details = []
    for y in range(1, int(years) + 1):
        current_users = current_users * (1 + user_growth - churn_rate)
        rev = current_users * revenue_per_user
        pv = rev / (1 + discount_rate) ** y
        total_pv += pv
        details.append({"year": y, "users": int(current_users), "revenue": round(rev, 0), "pv_revenue": round(pv, 0)})
    terminal_rev = current_users * revenue_per_user
    tg = min(0.03, discount_rate - 0.01)
    tv = terminal_rev * (1 + tg) / max(0.001, discount_rate - tg)
    pv_tv = tv / (1 + discount_rate) ** years
    total_value = total_pv + pv_tv
    if not shares or shares <= 0:
        return {"error": "Shares outstanding required"}
    iv = total_value / shares
    return {
        "intrinsic_value": iv,
        "upside_pct": (iv - price) / price * 100 if price else None,
        "total_pv_revenues": total_pv,
        "terminal_value": tv,
        "pv_terminal": pv_tv,
        "total_equity_value": total_value,
        "year_details": details,
    }
