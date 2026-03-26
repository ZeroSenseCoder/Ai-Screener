"""
Fetch financial statement data from screener.in
(which sources directly from NSE/BSE mandatory XBRL filings).

Returns structured P&L, Balance Sheet, and Cash Flow data
from the latest annual filing.
"""

import logging
import math
import re
import time
from typing import Optional

import requests

logger = logging.getLogger(__name__)

_SESSION = requests.Session()
_SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
})

_CACHE: dict[str, tuple[float, dict]] = {}  # symbol → (timestamp, data)
_CACHE_TTL = 3600  # 1 hour


def _cr_to_abs(value_cr: Optional[float]) -> Optional[float]:
    """Convert screener.in crore value to absolute rupees."""
    if value_cr is None:
        return None
    return value_cr * 1e7


def _parse_number(text: str) -> Optional[float]:
    """Parse a number string like '1,23,456.78' or '(45.67)' → float."""
    if not text:
        return None
    text = text.strip().replace(",", "").replace("\u20b9", "").replace("₹", "")
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        v = float(text)
        return -v if negative else v
    except ValueError:
        return None


def _get_csrf(session: requests.Session, url: str) -> Optional[str]:
    """Fetch page and extract CSRF token for subsequent requests."""
    try:
        r = session.get(url, timeout=15)
        match = re.search(r'csrfmiddlewaretoken["\s]+value["\s=]+([A-Za-z0-9]+)', r.text)
        if match:
            return match.group(1)
    except Exception:
        pass
    return None


def _search_screener(symbol: str) -> Optional[str]:
    """
    Search screener.in for a stock symbol.
    Returns the screener.in slug (e.g. 'HDFCBANK') or None.
    """
    try:
        url = f"https://www.screener.in/api/company/search/?q={symbol}&fields=name,url"
        r = _SESSION.get(url, timeout=10)
        r.raise_for_status()
        data = r.json()
        if isinstance(data, list) and len(data) > 0:
            # url looks like '/company/HDFCBANK/consolidated/'
            url_field = data[0].get("url", "")
            parts = [p for p in url_field.strip("/").split("/") if p]
            # parts = ['company', 'HDFCBANK', 'consolidated'] → index 1
            if len(parts) >= 2:
                return parts[1]
            elif parts:
                return parts[0]
    except Exception as e:
        logger.debug(f"screener search failed for {symbol}: {e}")
    return None


def _fetch_screener_page(slug: str, consolidated: bool = True) -> Optional[str]:
    """Fetch the screener.in company page HTML."""
    ctype = "consolidated" if consolidated else "standalone"
    url = f"https://www.screener.in/company/{slug}/{ctype}/"
    try:
        r = _SESSION.get(url, timeout=20)
        if r.status_code == 404 and consolidated:
            # Try standalone if consolidated not available
            r = _SESSION.get(
                f"https://www.screener.in/company/{slug}/", timeout=20
            )
        if r.ok:
            return r.text
    except Exception as e:
        logger.debug(f"screener fetch failed for {slug}: {e}")
    return None


def _extract_table(html: str, section_id: str) -> dict[str, list[float | None]]:
    """
    Extract a financial table from screener.in HTML by section id.
    Returns {row_label: [latest_value, prior_1, prior_2, ...]}
    """
    result: dict[str, list] = {}
    try:
        # Find the section
        pat = re.compile(
            rf'id="{section_id}".*?<tbody>(.*?)</tbody>',
            re.DOTALL | re.IGNORECASE,
        )
        m = pat.search(html)
        if not m:
            return result

        tbody = m.group(1)
        rows = re.findall(r"<tr[^>]*>(.*?)</tr>", tbody, re.DOTALL)

        for row in rows:
            cells = re.findall(r"<td[^>]*>(.*?)</td>", row, re.DOTALL)
            if not cells:
                continue
            # First cell = label, rest = values (most recent first)
            label_raw = re.sub(r"<[^>]+>", "", cells[0])
            # Clean HTML entities and screener-specific characters
            label_raw = label_raw.replace("&nbsp;", " ").replace("+", "").replace("%", "pct")
            label = re.sub(r"\s+", " ", label_raw).strip().lower()
            values = []
            for c in cells[1:]:
                clean = re.sub(r"<[^>]+>", "", c).replace("&nbsp;", "").replace(",", "").strip()
                values.append(_parse_number(clean))
            if label:
                result[label] = values

    except Exception as e:
        logger.debug(f"_extract_table({section_id}): {e}")
    return result


def _first_valid(*values) -> Optional[float]:
    for v in values:
        if v is not None and not (isinstance(v, float) and math.isnan(v)):
            return v
    return None


def _latest(table: dict, *keys) -> Optional[float]:
    """Get the most recent (index 0) value for any matching key."""
    for key in keys:
        for k, vals in table.items():
            if key in k and vals:
                v = vals[0]
                if v is not None:
                    return v
    return None


def fetch_bse_filings(symbol: str, isin: str = "") -> dict:
    """
    Fetch latest annual financial statement data for a stock from screener.in.
    Returns dict with all financial metrics in absolute ₹ (not crores).
    """
    # Cache check
    cached = _CACHE.get(symbol)
    if cached and (time.time() - cached[0]) < _CACHE_TTL:
        return cached[1]

    result: dict = {}

    # 1. Find the screener slug
    slug = _search_screener(symbol)
    if not slug:
        logger.debug(f"screener: no slug found for {symbol}")
        return result

    # 2. Fetch the company page
    html = _fetch_screener_page(slug, consolidated=True)
    if not html:
        return result

    # 3. Extract shares outstanding from key metrics section
    shares = None
    shares_match = re.search(
        r"(?:Shares outstanding|No\.? of Shares)[^\d]*([\d,\.]+)\s*(?:Cr|crore)?",
        html, re.IGNORECASE
    )
    if shares_match:
        v = _parse_number(shares_match.group(1))
        if v:
            shares = v * 1e7  # screener shows Cr shares

    # 4. P&L table  (section id = "profit-loss")
    pl = _extract_table(html, "profit-loss")

    # screener.in uses: "revenue", "sales", "net sales", "financing profit" (banks)
    revenue_cr   = _latest(pl,
        "revenue", "sales", "net sales", "revenue from operations",
        "total income", "total revenue", "income from operations")
    # EBITDA proxies — screener.in shows "operating profit" or "financing profit" (banks)
    ebitda_cr    = _latest(pl,
        "operating profit", "ebitda", "pbdit", "financing profit",
        "profit before depreciation")
    pat_cr       = _latest(pl, "net profit", "profit after tax", "pat")
    eps_val      = _latest(pl, "eps in rs", "eps", "diluted eps", "basic eps")
    dep_cr       = _latest(pl, "depreciation", "amortisation", "d&a")
    interest_cr  = _latest(pl, "interest", "finance cost", "finance charges")
    tax_cr       = _latest(pl, "tax", "provision for tax", "income tax")

    # 5. Balance Sheet table (section id = "balance-sheet")
    bs = _extract_table(html, "balance-sheet")

    # Screener shows "equity capital" + "reserves" separately (no combined equity row)
    eq_capital_cr = _latest(bs, "equity capital")
    reserves_cr   = _latest(bs, "reserves")
    equity_cr     = None
    if eq_capital_cr is not None and reserves_cr is not None:
        equity_cr = eq_capital_cr + reserves_cr
    else:
        equity_cr = _latest(bs,
            "shareholders equity", "total equity", "networth", "net worth",
            "total shareholder", "equity")

    # Debt: screener uses "borrowing" for banks, "borrowings" for others
    debt_cr = _latest(bs,
        "borrowing", "borrowings", "total debt",
        "long term borrowing", "total borrowings")
    if debt_cr is None:
        lt = _latest(bs, "long term borrowing") or 0
        st = _latest(bs, "short term borrowing") or 0
        debt_cr = lt + st if (lt or st) else None

    # Cash: screener.in often doesn't show cash separately for banks
    cash_cr = _latest(bs,
        "cash", "cash and bank", "cash equivalents",
        "cash & cash equivalents", "cash and cash equivalents")

    tot_assets_cr = _latest(bs, "total assets", "total liabilities", "balance sheet size")
    fixed_cr      = _latest(bs, "fixed assets", "net block", "tangible assets", "property")

    # Book value per share from BS equity / shares
    bvps = None
    if equity_cr and shares and shares > 0:
        bvps = round((equity_cr * 1e7) / shares, 2)

    # 6. Cash Flow table (section id = "cash-flow")
    cf = _extract_table(html, "cash-flow")

    op_cf_cr  = _latest(cf,
        "cash from operating", "operating activities",
        "net cash from operating", "cash flow from operations",
        "net cash generated from operating")
    capex_cr  = _latest(cf,
        "capital expenditure", "capex", "purchase of fixed",
        "acquisition of fixed", "purchase of property")
    if capex_cr is not None and capex_cr > 0:
        capex_cr = -capex_cr  # normalize to negative

    # FCF = Operating CF + Capex
    fcf_cr = None
    if op_cf_cr is not None and capex_cr is not None:
        fcf_cr = op_cf_cr + capex_cr
    elif op_cf_cr is not None:
        fcf_cr = op_cf_cr * 0.7  # rough: capex ≈ 30% of op CF

    # 7. Key ratios
    ratio_table = _extract_table(html, "ratios")
    roe_val  = _latest(ratio_table, "roe", "return on equity")
    roce_val = _latest(ratio_table, "roce", "return on capital")

    # Also try to get shares from the "Shares outstanding" field in ratios
    if shares is None:
        sh_cr = _latest(ratio_table, "number of shares", "no. of shares", "shares outstanding")
        if sh_cr:
            shares = sh_cr * 1e7  # in crores on screener.in

    # Convert everything from crores to absolute rupees
    def to_abs(cr_val):
        return round(cr_val * 1e7, 0) if cr_val is not None else None

    revenue       = to_abs(revenue_cr)
    ebitda        = to_abs(ebitda_cr)
    pat           = to_abs(pat_cr)
    depreciation  = to_abs(dep_cr)
    interest      = to_abs(interest_cr)
    equity_total  = to_abs(equity_cr)
    total_debt    = to_abs(debt_cr)
    cash          = to_abs(cash_cr)
    total_assets  = to_abs(tot_assets_cr)
    operating_cf  = to_abs(op_cf_cr)
    capex_abs     = to_abs(abs(capex_cr) if capex_cr else None)
    fcf           = to_abs(fcf_cr)
    fixed_assets  = to_abs(fixed_cr)

    # Derived
    total_liab    = (total_assets - equity_total) if (total_assets and equity_total) else None
    invested_cap  = (equity_total + (total_debt or 0)) if equity_total else None
    ebit          = (ebitda - (depreciation or 0)) if ebitda else pat  # fallback
    nopat         = ebit * 0.75 if ebit else pat  # rough 25% tax
    noi           = ebitda * 0.9 if ebitda else None  # for real estate
    roe_dec       = (roe_val / 100) if roe_val else (pat / equity_total if (pat and equity_total and equity_total > 0) else None)

    result = {
        # Income statement
        "revenue":            revenue,
        "ebitda":             ebitda,
        "net_income":         pat,
        "depreciation":       depreciation,
        "interest":           interest,
        "nopat":              nopat,
        # Balance sheet
        "total_assets":       total_assets,
        "total_debt":         total_debt,
        "cash":               cash,
        "total_equity":       equity_total,
        "total_liabilities":  total_liab,
        "book_value_total":   equity_total,
        "book_value_per_share": bvps,
        "invested_capital":   invested_cap,
        "fixed_assets":       fixed_assets,
        # Cash flow
        "operating_cf":       operating_cf,
        "capex":              capex_abs,
        "fcf":                fcf,
        "fcfe":               fcf,  # simplified
        # Per share
        "eps":                eps_val,
        "shares":             shares,
        # Ratios
        "roe":                roe_dec,
        "roce":               (roce_val / 100) if roce_val else None,
        "noi":                noi,
        # Metadata
        "_source":            "screener.in (BSE/NSE filings)",
        "_slug":              slug,
    }

    # Remove None values for cleanliness but keep structure
    result = {k: v for k, v in result.items()}

    _CACHE[symbol] = (time.time(), result)
    logger.info(f"bse_filings: fetched {symbol} from screener.in ({slug})")
    return result
