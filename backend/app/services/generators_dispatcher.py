"""
Dispatcher — routes model_id to its dedicated Excel generator file.
Each model has its own generator with a unique color scheme and analysis.
"""

import logging
logger = logging.getLogger(__name__)

# ── Import every individual generator ────────────────────────────────────────
from app.services.dcf_excel_generator        import generate_dcf_excel          as _dcf_fcff

def _load(name):
    try:
        mod = __import__(f"app.services.{name}", fromlist=["generate_excel"])
        return mod.generate_excel
    except Exception as e:
        logger.warning(f"Generator {name} not available: {e}")
        return None

_REGISTRY = {
    "dcf_fcff":              lambda fin, sym, co: _dcf_fcff(fin, sym, co),
    "dcf_fcfe":              _load("dcf_fcfe_generator"),
    "dcf_multistage":        _load("dcf_multistage_generator"),
    "gordon_growth":         _load("gordon_growth_generator"),
    "ddm_multistage":        _load("ddm_multistage_generator"),
    "residual_income":       _load("residual_income_generator"),
    "trading_comps":         _load("trading_comps_generator"),
    "precedent_transactions":_load("precedent_transactions_generator"),
    "peg":                   _load("peg_generator"),
    "revenue_multiple":      _load("revenue_multiple_generator"),
    "nav":                   _load("nav_generator"),
    "liquidation":           _load("liquidation_generator"),
    "replacement_cost":      _load("replacement_cost_generator"),
    "capitalized_earnings":  _load("capitalized_earnings_generator"),
    "excess_earnings":       _load("excess_earnings_generator"),
    "eva":                   _load("eva_generator"),
    "cfroi":                 _load("cfroi_generator"),
    "lbo":                   _load("lbo_generator"),
    "black_scholes":         _load("black_scholes_generator"),
    "real_options":          _load("real_options_generator"),
    "sum_of_parts":          _load("sum_of_parts_generator"),
    "pb_banks":              _load("pb_banks_generator"),
    "cap_rate":              _load("cap_rate_generator"),
    "user_based":            _load("user_based_generator"),
    "vc_method":             _load("vc_method_generator"),
}


def generate_for_model(model_id: str, fin: dict, symbol: str, company_name: str) -> bytes:
    """
    Route to the correct per-model generator.
    Falls back to the generic model_excel_generator if the dedicated one isn't ready.
    """
    gen = _REGISTRY.get(model_id)

    if gen is None:
        # Fallback: generic generator
        logger.warning(f"No dedicated generator for '{model_id}', using generic fallback")
        from app.services.model_excel_generator import generate_model_excel
        return generate_model_excel(
            model_id=model_id,
            fin=fin,
            result={"intrinsic_value": None, "upside_pct": None,
                    "current_price": fin.get("price", 0)},
            params={},
            symbol=symbol,
            company_name=company_name,
        )

    return gen(fin, symbol, company_name)
