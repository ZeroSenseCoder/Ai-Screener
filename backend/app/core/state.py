"""
Global in-memory state shared across requests.
Populated at startup and refreshed by background scheduler.

Enrichment flow:
  1. load_universe()        — fast (~60s): NSE + BSE meta + last_price + sector
  2. enrich_indicators()    — background (~10 min): OHLCV + indicators for NSE stocks
     Runs in a thread so API stays responsive while it computes.
"""

import logging
import sqlite3
import threading
import time
from datetime import datetime
from typing import Optional

import pandas as pd

logger = logging.getLogger(__name__)


class AppState:
    universe_df: Optional[pd.DataFrame] = None       # All stocks (meta only)
    indicators_df: Optional[pd.DataFrame] = None     # All computed indicators
    merged_df: Optional[pd.DataFrame] = None         # universe_df + indicators joined
    stocks_by_sector: dict = {}
    last_updated: Optional[str] = None
    enrichment_status: str = "pending"               # pending | running | done | error
    enrichment_progress: int = 0                     # 0-100 %


app_state = AppState()

# ── Indicator columns for DB cache ────────────────────────────────────────────
_IND_DB_COLS = [
    "yf_symbol", "last_price",
    "sma_20", "sma_50", "sma_200", "ema_20", "ema_50", "ema_200",
    "rsi_14", "macd", "macd_signal", "macd_hist", "beta",
    "max_drawdown_52w", "daily_return", "return_5d",
    "return_1m", "return_3m", "return_6m", "return_1y", "avg_volume_20d",
    # Fundamentals — persisted so MC/PE survive server restarts
    "market_cap", "pe_ratio", "forward_pe", "dividend_yield",
    "price_to_book", "debt_to_equity", "revenue_growth", "earnings_growth",
    "profit_margins", "eps",
]

_DB_PATH = "./fintech.db"
# Use a dedicated cache table (avoids conflicts with SQLAlchemy's Indicators model)
_CACHE_TABLE = "indicators_cache"
_CREATE_CACHE_DDL = """
    CREATE TABLE IF NOT EXISTS indicators_cache (
        yf_symbol    TEXT PRIMARY KEY,
        last_price   REAL, sma_20 REAL, sma_50 REAL, sma_200 REAL,
        ema_20       REAL, ema_50 REAL, ema_200 REAL,
        rsi_14       REAL, macd   REAL, macd_signal REAL, macd_hist REAL,
        beta         REAL, max_drawdown_52w REAL, daily_return REAL,
        return_5d    REAL, return_1m REAL, return_3m REAL, return_6m REAL,
        return_1y    REAL, avg_volume_20d REAL,
        market_cap   REAL, pe_ratio REAL, forward_pe REAL, dividend_yield REAL,
        price_to_book REAL, debt_to_equity REAL, revenue_growth REAL,
        earnings_growth REAL, profit_margins REAL, eps REAL
    )
"""


def _save_indicators_to_db(rows: list[dict]) -> None:
    """Upsert a batch of indicator rows into the indicators_cache table."""
    if not rows:
        return
    try:
        ind_df = pd.DataFrame(rows)
        available = [c for c in _IND_DB_COLS if c in ind_df.columns]
        if not available:
            return
        ind_df = ind_df[available]
        conn = sqlite3.connect(_DB_PATH)
        try:
            conn.execute(_CREATE_CACHE_DDL)
            # Migrate: add any missing columns (e.g. return_6m added later)
            existing_cols = {row[1] for row in conn.execute(f"PRAGMA table_info({_CACHE_TABLE})")}
            for col in _IND_DB_COLS:
                if col != "yf_symbol" and col not in existing_cols:
                    try:
                        conn.execute(f"ALTER TABLE {_CACHE_TABLE} ADD COLUMN {col} REAL")
                    except Exception:
                        pass
            conn.commit()
            ind_df.to_sql("_ind_staging", conn, if_exists="replace", index=False)
            cols_sql = ", ".join(available)
            conn.execute(
                f"INSERT OR REPLACE INTO {_CACHE_TABLE} ({cols_sql}) "
                f"SELECT {cols_sql} FROM _ind_staging"
            )
            conn.execute("DROP TABLE IF EXISTS _ind_staging")
            conn.commit()
        finally:
            conn.close()
    except Exception as exc:
        logger.warning(f"_save_indicators_to_db failed: {exc}")


def _load_indicators_from_db() -> Optional[pd.DataFrame]:
    """Load all previously saved indicator rows from indicators_cache."""
    try:
        conn = sqlite3.connect(_DB_PATH)
        try:
            df = pd.read_sql(f"SELECT * FROM {_CACHE_TABLE}", conn)
        finally:
            conn.close()
        if df.empty:
            return None
        return df
    except Exception as exc:
        logger.debug(f"_load_indicators_from_db: {exc}")
        return None


async def load_universe():
    """Fetch NSE+BSE stock list, store in app_state, then kick off background enrichment."""
    from app.services.stock_universe import build_universe
    try:
        logger.info("Loading stock universe...")
        universe_df, by_sector = build_universe(enrich=False)
        app_state.universe_df = universe_df
        app_state.stocks_by_sector = by_sector
        app_state.last_updated = datetime.utcnow().isoformat()

        # Build merged_df with empty indicator columns initially
        merged = universe_df.copy()
        ind_cols = [
            "sma_20", "sma_50", "sma_200", "ema_20", "ema_50", "ema_200",
            "rsi_14", "macd", "macd_signal", "macd_hist", "beta",
            "max_drawdown_52w", "daily_return", "return_5d", "return_1m",
            "return_3m", "return_6m", "return_1y", "avg_volume_20d",
            "pe_ratio", "market_cap", "forward_pe", "dividend_yield",
            "price_to_book", "debt_to_equity", "revenue_growth", "earnings_growth",
            "profit_margins",
        ]
        for col in ind_cols:
            if col not in merged.columns:
                merged[col] = None

        # ── Load cached indicators from DB so technicals are instantly available ──
        cached = _load_indicators_from_db()
        if cached is not None and not cached.empty:
            # Drop placeholder None columns, replace with real DB values
            drop_cols = [c for c in cached.columns if c != "yf_symbol" and c in merged.columns]
            merged = merged.drop(columns=drop_cols, errors="ignore")
            merged = merged.merge(cached, on="yf_symbol", how="left")
            logger.info(f"Loaded {len(cached)} indicator rows from DB cache")

        app_state.merged_df = merged

        logger.info(f"Universe loaded: {len(universe_df)} stocks, {len(by_sector)} sectors")

        # Start background enrichment in a daemon thread
        t = threading.Thread(target=_run_enrichment, daemon=True, name="enrichment")
        t.start()

    except Exception as e:
        logger.error(f"Failed to load universe: {e}", exc_info=True)


def _run_enrichment():
    """
    Background thread: batch-download 1Y OHLCV for NSE stocks, compute indicators,
    then fetch fundamentals (market_cap, pe_ratio, beta) from yfinance fast_info.

    Processes stocks in batches of 50; updates merged_df after each batch so
    the screener shows data progressively.
    """
    import yfinance as yf
    from app.services.indicators import compute_all_indicators

    app_state.enrichment_status = "running"
    logger.info("Background enrichment started")

    try:
        df = app_state.universe_df
        if df is None:
            return

        # Only enrich NSE stocks — BSE-only tickers have near-zero yfinance coverage
        # and processing 4000+ BSE symbols would take 30+ extra minutes for no gain.
        nse = df[df["exchange"] == "NSE"]["yf_symbol"].tolist()
        symbols = nse
        total = len(symbols)
        logger.info(f"Enrichment scope: {total} NSE stocks (BSE-only skipped)")

        # Fetch Nifty 50 as benchmark for beta
        try:
            nifty_raw = yf.Ticker("^NSEI").history(period="1y", interval="1d", auto_adjust=True)
            nifty_df = nifty_raw.reset_index()
            # Flatten MultiIndex if present (newer yfinance may return one for single ticker)
            if isinstance(nifty_df.columns, pd.MultiIndex):
                nifty_df.columns = nifty_df.columns.get_level_values(0)
            nifty_df.columns = [c.lower() for c in nifty_df.columns]
            if "datetime" in nifty_df.columns:
                nifty_df = nifty_df.rename(columns={"datetime": "date"})
        except Exception:
            nifty_df = None

        rows: list[dict] = []
        BATCH = 100

        for i in range(0, total, BATCH):
            batch = symbols[i: i + BATCH]
            batch_str = " ".join(batch)
            try:
                # Batch OHLCV download (much faster than one-by-one)
                raw = yf.download(
                    batch_str,
                    period="1y",
                    interval="1d",
                    auto_adjust=True,
                    progress=False,
                    threads=True,
                    group_by="ticker",
                )
            except Exception as e:
                logger.warning(f"Batch download error (batch {i//BATCH}): {e}")
                time.sleep(2)
                continue

            batch_rows: list[dict] = []
            for sym in batch:
                try:
                    # Extract per-symbol DataFrame — handle both yfinance MultiIndex formats:
                    #   Old (<0.2.50): (Ticker, PriceType) → level 0 = ticker
                    #   New (>=0.2.50): (PriceType, Ticker) → level 1 = ticker
                    if len(batch) == 1:
                        sym_df = raw.copy()
                        # Flatten MultiIndex columns for single-ticker result
                        if isinstance(sym_df.columns, pd.MultiIndex):
                            sym_df.columns = sym_df.columns.get_level_values(0)
                    elif isinstance(raw.columns, pd.MultiIndex):
                        if sym in raw.columns.get_level_values(0):
                            # Old format: ticker at level 0
                            sym_df = raw[sym].copy()
                        elif raw.columns.nlevels > 1 and sym in raw.columns.get_level_values(1):
                            # New format: ticker at level 1 (price type at level 0)
                            sym_df = raw.xs(sym, axis=1, level=1).copy()
                        else:
                            continue
                    else:
                        continue

                    sym_df = sym_df.dropna(how="all").reset_index()
                    sym_df.columns = [
                        (c[0] if isinstance(c, tuple) else c).lower()
                        for c in sym_df.columns
                    ]
                    if "datetime" in sym_df.columns:
                        sym_df = sym_df.rename(columns={"datetime": "date"})
                    sym_df = sym_df.dropna(subset=["close"])

                    if len(sym_df) < 20:
                        continue

                    ind = compute_all_indicators(sym_df, nifty_df)
                    ind["yf_symbol"] = sym
                    # last_price is already set by compute_all_indicators from OHLCV close
                    # market_cap is fetched in the slow fundamentals pass below

                    rows.append(ind)
                    batch_rows.append(ind)

                except Exception as e:
                    logger.debug(f"Indicator compute failed for {sym}: {e}")

            # Update merged_df after each batch
            if rows:
                _update_merged(rows)

            # Persist this batch's indicators to DB for fast restart
            if batch_rows:
                _save_indicators_to_db(batch_rows)

            app_state.enrichment_progress = min(100, int((i + BATCH) / total * 100))
            logger.info(
                f"Enrichment: {min(i + BATCH, total)}/{total} stocks "
                f"({app_state.enrichment_progress}%)"
            )
            time.sleep(0.1)  # brief pause between batches

        # Second pass: fetch pe_ratio, beta, market_cap, dividend_yield from .info for NSE stocks
        _fetch_slow_fundamentals(nse[:500])

        # Third pass: fix sectors for top Unknown stocks (capped to avoid slow runtime)
        _fix_unknown_sectors()

        app_state.enrichment_status = "done"
        app_state.enrichment_progress = 100
        logger.info(f"Background enrichment complete — {len(rows)} stocks enriched")

    except Exception as e:
        app_state.enrichment_status = "error"
        logger.error(f"Background enrichment failed: {e}", exc_info=True)


def _update_merged(rows: list[dict]):
    """Upsert new indicator rows into app_state.merged_df without wiping existing data."""
    try:
        ind_df = pd.DataFrame(rows)
        if "yf_symbol" not in ind_df.columns:
            return

        existing = app_state.merged_df
        if existing is None:
            universe = app_state.universe_df
            if universe is None:
                return
            existing = universe.copy()

        merged = existing.copy()

        # Add any new columns from ind_df that don't exist yet in merged
        for col in ind_df.columns:
            if col != "yf_symbol" and col not in merged.columns:
                merged[col] = None

        # Upsert: update only non-null values for matching yf_symbol rows
        merged = merged.set_index("yf_symbol")
        ind_indexed = ind_df.set_index("yf_symbol")
        for col in ind_indexed.columns:
            if col not in merged.columns:
                merged[col] = None
            new_vals = ind_indexed[col].dropna()
            if not new_vals.empty:
                merged.loc[merged.index.intersection(new_vals.index), col] = new_vals.reindex(
                    merged.index.intersection(new_vals.index)
                )

        app_state.merged_df = merged.reset_index()
    except Exception as e:
        logger.warning(f"_update_merged failed: {e}")


def _fetch_slow_fundamentals(nse_symbols: list[str]):
    """Fetch pe_ratio, market_cap, beta, etc. from yfinance .info.
    Flushes every 50 stocks so MC/PE appear on screener progressively.
    """
    import yfinance as yf
    from app.services.stock_universe import YF_SECTOR_MAP

    logger.info(f"Fetching slow fundamentals for {len(nse_symbols)} NSE stocks...")
    batch_rows: list[dict] = []
    all_rows:   list[dict] = []
    sector_updates: dict[str, str] = {}

    FLUSH_EVERY = 50   # write MC/PE to merged_df + DB every N stocks

    for i, sym in enumerate(nse_symbols):
        try:
            info = yf.Ticker(sym).info
            row = {
                "yf_symbol":      sym,
                "market_cap":     info.get("marketCap"),
                "pe_ratio":       info.get("trailingPE"),
                "forward_pe":     info.get("forwardPE"),
                "beta":           info.get("beta"),
                "dividend_yield": info.get("dividendYield"),
                "price_to_book":  info.get("priceToBook"),
                "debt_to_equity": info.get("debtToEquity"),
                "revenue_growth": info.get("revenueGrowth"),
                "earnings_growth":info.get("earningsGrowth"),
                "profit_margins": info.get("profitMargins"),
                "eps":            info.get("trailingEps"),
            }
            batch_rows.append(row)
            all_rows.append(row)
            yf_sector = info.get("sector", "")
            if yf_sector and yf_sector in YF_SECTOR_MAP:
                sector_updates[sym] = YF_SECTOR_MAP[yf_sector]
        except Exception:
            pass

        # Flush every FLUSH_EVERY stocks — makes MC/PE visible on screener quickly
        if len(batch_rows) >= FLUSH_EVERY:
            _update_merged(batch_rows)
            _save_indicators_to_db(batch_rows)
            logger.info(f"  slow fundamentals: {i+1}/{len(nse_symbols)} (flushed {len(batch_rows)})")
            batch_rows = []
            time.sleep(0.5)
        else:
            time.sleep(0.1)

    # Flush remainder
    if batch_rows:
        _update_merged(batch_rows)
        _save_indicators_to_db(batch_rows)

    # Patch sector for Unknown stocks
    if sector_updates and app_state.merged_df is not None:
        try:
            df = app_state.merged_df.copy()
            for yf_sym, sector in sector_updates.items():
                mask = (df["yf_symbol"] == yf_sym) & (df["sector"] == "Unknown")
                df.loc[mask, "sector"] = sector
            app_state.merged_df = df
            logger.info(f"Sector updated for {len(sector_updates)} stocks via yfinance")
        except Exception as e:
            logger.warning(f"Sector patch failed: {e}")

    logger.info(f"Slow fundamentals complete — {len(all_rows)} stocks")


def _fix_unknown_sectors():
    """
    For remaining Unknown stocks: fetch yfinance .info sector and apply it.
    Processes only stocks still marked Unknown after name/industry matching.
    Caps at 500 stocks to avoid excessive runtime.
    """
    import yfinance as yf
    from app.services.stock_universe import YF_SECTOR_MAP

    df = app_state.merged_df
    if df is None:
        return

    unknown_mask = df["sector"] == "Unknown"
    unknown_syms = df.loc[unknown_mask, "yf_symbol"].dropna().tolist()
    if not unknown_syms:
        return

    logger.info(f"Fixing sectors for {len(unknown_syms)} Unknown stocks via yfinance...")
    # Limit to avoid very long runtime; BSE stocks are usually small-cap
    unknown_syms = unknown_syms[:100]

    updated = 0
    for i, sym in enumerate(unknown_syms):
        try:
            info = yf.Ticker(sym).info
            yf_sector = info.get("sector", "")
            if yf_sector and yf_sector in YF_SECTOR_MAP:
                mapped = YF_SECTOR_MAP[yf_sector]
                curr_df = app_state.merged_df
                if curr_df is not None:
                    mask = (curr_df["yf_symbol"] == sym) & (curr_df["sector"] == "Unknown")
                    if mask.any():
                        curr_df = curr_df.copy()
                        curr_df.loc[mask, "sector"] = mapped
                        app_state.merged_df = curr_df
                        updated += 1
        except Exception:
            pass
        if (i + 1) % 50 == 0:
            logger.info(f"  sector fix: {i+1}/{len(unknown_syms)} ({updated} updated)")
            time.sleep(1)
        else:
            time.sleep(0.15)

    logger.info(f"Sector fix complete: {updated}/{len(unknown_syms)} stocks updated")


def _merge_dataframes():
    """Join universe_df with indicators_df on yf_symbol."""
    if app_state.universe_df is None or app_state.indicators_df is None:
        return
    merged = app_state.universe_df.merge(
        app_state.indicators_df,
        on="yf_symbol",
        how="left",
    )
    app_state.merged_df = merged
