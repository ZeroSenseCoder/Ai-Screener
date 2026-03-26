"""
Compute technical indicators from OHLCV data using pandas.
No external API needed — pure pandas/numpy.
"""

import numpy as np
import pandas as pd


def compute_sma(series: pd.Series, period: int) -> float | None:
    if len(series) < period:
        return None
    return float(series.iloc[-period:].mean())


def compute_ema(series: pd.Series, period: int) -> float | None:
    if len(series) < period:
        return None
    ema = series.ewm(span=period, adjust=False).mean()
    return float(ema.iloc[-1])


def compute_rsi(series: pd.Series, period: int = 14) -> float | None:
    if len(series) < period + 1:
        return None
    delta = series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(com=period - 1, adjust=False).mean()
    avg_loss = loss.ewm(com=period - 1, adjust=False).mean()
    rs = avg_gain / avg_loss.replace(0, np.nan)
    rsi = 100 - (100 / (1 + rs))
    return float(rsi.iloc[-1])


def compute_macd(series: pd.Series) -> dict | None:
    if len(series) < 26:
        return None
    ema12 = series.ewm(span=12, adjust=False).mean()
    ema26 = series.ewm(span=26, adjust=False).mean()
    macd_line = ema12 - ema26
    signal = macd_line.ewm(span=9, adjust=False).mean()
    hist = macd_line - signal
    return {
        "macd": float(macd_line.iloc[-1]),
        "signal": float(signal.iloc[-1]),
        "hist": float(hist.iloc[-1]),
    }


def compute_max_drawdown(series: pd.Series) -> float | None:
    """Max drawdown over the supplied series (trailing 252 trading days = ~1yr)."""
    if len(series) < 2:
        return None
    tail = series.iloc[-252:] if len(series) > 252 else series
    rolling_max = tail.cummax()
    drawdown = (tail - rolling_max) / rolling_max
    return float(drawdown.min()) * 100  # as percentage


def compute_beta(stock_returns: pd.Series, index_returns: pd.Series, period: int = 252) -> float | None:
    """Beta of stock vs benchmark (Nifty 50) over trailing `period` days."""
    stock_r = stock_returns.iloc[-period:]
    index_r = index_returns.iloc[-period:]
    if len(stock_r) < 30:
        return None
    aligned = pd.concat([stock_r, index_r], axis=1).dropna()
    if len(aligned) < 30:
        return None
    cov = aligned.cov()
    variance = aligned.iloc[:, 1].var()
    if variance == 0:
        return None
    return float(cov.iloc[0, 1] / variance)


def compute_all_indicators(ohlcv_df: pd.DataFrame, nifty_df: pd.DataFrame | None = None) -> dict:
    """
    Given a DataFrame with columns [date, open, high, low, close, volume]
    sorted ascending by date, return all computed indicators as a dict.
    """
    if ohlcv_df.empty or len(ohlcv_df) < 2:
        return {}

    close = ohlcv_df["close"].astype(float)
    volume = ohlcv_df["volume"].astype(float)
    daily_returns = close.pct_change().dropna()

    result: dict = {}

    # Price levels
    result["last_price"] = float(close.iloc[-1])
    result["sma_20"] = compute_sma(close, 20)
    result["sma_50"] = compute_sma(close, 50)
    result["sma_200"] = compute_sma(close, 200)
    result["ema_20"] = compute_ema(close, 20)
    result["ema_50"] = compute_ema(close, 50)
    result["ema_200"] = compute_ema(close, 200)

    # Momentum
    result["rsi_14"] = compute_rsi(close)
    macd = compute_macd(close)
    if macd:
        result["macd"] = macd["macd"]
        result["macd_signal"] = macd["signal"]
        result["macd_hist"] = macd["hist"]

    # Risk
    result["max_drawdown_52w"] = compute_max_drawdown(close)

    # Returns
    if len(daily_returns) >= 1:
        result["daily_return"] = float(daily_returns.iloc[-1]) * 100
    if len(close) >= 5:
        result["return_5d"] = float((close.iloc[-1] / close.iloc[-5] - 1) * 100)
    if len(close) >= 20:
        result["return_1m"] = float((close.iloc[-1] / close.iloc[-20] - 1) * 100)
    if len(close) >= 63:
        result["return_3m"] = float((close.iloc[-1] / close.iloc[-63] - 1) * 100)
    if len(close) >= 126:
        result["return_6m"] = float((close.iloc[-1] / close.iloc[-126] - 1) * 100)
    if len(close) >= 245:
        result["return_1y"] = float((close.iloc[-1] / close.iloc[-245] - 1) * 100)

    # Volume
    if len(volume) >= 20:
        result["avg_volume_20d"] = int(volume.iloc[-20:].mean())

    # Beta vs Nifty
    if nifty_df is not None and not nifty_df.empty:
        nifty_returns = nifty_df["close"].astype(float).pct_change().dropna()
        # Align by date index
        stock_r = daily_returns.reset_index(drop=True)
        nifty_r = nifty_returns.reset_index(drop=True)
        min_len = min(len(stock_r), len(nifty_r))
        result["beta"] = compute_beta(stock_r.iloc[-min_len:], nifty_r.iloc[-min_len:])

    return result
