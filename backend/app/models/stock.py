from sqlalchemy import Column, String, Float, BigInteger, DateTime, Boolean, Text
from sqlalchemy.sql import func
from app.db.database import Base


class StockMeta(Base):
    __tablename__ = "stock_meta"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    symbol = Column(String(50), nullable=False, index=True)
    yf_symbol = Column(String(60), unique=True, nullable=False, index=True)
    company_name = Column(String(200), nullable=False)
    isin = Column(String(12), index=True)
    exchange = Column(String(10), nullable=False)  # NSE | BSE
    sector = Column(String(100), index=True, default="Unknown")
    industry = Column(String(150))
    market_cap = Column(Float)
    pe_ratio = Column(Float)
    beta = Column(Float)
    dividend_yield = Column(Float)
    face_value = Column(Float)
    is_active = Column(Boolean, default=True)
    last_updated = Column(DateTime, server_default=func.now(), onupdate=func.now())


class DailyOHLCV(Base):
    __tablename__ = "daily_ohlcv"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    yf_symbol = Column(String(60), nullable=False, index=True)
    date = Column(String(10), nullable=False, index=True)  # YYYY-MM-DD
    open = Column(Float)
    high = Column(Float)
    low = Column(Float)
    close = Column(Float)
    volume = Column(BigInteger)
    daily_return = Column(Float)  # % change


class Indicators(Base):
    __tablename__ = "indicators"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    yf_symbol = Column(String(60), unique=True, nullable=False, index=True)
    sma_20 = Column(Float)
    sma_50 = Column(Float)
    sma_200 = Column(Float)
    ema_20 = Column(Float)
    ema_50 = Column(Float)
    ema_200 = Column(Float)
    rsi_14 = Column(Float)
    macd = Column(Float)
    macd_signal = Column(Float)
    macd_hist = Column(Float)
    beta = Column(Float)
    max_drawdown_52w = Column(Float)
    daily_return = Column(Float)
    return_5d = Column(Float)
    return_1m = Column(Float)
    return_3m = Column(Float)
    return_1y = Column(Float)
    avg_volume_20d = Column(BigInteger)
    last_price = Column(Float)
    last_updated = Column(DateTime, server_default=func.now(), onupdate=func.now())


class NewsCache(Base):
    __tablename__ = "news_cache"

    id = Column(BigInteger, primary_key=True, autoincrement=True)
    symbol = Column(String(50), nullable=False, index=True)
    title = Column(Text)
    description = Column(Text)
    url = Column(Text)
    source = Column(String(100))
    published_at = Column(String(30))
    importance = Column(String(10))  # high | medium | low
    fetched_at = Column(DateTime, server_default=func.now())
