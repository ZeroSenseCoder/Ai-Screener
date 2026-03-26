# FinScope — Indian Stock Screener & Valuation Platform

A full-stack platform for screening, analysing, and valuing Indian equities (NSE + BSE).
Built with **FastAPI + Python** (backend) and **React + TypeScript + Vite** (frontend).

---

## Table of Contents

- [Features](#features)
- [Architecture Overview](#architecture-overview)
- [Project Structure](#project-structure)
- [Tech Stack](#tech-stack)
- [Getting Started](#getting-started)
- [API Reference](#api-reference)
- [Valuation Models](#valuation-models)
- [Data Sources](#data-sources)
- [Workflow](#workflow)

---

## Features

| Category | Details |
|---|---|
| **Universe** | ~6,000+ NSE + BSE stocks with sector classification |
| **Basic Screener** | Filter by price, market cap, P/E, beta, RSI, MACD, moving averages, returns, volume, drawdown |
| **Pro Screener** | Bloomberg-style weighted multi-condition scoring across 50+ filters |
| **Valuation Engine** | 24 models — DCF, DDM, LBO, NAV, VC, PEG, EVA, CFROI, Black-Scholes, and more |
| **Auto-filled Financials** | Balance sheet, P&L, and cash flow data pre-filled from BSE/NSE XBRL filings via screener.in |
| **Excel Export** | Download valuation results as a multi-sheet Excel workbook |
| **Stock Detail** | Live quote, OHLCV candlestick chart (8 timeframes), technical indicators, fundamentals, news |
| **Macro Dashboard** | Global indices, forex, commodities, India VIX, FII/DII flows |
| **Options Data** | F&O open interest |
| **News** | Aggregated from Yahoo Finance, RSS feeds, and optional NewsAPI |

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                          BROWSER                                │
│                                                                 │
│   React SPA  (Vite + TypeScript + TailwindCSS)                  │
│   ┌─────────┐ ┌────────────┐ ┌───────────┐ ┌───────────────┐  │
│   │Screener │ │ProScreener │ │StockDetail│ │  Valuation    │  │
│   │  Page   │ │   Page     │ │   Page    │ │    Page       │  │
│   └────┬────┘ └─────┬──────┘ └─────┬─────┘ └───────┬───────┘  │
│        └────────────┴──────────────┴───────────────┘           │
│                          Axios (api.ts)                         │
└──────────────────────────┬──────────────────────────────────────┘
                           │ HTTP  /api/v1/*
                           │ (Vite proxy → localhost:8000)
┌──────────────────────────▼──────────────────────────────────────┐
│                     FastAPI Backend                             │
│                                                                 │
│  ┌─────────────────────── API Layer ──────────────────────────┐ │
│  │  /universe  /screener  /pro-screener  /stocks  /macro      │ │
│  │  /valuation                                                │ │
│  └──────────────────────────┬─────────────────────────────────┘ │
│                             │                                   │
│  ┌──────────────── Core State ────────────────────────────────┐ │
│  │  universe_df  │  indicators_df  │  merged_df               │ │
│  │  (in-memory Pandas DataFrames — all queries at RAM speed)  │ │
│  └───────────┬─────────────────────────────────────┬──────────┘ │
│              │ background thread                   │ startup    │
│  ┌───────────▼──────────────┐        ┌─────────────▼──────────┐ │
│  │  Enrichment Pipeline     │        │   SQLite DB Cache      │ │
│  │  (OHLCV → Indicators →   │        │   indicators_cache     │ │
│  │   Fundamentals)          │        │   (survives restarts)  │ │
│  └───────────┬──────────────┘        └────────────────────────┘ │
│              │                                                   │
│  ┌───────────▼──────────────────────────────────────────────┐   │
│  │                     Services Layer                        │   │
│  │  stock_universe  │  screener_engine  │  pro_screener_engine│  │
│  │  valuation_engine│  indicators       │  market_data        │  │
│  │  bse_filings     │  macro_service    │  news_service       │  │
│  └───────────┬──────────────────────────────────────────────┘   │
└──────────────┼──────────────────────────────────────────────────┘
               │
┌──────────────▼──────────────────────────────────────────────────┐
│                    External Data Sources                        │
│                                                                 │
│  NSE API         BSE Open Data       Yahoo Finance (yfinance)   │
│  (stock list)    (stock list)        (OHLCV, quotes, fundamentals│
│                                       options chain, news)      │
│                                                                 │
│  screener.in     RSS Feeds           NewsAPI (optional)         │
│  (BSE/NSE XBRL   (news headlines)    (news headlines)           │
│   filings)                                                      │
└─────────────────────────────────────────────────────────────────┘
```

---

## Project Structure

```
FINTECH/
├── backend/
│   ├── app/
│   │   ├── api/v1/
│   │   │   ├── universe.py         # Stock universe endpoints
│   │   │   ├── screener.py         # Basic screener endpoints
│   │   │   ├── pro_screener.py     # Pro screener endpoints
│   │   │   ├── stocks.py           # Per-stock detail endpoints
│   │   │   ├── macro.py            # Macro data endpoints
│   │   │   └── valuation.py        # Valuation endpoints
│   │   ├── services/
│   │   │   ├── stock_universe.py   # NSE/BSE stock list builder
│   │   │   ├── screener_engine.py  # Filter logic (pandas)
│   │   │   ├── pro_screener_engine.py  # Weighted scoring engine
│   │   │   ├── valuation_engine.py # 24 valuation model functions
│   │   │   ├── indicators.py       # Technical indicator computation
│   │   │   ├── market_data.py      # OHLCV + fundamentals via yfinance
│   │   │   ├── bse_filings.py      # XBRL filing parser (screener.in)
│   │   │   ├── macro_service.py    # Macro data aggregation
│   │   │   ├── news_service.py     # News aggregation
│   │   │   └── oi_service.py       # Options open interest
│   │   ├── core/
│   │   │   └── state.py            # Global app state + enrichment pipeline
│   │   ├── models/
│   │   │   └── stock.py            # SQLAlchemy ORM models
│   │   ├── db/
│   │   │   └── database.py         # Async SQLite engine
│   │   ├── config.py               # App settings
│   │   └── main.py                 # FastAPI app + routers
│   └── requirements.txt
│
├── frontend/
│   ├── src/
│   │   ├── pages/
│   │   │   ├── ScreenerPage.tsx    # Basic screener UI
│   │   │   ├── ProScreenerPage.tsx # Pro screener UI
│   │   │   ├── StockDetailPage.tsx # Stock detail + chart
│   │   │   └── ValuationPage.tsx   # 24-model valuation calculator
│   │   ├── components/
│   │   │   ├── common/Navbar.tsx
│   │   │   ├── screener/           # Filter panel, sector bar, stock table
│   │   │   └── macro/MacroDashboard.tsx
│   │   ├── services/
│   │   │   └── api.ts              # Typed Axios API client
│   │   ├── utils/
│   │   │   └── formatters.ts
│   │   ├── App.tsx                 # Router setup
│   │   └── main.tsx
│   ├── vite.config.ts              # Dev server + /api proxy
│   ├── tailwind.config.js
│   └── package.json
│
└── start.bat                       # One-click start (Windows)
```

---

## Tech Stack

### Backend
| Package | Purpose |
|---|---|
| FastAPI | Async REST API framework |
| Uvicorn | ASGI server |
| SQLAlchemy + aiosqlite | ORM + async SQLite |
| pandas + numpy | In-memory data processing |
| yfinance | Market data (OHLCV, fundamentals, options) |
| requests + BeautifulSoup4 | Web scraping (screener.in filings) |
| feedparser | RSS news feeds |
| httpx | Async HTTP client |
| pydantic-settings | Configuration management |

### Frontend
| Package | Purpose |
|---|---|
| React 18 + TypeScript | UI library |
| Vite | Build tool + dev server |
| React Router v6 | Client-side routing |
| TailwindCSS | Utility-first styling |
| Axios | HTTP client |
| TanStack Table + Virtual | Virtualized data tables |
| Recharts + Lightweight Charts | Charts + candlesticks |
| xlsx (SheetJS) | Excel export |

---

## Getting Started

### Prerequisites
- Python 3.10+
- Node.js 18+
- Windows (or adapt paths for Linux/macOS)

### Quick Start (Windows)
```bat
double-click start.bat
```
This automatically creates the Python venv, installs all dependencies, and starts both servers.

### Manual Start

**Terminal 1 — Backend:**
```bash
cd backend
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

**Terminal 2 — Frontend:**
```bash
cd frontend
npm install
npm run dev
```

**Access:**
- Frontend: http://localhost:5173
- Backend API: http://localhost:8000
- Swagger UI: http://localhost:8000/docs

### Optional Configuration
Create `backend/.env`:
```env
NEWS_API_KEY=your_newsapi_key   # optional — for enhanced news
DATABASE_URL=sqlite+aiosqlite:///./finscope.db
```

---

## API Reference

### Universe
| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/v1/universe/stats` | Total stocks, NSE/BSE counts, sector count |
| GET | `/api/v1/universe/sectors/list` | Sector names + stock counts |
| GET | `/api/v1/universe/sectors/{name}` | Stocks in a specific sector |
| GET | `/api/v1/universe/stocks` | All stocks (supports `exchange`, `sector`, `search`, `page` params) |

### Screener
| Method | Endpoint | Description |
|---|---|---|
| POST | `/api/v1/screener/screen` | Apply filters + sorting + pagination |
| GET | `/api/v1/screener/sector-summary` | Sector aggregates for filter sidebar |
| GET | `/api/v1/screener/meta` | Available filter options |

### Pro Screener
| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/v1/pro-screener/catalog` | All 50+ available filters grouped by category |
| POST | `/api/v1/pro-screener/screen` | Weighted multi-condition scoring |
| GET | `/api/v1/pro-screener/sectors` | Sectors with stock counts |

### Stock Detail
| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/v1/stocks/search?q=` | Search by symbol or name |
| GET | `/api/v1/stocks/{symbol}/ohlcv?timeframe=1D` | Candlestick data |
| GET | `/api/v1/stocks/{symbol}/indicators` | Technical indicators |
| GET | `/api/v1/stocks/{symbol}/fundamentals` | P/E, P/B, dividend yield, etc. |
| GET | `/api/v1/stocks/{symbol}/news?days=7` | News articles |
| GET | `/api/v1/stocks/{symbol}/summary` | All-in-one: quote + indicators + fundamentals |
| GET | `/api/v1/stocks/{symbol}/oi` | Options open interest |

### Macro
| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/v1/macro/overview` | All macro data in one payload |
| GET | `/api/v1/macro/indices` | Global indices |
| GET | `/api/v1/macro/forex` | INR/USD, INR/EUR, etc. |
| GET | `/api/v1/macro/commodities` | Gold, oil, natural gas |
| GET | `/api/v1/macro/fii-dii` | FII/DII flows |

### Valuation
| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/v1/valuation/{symbol}/inputs` | Pre-filled financials from BSE/NSE filings + suggested params |
| POST | `/api/v1/valuation/{symbol}/run` | Run a valuation model; returns intrinsic value + upside % |

---

## Valuation Models

24 models across 9 categories, all with financials **pre-filled from BSE/NSE XBRL filings**:

| Category | Models |
|---|---|
| **DCF** | FCFF DCF, FCFE DCF, Multistage DCF |
| **Dividend** | Gordon Growth Model, Multistage DDM |
| **Equity Value** | Residual Income, Capitalized Earnings |
| **Asset-Based** | NAV, Liquidation Value, Replacement Cost |
| **Market Multiples** | Trading Comps (EV/EBITDA), Revenue Multiple |
| **Special Situations** | Precedent Transactions (M&A), LBO, VC Method |
| **Advanced** | EVA (Economic Value Added), CFROI, Excess Earnings, Black-Scholes (Real Options) |
| **Sector-Specific** | Cap Rate / NOI (Real Estate), P/B Ratio (Banks), PEG Ratio (Growth), Sum-of-Parts |
| **Alternative** | User-Based Valuation (SaaS/Tech) |

Each model UI has two sections:
- **Balance Sheet / Financials** — auto-filled from latest annual filing (editable override)
- **Assumptions & Inputs** — required fields marked `*`, optional fields with defaults

Results can be exported as a multi-sheet Excel workbook (Summary + Year Details + Sensitivity).

---

## Data Sources

| Source | Data |
|---|---|
| **NSE API** | NSE stock universe (1,978+ equities) |
| **BSE Open Data** | BSE stock universe (4,838+ equities) |
| **Yahoo Finance (yfinance)** | OHLCV history, live quotes, fundamentals, options chain, dividends, news |
| **screener.in** | BSE/NSE mandatory XBRL filings — P&L, Balance Sheet, Cash Flow (latest annual) |
| **Yahoo Finance RSS** | Stock news headlines |
| **NewsAPI** *(optional)* | Enhanced news coverage |

---

## Workflow

### Startup Sequence

```
Server start
    │
    ▼
1. Database init (SQLAlchemy, create tables if new)
    │
    ▼
2. Load stock universe  (~30–60 sec)
   ├── Fetch NSE stock list
   ├── Fetch BSE stock list
   ├── Map symbols to yfinance tickers
   └── Assign sectors from NSE indices
    │
    ▼
3. Load cached indicators from SQLite DB
   └── merged_df now has indicators for ~1,947+ stocks instantly
    │
    ▼
4. API server ready  (screener works immediately with cached data)
    │
    ▼
5. Background enrichment  (~10 min, non-blocking)
   ├── Pass 1: Download 1Y OHLCV for all NSE stocks (batches of 100)
   │          Compute SMA, EMA, RSI, MACD, beta, drawdown, returns
   │          Flush every 50 stocks to DB + update merged_df live
   │
   ├── Pass 2: Fetch slow fundamentals (P/E, P/B, market cap, dividend yield)
   │          from yfinance — flush every 50 stocks to DB
   │
   └── Pass 3: Attempt to fix "Unknown" sector stocks
              Progress: 0% → 100% reported via /api/v1/universe/stats
```

### Screener Request Flow

```
User applies filters in browser
    │
    ▼
POST /api/v1/screener/screen  (or /pro-screener/screen)
    │
    ▼
screener_engine.apply_filters(merged_df, filters)
    │  (pure pandas in-memory — no DB query)
    ▼
Returns paginated results (JSON)
    │
    ▼
StockTable renders with TanStack Virtual (smooth even for 1000+ rows)
```

### Valuation Request Flow

```
User selects stock → navigates to Valuation page
    │
    ▼
GET /api/v1/valuation/{symbol}/inputs
    │
    ├── bse_filings.fetch_bse_filings(symbol)
    │   ├── Search screener.in for slug
    │   ├── Fetch company page HTML
    │   ├── Parse profit-loss, balance-sheet, cash-flow, ratios tables
    │   └── Return values in absolute ₹
    │
    ├── yfinance (for live price, shares, dividends, volatility, growth)
    │
    └── Merge: filing data takes priority for financials
              yfinance fills market data
    │
    ▼
UI pre-fills all Balance Sheet fields (from filing)
User adjusts Assumptions & required (*) fields
    │
    ▼
POST /api/v1/valuation/{symbol}/run  { model, params }
    │
    ▼
valuation_engine.<model_function>(params)
    │
    ▼
Returns: intrinsic_value, upside_pct, enterprise_value, year_details, ...
    │
    ▼
UI shows result card + optional Excel export
```

---

## Database Schema

### `indicators_cache` (SQLite — fast restart)
| Column | Type | Description |
|---|---|---|
| `yf_symbol` | TEXT PK | Yahoo Finance ticker |
| `last_price` | REAL | Latest closing price |
| `sma_20/50/200` | REAL | Simple moving averages |
| `ema_20/50/200` | REAL | Exponential moving averages |
| `rsi_14` | REAL | RSI (14-period) |
| `macd`, `macd_signal`, `macd_hist` | REAL | MACD line, signal, histogram |
| `beta` | REAL | Beta vs Nifty 50 |
| `max_drawdown_52w` | REAL | 52-week max drawdown |
| `daily_return`, `return_1m`, `return_3m`, `return_1y` | REAL | Return metrics |
| `avg_volume_20d` | INTEGER | 20-day average volume |
| `market_cap`, `pe_ratio`, `forward_pe` | REAL | Fundamental metrics |
| `dividend_yield`, `price_to_book` | REAL | Valuation ratios |
| `debt_to_equity`, `profit_margins` | REAL | Quality metrics |
| `revenue_growth`, `earnings_growth` | REAL | Growth metrics |
| `eps` | REAL | Earnings per share |

---

## Contributing

1. Fork the repo
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit: `git commit -m "Add your feature"`
4. Push: `git push origin feature/your-feature`
5. Open a Pull Request

---

## License

MIT License — see [LICENSE](LICENSE) for details.
