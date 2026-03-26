"""
Fetches all NSE + BSE stocks and organises them by sector.

Working sources (all free, no API key needed):
  NSE  → nseindia.com/api (requires session cookie — handled below)
  BSE  → api.bseindia.com/BseIndiaAPI/api/ListofScripData (open JSON)
  Sector → NSE meta field + curated override map
"""

import logging
import time
import urllib.parse
from typing import Optional

import pandas as pd
import requests
import yfinance as yf

logger = logging.getLogger(__name__)

NSE_SESSION_URL = "https://www.nseindia.com"
NSE_TOTAL_MARKET_URL = "https://www.nseindia.com/api/equity-stockIndices?index=NIFTY%20TOTAL%20MARKET"
NSE_ALL_PREOPEN_URL = "https://www.nseindia.com/api/market-data-pre-open?key=ALL"
NSE_INDEX_URL = "https://www.nseindia.com/api/equity-stockIndices?index={}"
NSE_QUOTE_URL = "https://www.nseindia.com/api/quote-equity?symbol={}"

# NSE sector indices → maps every constituent symbol directly to a sector
NSE_SECTOR_INDICES: list[tuple[str, str]] = [
    ("NIFTY BANK",                       "Financial Services"),
    ("NIFTY FINANCIAL SERVICES",         "Financial Services"),
    ("NIFTY FINANCIAL SERVICES 25/50",   "Financial Services"),
    ("NIFTY IT",                         "Information Technology"),
    ("NIFTY PHARMA",                     "Healthcare"),
    ("NIFTY AUTO",                       "Automobile"),
    ("NIFTY REALTY",                     "Real Estate"),
    ("NIFTY ENERGY",                     "Energy"),
    ("NIFTY FMCG",                       "Consumer Staples"),
    ("NIFTY METAL",                      "Metals & Mining"),
    ("NIFTY MEDIA",                      "Consumer Discretionary"),
    ("NIFTY INFRA",                      "Infrastructure"),
    ("NIFTY COMMODITIES",                "Chemicals"),
    ("NIFTY CONSUMER DURABLES",          "Consumer Discretionary"),
    ("NIFTY HEALTHCARE INDEX",           "Healthcare"),
    ("NIFTY OIL AND GAS",                "Energy"),
    ("NIFTY SERVICES SECTOR",            "Financial Services"),
    ("NIFTY INDIA DIGITAL",              "Information Technology"),
    ("NIFTY INDIA CONSUMPTION",          "Consumer Staples"),
    ("NIFTY INDIA MANUFACTURING",        "Capital Goods"),
    ("NIFTY CAPITAL MARKETS",            "Financial Services"),
    ("NIFTY TRANSPORTATION & LOGISTICS", "Infrastructure"),
    ("NIFTY INDIA DEFENCE",              "Capital Goods"),
    ("NIFTY PSU BANK",                   "Financial Services"),
    ("NIFTY PRIVATE BANK",               "Financial Services"),
    ("NIFTY CPSE",                       None),   # mixed — use industry fallback
    ("NIFTY50 VALUE 20",                 None),
    ("NIFTY NEXT 50",                    None),
    ("NIFTY MIDCAP 50",                  None),
    ("NIFTY MIDCAP 100",                 None),
    ("NIFTY MIDCAP 150",                 None),
    ("NIFTY SMALLCAP 50",                None),
    ("NIFTY SMALLCAP 100",               None),
    ("NIFTY SMALLCAP 250",               None),
    ("NIFTY LARGEMIDCAP 250",            None),
    ("NIFTY500 MULTICAP 50:25:25",       None),
]

BSE_LIST_URL = (
    "https://api.bseindia.com/BseIndiaAPI/api/ListofScripData/w"
    "?Group=&Scripcode=&industry=&segment=Equity&status=Active"
)

NSE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/market-data/live-equity-market",
}

BSE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer": "https://www.bseindia.com/",
}

# ── yfinance sector → our sector ──────────────────────────────────────────────
YF_SECTOR_MAP: dict[str, str] = {
    "Technology":              "Information Technology",
    "Healthcare":              "Healthcare",
    "Financial Services":      "Financial Services",
    "Consumer Cyclical":       "Consumer Discretionary",
    "Consumer Defensive":      "Consumer Staples",
    "Basic Materials":         "Chemicals",
    "Industrials":             "Capital Goods",
    "Energy":                  "Energy",
    "Utilities":               "Energy",
    "Real Estate":             "Real Estate",
    "Communication Services":  "Telecom",
}

# ── Industry string → Sector  (case-insensitive substring match) ──────────────
# Ordered: more specific patterns first so they win on substring match
INDUSTRY_TO_SECTOR: dict[str, str] = {
    # ── Financial Services ──
    "private sector bank":              "Financial Services",
    "public sector bank":               "Financial Services",
    "bank":                             "Financial Services",
    "banking":                          "Financial Services",
    "housing finance":                  "Financial Services",
    "home finance":                     "Financial Services",
    "microfinance":                     "Financial Services",
    "nbfc-mfi":                         "Financial Services",
    "nbfc":                             "Financial Services",
    "non-banking finance":              "Financial Services",
    "non banking finance":              "Financial Services",
    "finance":                          "Financial Services",
    "financial services":               "Financial Services",
    "insurance":                        "Financial Services",
    "life insurance":                   "Financial Services",
    "general insurance":                "Financial Services",
    "reinsurance":                      "Financial Services",
    "stock exchange":                   "Financial Services",
    "stockbroking":                     "Financial Services",
    "brokerage":                        "Financial Services",
    "depository":                       "Financial Services",
    "investment company":               "Financial Services",
    "mutual fund":                      "Financial Services",
    "asset management":                 "Financial Services",
    "wealth management":                "Financial Services",
    "payment":                          "Financial Services",
    "credit card":                      "Financial Services",
    "capital market":                   "Financial Services",
    "holding company":                  "Financial Services",
    "investment":                       "Financial Services",
    "fintech":                          "Financial Services",

    # ── Information Technology ──
    "it consulting & software":         "Information Technology",
    "it services & consulting":         "Information Technology",
    "computers - software":             "Information Technology",
    "software & consulting":            "Information Technology",
    "software services":                "Information Technology",
    "information technology":           "Information Technology",
    "it-software":                      "Information Technology",
    "software":                         "Information Technology",
    "it ":                              "Information Technology",
    " it":                              "Information Technology",
    "^it$":                             "Information Technology",
    "electronics - software":           "Information Technology",
    "internet services":                "Information Technology",
    "internet":                         "Information Technology",
    "ecommerce":                        "Information Technology",
    "e-commerce":                       "Information Technology",
    "data processing":                  "Information Technology",
    "bpo":                              "Information Technology",
    "ites":                             "Information Technology",

    # ── Healthcare / Pharma ──
    "pharmaceuticals & biotechnology":  "Healthcare",
    "pharmaceutical":                   "Healthcare",
    "pharma":                           "Healthcare",
    "biotechnology":                    "Healthcare",
    "biotech":                          "Healthcare",
    "healthcare services":              "Healthcare",
    "healthcare":                       "Healthcare",
    "hospital":                         "Healthcare",
    "diagnostic":                       "Healthcare",
    "medical device":                   "Healthcare",
    "medical equipment":                "Healthcare",
    "surgical":                         "Healthcare",
    "laboratory":                       "Healthcare",
    "wellness":                         "Healthcare",
    "ayurveda":                         "Healthcare",
    "veterinary":                       "Healthcare",

    # ── Consumer Staples ──
    "fmcg":                             "Consumer Staples",
    "consumer food":                    "Consumer Staples",
    "packaged foods":                   "Consumer Staples",
    "food products":                    "Consumer Staples",
    "food processing":                  "Consumer Staples",
    "food & beverages":                 "Consumer Staples",
    "beverages":                        "Consumer Staples",
    "personal care":                    "Consumer Staples",
    "household goods":                  "Consumer Staples",
    "household products":               "Consumer Staples",
    "tobacco":                          "Consumer Staples",
    "edible oil":                       "Consumer Staples",
    "dairy":                            "Consumer Staples",
    "sugar":                            "Consumer Staples",
    "breweries":                        "Consumer Staples",
    "distilleries":                     "Consumer Staples",
    "liquor":                           "Consumer Staples",
    "alcoholic beverage":               "Consumer Staples",
    "tea":                              "Consumer Staples",
    "coffee":                           "Consumer Staples",
    "rice":                             "Consumer Staples",
    "flour":                            "Consumer Staples",
    "wheat":                            "Consumer Staples",
    "agri":                             "Consumer Staples",
    "agricultural":                     "Consumer Staples",
    "seeds":                            "Consumer Staples",
    "aquaculture":                      "Consumer Staples",
    "fisheries":                        "Consumer Staples",

    # ── Consumer Discretionary ──
    "consumer durables":                "Consumer Discretionary",
    "consumer electronics":             "Consumer Discretionary",
    "retail":                           "Consumer Discretionary",
    "retailing":                        "Consumer Discretionary",
    "jewellery":                        "Consumer Discretionary",
    "gems":                             "Consumer Discretionary",
    "watches":                          "Consumer Discretionary",
    "apparels":                         "Consumer Discretionary",
    "apparel":                          "Consumer Discretionary",
    "garment":                          "Consumer Discretionary",
    "fashion":                          "Consumer Discretionary",
    "footwear":                         "Consumer Discretionary",
    "leather":                          "Consumer Discretionary",
    "hotels":                           "Consumer Discretionary",
    "hospitality":                      "Consumer Discretionary",
    "restaurant":                       "Consumer Discretionary",
    "travel":                           "Consumer Discretionary",
    "tourism":                          "Consumer Discretionary",
    "airlines":                         "Consumer Discretionary",
    "aviation":                         "Consumer Discretionary",
    "media":                            "Consumer Discretionary",
    "entertainment":                    "Consumer Discretionary",
    "film":                             "Consumer Discretionary",
    "publishing":                       "Consumer Discretionary",
    "newspaper":                        "Consumer Discretionary",
    "broadcasting":                     "Consumer Discretionary",
    "recreation":                       "Consumer Discretionary",
    "sport":                            "Consumer Discretionary",
    "gaming":                           "Consumer Discretionary",
    "education":                        "Consumer Discretionary",
    "toy":                              "Consumer Discretionary",
    "stationery":                       "Consumer Discretionary",
    "furniture":                        "Consumer Discretionary",
    "home furnishing":                  "Consumer Discretionary",
    "paints":                           "Consumer Discretionary",
    "decorative":                       "Consumer Discretionary",
    "packaging":                        "Consumer Discretionary",

    # ── Automobile ──
    "automobiles & auto components":    "Automobile",
    "automobile":                       "Automobile",
    "auto ancillar":                    "Automobile",
    "auto components":                  "Automobile",
    "auto parts":                       "Automobile",
    "two & three wheeler":              "Automobile",
    "commercial vehicle":               "Automobile",
    "passenger car":                    "Automobile",
    "utility vehicle":                  "Automobile",
    "tractor":                          "Automobile",
    "tyre":                             "Automobile",
    "tire":                             "Automobile",
    "motor":                            "Automobile",
    "vehicle":                          "Automobile",
    "bus":                              "Automobile",
    "ev":                               "Automobile",
    "electric vehicle":                 "Automobile",

    # ── Energy ──
    "crude oil & natural gas":          "Energy",
    "oil exploration":                  "Energy",
    "oil & gas":                        "Energy",
    "oil and gas":                      "Energy",
    "refineries & marketing":           "Energy",
    "refineries":                       "Energy",
    "refinery":                         "Energy",
    "petroleum":                        "Energy",
    "gas distribution":                 "Energy",
    "gas transmission":                 "Energy",
    "lng":                              "Energy",
    "cng":                              "Energy",
    "power generation":                 "Energy",
    "power transmission":               "Energy",
    "power distribution":               "Energy",
    "power sector":                     "Energy",
    "electric utility":                 "Energy",
    "electricity":                      "Energy",
    "solar energy":                     "Energy",
    "renewable energy":                 "Energy",
    "wind energy":                      "Energy",
    "hydro":                            "Energy",
    "thermal power":                    "Energy",
    "nuclear":                          "Energy",
    "coal":                             "Energy",
    "power":                            "Energy",
    "energy":                           "Energy",

    # ── Metals & Mining ──
    "ferrous metals":                   "Metals & Mining",
    "non-ferrous metals":               "Metals & Mining",
    "non ferrous metals":               "Metals & Mining",
    "steel":                            "Metals & Mining",
    "aluminium":                        "Metals & Mining",
    "aluminum":                         "Metals & Mining",
    "copper":                           "Metals & Mining",
    "zinc":                             "Metals & Mining",
    "lead":                             "Metals & Mining",
    "iron":                             "Metals & Mining",
    "mining & minerals":                "Metals & Mining",
    "mining":                           "Metals & Mining",
    "minerals":                         "Metals & Mining",
    "precious metals":                  "Metals & Mining",
    "gold":                             "Metals & Mining",
    "silver":                           "Metals & Mining",
    "platinum":                         "Metals & Mining",
    "industrial minerals":              "Metals & Mining",
    "metal":                            "Metals & Mining",

    # ── Chemicals ──
    "chemicals & petrochemicals":       "Chemicals",
    "specialty chemicals":              "Chemicals",
    "specialty chemical":               "Chemicals",
    "speciality chemical":              "Chemicals",
    "chemical":                         "Chemicals",
    "petrochemical":                    "Chemicals",
    "fertilizer":                       "Chemicals",
    "fertiliser":                       "Chemicals",
    "agrochemical":                     "Chemicals",
    "pesticide":                        "Chemicals",
    "herbicide":                        "Chemicals",
    "insecticide":                      "Chemicals",
    "dye":                              "Chemicals",
    "paint raw material":               "Chemicals",
    "industrial gases":                 "Chemicals",
    "adhesive":                         "Chemicals",
    "resin":                            "Chemicals",
    "polymer":                          "Chemicals",
    "plastic":                          "Chemicals",
    "rubber":                           "Chemicals",
    "lubricant":                        "Chemicals",
    "ink":                              "Chemicals",
    "coatings":                         "Chemicals",
    "surfactant":                       "Chemicals",
    "pigment":                          "Chemicals",
    "commodity chemicals":              "Chemicals",

    # ── Capital Goods / Engineering ──
    "capital goods-non electrical":     "Capital Goods",
    "capital goods-electrical":         "Capital Goods",
    "capital goods":                    "Capital Goods",
    "industrial manufacturing":         "Capital Goods",
    "industrial machinery":             "Capital Goods",
    "engineering":                      "Capital Goods",
    "electrical equipment":             "Capital Goods",
    "electronic components":            "Capital Goods",
    "electronics":                      "Capital Goods",
    "heavy equipment":                  "Capital Goods",
    "heavy electrical":                 "Capital Goods",
    "industrial products":              "Capital Goods",
    "machine tools":                    "Capital Goods",
    "compressor":                       "Capital Goods",
    "pump":                             "Capital Goods",
    "valve":                            "Capital Goods",
    "bearing":                          "Capital Goods",
    "gear":                             "Capital Goods",
    "casting":                          "Capital Goods",
    "forging":                          "Capital Goods",
    "fastener":                         "Capital Goods",
    "wire":                             "Capital Goods",
    "cable":                            "Capital Goods",
    "transformer":                      "Capital Goods",
    "switchgear":                       "Capital Goods",
    "battery":                          "Capital Goods",
    "defence":                          "Capital Goods",
    "aerospace":                        "Capital Goods",
    "shipbuilding":                     "Capital Goods",
    "ship":                             "Capital Goods",
    "railway wagon":                    "Capital Goods",
    "locomotive":                       "Capital Goods",
    "boiler":                           "Capital Goods",
    "textile machinery":                "Capital Goods",
    "packaging machinery":              "Capital Goods",
    "diversified commercial":           "Capital Goods",
    "trading":                          "Capital Goods",
    "manufacturing":                    "Capital Goods",

    # ── Infrastructure / Logistics ──
    "infrastructure developers":        "Infrastructure",
    "infrastructure":                   "Infrastructure",
    "construction material":            "Infrastructure",
    "civil construction":               "Infrastructure",
    "road":                             "Infrastructure",
    "highway":                          "Infrastructure",
    "port":                             "Infrastructure",
    "airport":                          "Infrastructure",
    "logistics":                        "Infrastructure",
    "freight":                          "Infrastructure",
    "warehousing":                      "Infrastructure",
    "cold chain":                       "Infrastructure",
    "courier":                          "Infrastructure",
    "supply chain":                     "Infrastructure",
    "shipping":                         "Infrastructure",
    "fleet management":                 "Infrastructure",
    "water supply":                     "Infrastructure",
    "water treatment":                  "Infrastructure",
    "sanitation":                       "Infrastructure",
    "waste management":                 "Infrastructure",

    # ── Real Estate ──
    "real estate":                      "Real Estate",
    "realty":                           "Real Estate",
    "residential commercial":           "Real Estate",
    "commercial real estate":           "Real Estate",
    "residential real estate":          "Real Estate",
    "housing":                          "Real Estate",
    "construction":                     "Real Estate",
    "property":                         "Real Estate",

    # ── Telecom ──
    "telecom services":                 "Telecom",
    "telecom":                          "Telecom",
    "telecommunication":                "Telecom",
    "wireless":                         "Telecom",
    "broadband":                        "Telecom",
    "satellite":                        "Telecom",
    "data center":                      "Telecom",

    # ── Cement ──
    "cement & cement products":         "Cement",
    "cement":                           "Cement",
    "building materials":               "Cement",
    "ceramic":                          "Cement",
    "glass":                            "Cement",
    "pipes & fittings":                 "Cement",
    "sanitary ware":                    "Cement",
    "tiles":                            "Cement",
    "marble":                           "Cement",
    "granite":                          "Cement",

    # ── Textiles ──
    "textiles & apparel":               "Textiles",
    "textiles":                         "Textiles",
    "textile":                          "Textiles",
    "spinning":                         "Textiles",
    "weaving":                          "Textiles",
    "cotton":                           "Textiles",
    "yarn":                             "Textiles",
    "synthetic":                        "Textiles",
    "wool":                             "Textiles",
    "jute":                             "Textiles",
    "silk":                             "Textiles",
    "denim":                            "Textiles",

    # ── Diversified ──
    "diversified":                      "Diversified",
    "conglomerate":                     "Diversified",
    "miscellaneous":                    "Diversified",
}

# Direct symbol overrides — always win over any mapping
SECTOR_OVERRIDES: dict[str, str] = {
    # Nifty 50
    "RELIANCE":    "Energy",
    "TCS":         "Information Technology",
    "HDFCBANK":    "Financial Services",
    "ICICIBANK":   "Financial Services",
    "INFY":        "Information Technology",
    "HINDUNILVR":  "Consumer Staples",
    "ITC":         "Consumer Staples",
    "LT":          "Capital Goods",
    "SBIN":        "Financial Services",
    "BHARTIARTL":  "Telecom",
    "KOTAKBANK":   "Financial Services",
    "AXISBANK":    "Financial Services",
    "BAJFINANCE":  "Financial Services",
    "MARUTI":      "Automobile",
    "TATAMOTORS":  "Automobile",
    "SUNPHARMA":   "Healthcare",
    "WIPRO":       "Information Technology",
    "ULTRACEMCO":  "Cement",
    "ONGC":        "Energy",
    "NTPC":        "Energy",
    "POWERGRID":   "Energy",
    "COALINDIA":   "Metals & Mining",
    "TATASTEEL":   "Metals & Mining",
    "JSWSTEEL":    "Metals & Mining",
    "NESTLEIND":   "Consumer Staples",
    "M&M":         "Automobile",
    "ADANIPORTS":  "Infrastructure",
    "ADANIENT":    "Infrastructure",
    "ADANIGREEN":  "Energy",
    "TITAN":       "Consumer Discretionary",
    "ASIANPAINT":  "Consumer Discretionary",
    "DRREDDY":     "Healthcare",
    "CIPLA":       "Healthcare",
    "DIVISLAB":    "Healthcare",
    "APOLLOHOSP":  "Healthcare",
    # Nifty Next 50 additions
    "HCLTECH":     "Information Technology",
    "TECHM":       "Information Technology",
    "LTIM":        "Information Technology",
    "PERSISTENT":  "Information Technology",
    "COFORGE":     "Information Technology",
    "MPHASIS":     "Information Technology",
    "HEXAWARE":    "Information Technology",
    "ZOMATO":      "Consumer Discretionary",
    "NYKAA":       "Consumer Discretionary",
    "DMART":       "Consumer Staples",
    "TATACONSUM":  "Consumer Staples",
    "BRITANNIA":   "Consumer Staples",
    "DABUR":       "Consumer Staples",
    "MARICO":      "Consumer Staples",
    "GODREJCP":    "Consumer Staples",
    "EMAMILTD":    "Consumer Staples",
    "BAJAJ-AUTO":  "Automobile",
    "HEROMOTOCO":  "Automobile",
    "EICHERMOT":   "Automobile",
    "TVSMOTOR":    "Automobile",
    "BAJAJFINSV":  "Financial Services",
    "HDFCLIFE":    "Financial Services",
    "SBILIFE":     "Financial Services",
    "ICICIPRULI":  "Financial Services",
    "GICRE":       "Financial Services",
    "NIACL":       "Financial Services",
    "ICICIGI":     "Financial Services",
    "RECLTD":      "Financial Services",
    "PFC":         "Financial Services",
    "IRFC":        "Financial Services",
    "CHOLAFIN":    "Financial Services",
    "MUTHOOTFIN":  "Financial Services",
    "MANAPPURAM":  "Financial Services",
    "SHRIRAMFIN":  "Financial Services",
    "LICHSGFIN":   "Financial Services",
    "GODREJPROP":  "Real Estate",
    "DLF":         "Real Estate",
    "PRESTIGE":    "Real Estate",
    "OBEROIRLTY":  "Real Estate",
    "PHOENIXLTD":  "Real Estate",
    "BRIGADE":     "Real Estate",
    "SOBHA":       "Real Estate",
    "SUNPHARMA":   "Healthcare",
    "LUPIN":       "Healthcare",
    "AUROPHARMA":  "Healthcare",
    "BIOCON":      "Healthcare",
    "ALKEM":       "Healthcare",
    "TORNTPHARM":  "Healthcare",
    "IPCALAB":     "Healthcare",
    "GLAND":       "Healthcare",
    "ABBOTINDIA":  "Healthcare",
    "PFIZER":      "Healthcare",
    "SANOFI":      "Healthcare",
    "TATAPOWER":   "Energy",
    "ADANIPOWER":  "Energy",
    "TORNTPOWER":  "Energy",
    "CESC":        "Energy",
    "JSWENERGY":   "Energy",
    "SJVN":        "Energy",
    "NHPC":        "Energy",
    "HINDPETRO":   "Energy",
    "BPCL":        "Energy",
    "IOC":         "Energy",
    "GAIL":        "Energy",
    "PETRONET":    "Energy",
    "MGL":         "Energy",
    "IGL":         "Energy",
    "VEDL":        "Metals & Mining",
    "HINDALCO":    "Metals & Mining",
    "NATIONALUM":  "Metals & Mining",
    "NMDC":        "Metals & Mining",
    "SAIL":        "Metals & Mining",
    "JSWINFRA":    "Infrastructure",
    "GMRINFRA":    "Infrastructure",
    "CONCOR":      "Infrastructure",
    "BLUEDART":    "Infrastructure",
    "DELHIVERY":   "Infrastructure",
    "IRCTC":       "Infrastructure",
    "GRINFRA":     "Infrastructure",
    "IRB":         "Infrastructure",
    "ASHOKA":      "Infrastructure",
    "KPITTECH":    "Information Technology",
    "TATAELXSI":   "Information Technology",
    "LTTS":        "Information Technology",
    "INFY":        "Information Technology",
    "ULTRACEMCO":  "Cement",
    "AMBUJACEM":   "Cement",
    "ACC":         "Cement",
    "DALMIACEMEN": "Cement",
    "JKCEMENT":    "Cement",
    "SHREECEM":    "Cement",
    "HEIDELBERG":  "Cement",
    "INDUSINDBK":  "Financial Services",
    "FEDERALBNK":  "Financial Services",
    "IDFCFIRSTB":  "Financial Services",
    "BANDHANBNK":  "Financial Services",
    "RBLBANK":     "Financial Services",
    "CANBK":       "Financial Services",
    "BANKBARODA":  "Financial Services",
    "UNIONBANK":   "Financial Services",
    "PNB":         "Financial Services",
    "UCOBANK":     "Financial Services",
    "CENTRALBK":   "Financial Services",
    "IOB":         "Financial Services",
    "MAHABANK":    "Financial Services",
    "KARURVYSYA":  "Financial Services",
    "DCBBANK":     "Financial Services",
    "JKBANK":      "Financial Services",
    "SOUTHBANK":   "Financial Services",
    "CEATLTD":     "Automobile",
    "MRF":         "Automobile",
    "APOLLOTYRE":  "Automobile",
    "BALKRISIND":  "Automobile",
    "JKTYRE":      "Automobile",
    "BOSCHLTD":    "Automobile",
    "MOTHERSON":   "Automobile",
    "BHARATFORG":  "Automobile",
    "SUNDRMFAST":  "Automobile",
    "EXIDEIND":    "Automobile",
    "AMARARAJA":   "Automobile",
    "TATACOMM":    "Telecom",
    "IDEA":        "Telecom",
    "MTNL":        "Telecom",
    "TTML":        "Telecom",
    "PIIND":       "Chemicals",
    "SRF":         "Chemicals",
    "AAVAS":       "Financial Services",
    "HOMEFIRST":   "Financial Services",
    "APTUS":       "Financial Services",
    "AARTIIND":    "Chemicals",
    "DEEPAKNITR":  "Chemicals",
    "NAVINFLUOR":  "Chemicals",
    "FLUOROCHEM":  "Chemicals",
    "CLEAN":       "Chemicals",
    "TATACHEM":    "Chemicals",
    "GSFC":        "Chemicals",
    "COROMANDEL":  "Chemicals",
    "CHAMBALFERT": "Chemicals",
    "GNFC":        "Chemicals",
}


def _make_nse_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(NSE_HEADERS)
    try:
        session.get(NSE_SESSION_URL, timeout=10)
        time.sleep(0.5)
    except Exception as e:
        logger.warning(f"NSE session init failed: {e}")
    return session


# Company name keyword → sector (ordered: most specific first)
_NAME_KEYWORDS: list[tuple[str, str]] = [
    # Financial Services
    ("bank",            "Financial Services"),
    ("finance",         "Financial Services"),
    ("financial",       "Financial Services"),
    ("insurance",       "Financial Services"),
    ("nbfc",            "Financial Services"),
    ("leasing",         "Financial Services"),
    ("investment",      "Financial Services"),
    ("capital",         "Financial Services"),
    ("credit",          "Financial Services"),
    ("securities",      "Financial Services"),
    ("broking",         "Financial Services"),
    ("brokerage",       "Financial Services"),
    ("housing finance", "Financial Services"),
    ("microfinance",    "Financial Services"),
    # Information Technology
    ("software",        "Information Technology"),
    ("infotech",        "Information Technology"),
    ("infosys",         "Information Technology"),
    ("techno",          "Information Technology"),
    ("technology",      "Information Technology"),
    ("systems",         "Information Technology"),
    ("computers",       "Information Technology"),
    ("digital",         "Information Technology"),
    ("data",            "Information Technology"),
    ("cyber",           "Information Technology"),
    # Healthcare
    ("pharma",          "Healthcare"),
    ("drug",            "Healthcare"),
    ("medicine",        "Healthcare"),
    ("healthcare",      "Healthcare"),
    ("hospital",        "Healthcare"),
    ("diagnostic",      "Healthcare"),
    ("biotech",         "Healthcare"),
    ("lifescience",     "Healthcare"),
    ("life science",    "Healthcare"),
    ("medical",         "Healthcare"),
    ("labs",            "Healthcare"),
    ("lab ",            "Healthcare"),
    # Automobile
    ("auto",            "Automobile"),
    ("motor",           "Automobile"),
    ("vehicle",         "Automobile"),
    ("tractor",         "Automobile"),
    ("tyres",           "Automobile"),
    ("tyre",            "Automobile"),
    # Energy
    ("petroleum",       "Energy"),
    ("petrochem",       "Energy"),
    ("oil",             "Energy"),
    ("gas",             "Energy"),
    ("power",           "Energy"),
    ("energy",          "Energy"),
    ("solar",           "Energy"),
    ("wind energy",     "Energy"),
    ("coal",            "Energy"),
    # Consumer Staples
    ("food",            "Consumer Staples"),
    ("beverages",       "Consumer Staples"),
    ("agro",            "Consumer Staples"),
    ("agriculture",     "Consumer Staples"),
    ("dairy",           "Consumer Staples"),
    ("tobacco",         "Consumer Staples"),
    ("fmcg",            "Consumer Staples"),
    ("flour",           "Consumer Staples"),
    ("sugar",           "Consumer Staples"),
    ("edible",          "Consumer Staples"),
    # Consumer Discretionary
    ("hotel",           "Consumer Discretionary"),
    ("resort",          "Consumer Discretionary"),
    ("retail",          "Consumer Discretionary"),
    ("media",           "Consumer Discretionary"),
    ("entertainment",   "Consumer Discretionary"),
    ("fashion",         "Consumer Discretionary"),
    ("apparel",         "Consumer Discretionary"),
    ("jewel",           "Consumer Discretionary"),
    # Metals & Mining
    ("steel",           "Metals & Mining"),
    ("iron",            "Metals & Mining"),
    ("copper",          "Metals & Mining"),
    ("alumin",          "Metals & Mining"),
    ("zinc",            "Metals & Mining"),
    ("metal",           "Metals & Mining"),
    ("mining",          "Metals & Mining"),
    ("mineral",         "Metals & Mining"),
    # Chemicals
    ("chemical",        "Chemicals"),
    ("fertiliser",      "Chemicals"),
    ("fertilizer",      "Chemicals"),
    ("pesticide",       "Chemicals"),
    ("paint",           "Chemicals"),
    ("dye",             "Chemicals"),
    ("resin",           "Chemicals"),
    ("polymer",         "Chemicals"),
    ("petro",           "Chemicals"),
    ("agrochem",        "Chemicals"),
    # Capital Goods
    ("engineer",        "Capital Goods"),
    ("engineering",     "Capital Goods"),
    ("machinery",       "Capital Goods"),
    ("equipment",       "Capital Goods"),
    ("industrial",      "Capital Goods"),
    ("electric",        "Capital Goods"),
    ("electronics",     "Capital Goods"),
    ("pumps",           "Capital Goods"),
    ("valve",           "Capital Goods"),
    ("bearing",         "Capital Goods"),
    ("transmission",    "Capital Goods"),
    ("cable",           "Capital Goods"),
    ("wire",            "Capital Goods"),
    ("transformer",     "Capital Goods"),
    # Infrastructure
    ("infrastructure",  "Infrastructure"),
    ("port",            "Infrastructure"),
    ("airport",         "Infrastructure"),
    ("road",            "Infrastructure"),
    ("highway",         "Infrastructure"),
    ("logistics",       "Infrastructure"),
    ("shipping",        "Infrastructure"),
    ("railway",         "Infrastructure"),
    ("transport",       "Infrastructure"),
    ("telecom",         "Telecom"),
    # Cement
    ("cement",          "Cement"),
    # Real Estate
    ("realty",          "Real Estate"),
    ("real estate",     "Real Estate"),
    ("property",        "Real Estate"),
    ("builder",         "Real Estate"),
    ("developer",       "Real Estate"),
    # Textiles
    ("textile",         "Textiles"),
    ("cotton",          "Textiles"),
    ("yarn",            "Textiles"),
    ("fabric",          "Textiles"),
    ("spinning",        "Textiles"),
    ("weaving",         "Textiles"),
    ("denim",           "Textiles"),
    ("garment",         "Textiles"),
    ("apparel",         "Textiles"),
    # More Chemicals
    ("aromatics",       "Chemicals"),
    ("plastics",        "Chemicals"),
    ("plastic",         "Chemicals"),
    ("rubber",          "Chemicals"),
    ("lubricant",       "Chemicals"),
    ("coatings",        "Chemicals"),
    ("adhesive",        "Chemicals"),
    ("solvent",         "Chemicals"),
    ("benzoplast",      "Chemicals"),
    ("benzene",         "Chemicals"),
    ("refractory",      "Chemicals"),
    ("ink",             "Chemicals"),
    ("pigment",         "Chemicals"),
    ("agrochem",        "Chemicals"),
    ("organic",         "Chemicals"),
    # More Metals & Mining
    ("extrusion",       "Metals & Mining"),
    ("casting",         "Metals & Mining"),
    ("forging",         "Metals & Mining"),
    ("alloy",           "Metals & Mining"),
    ("graphite",        "Metals & Mining"),
    ("foundry",         "Metals & Mining"),
    ("smelting",        "Metals & Mining"),
    ("ferrous",         "Metals & Mining"),
    ("nonferrous",      "Metals & Mining"),
    ("stainless",       "Metals & Mining"),
    ("pipes",           "Metals & Mining"),
    ("tubes",           "Metals & Mining"),
    # More Capital Goods
    ("welding",         "Capital Goods"),
    ("compressor",      "Capital Goods"),
    ("gears",           "Capital Goods"),
    ("boiler",          "Capital Goods"),
    ("turbine",         "Capital Goods"),
    ("crane",           "Capital Goods"),
    ("elevator",        "Capital Goods"),
    ("hydraulic",       "Capital Goods"),
    ("pneumatic",       "Capital Goods"),
    ("projects",        "Capital Goods"),
    ("precision",       "Capital Goods"),
    ("fabrication",     "Capital Goods"),
    ("switchgear",      "Capital Goods"),
    ("motor",           "Capital Goods"),
    # Consumer Discretionary additions
    ("packaging",       "Consumer Discretionary"),
    ("paper",           "Consumer Discretionary"),
    ("printing",        "Consumer Discretionary"),
    ("publishing",      "Consumer Discretionary"),
    ("glass",           "Consumer Discretionary"),
    ("tiles",           "Consumer Discretionary"),
    ("ceramic",         "Consumer Discretionary"),
    ("sanitary",        "Consumer Discretionary"),
    ("furniture",       "Consumer Discretionary"),
    ("decor",           "Consumer Discretionary"),
    ("travel",          "Consumer Discretionary"),
    ("tourism",         "Consumer Discretionary"),
    ("film",            "Consumer Discretionary"),
    ("broadcast",       "Consumer Discretionary"),
    ("vyapar",          "Consumer Discretionary"),
    ("trading",         "Consumer Discretionary"),
    # Consumer Staples additions
    ("agri",            "Consumer Staples"),
    ("seeds",           "Consumer Staples"),
    ("spice",           "Consumer Staples"),
    ("tea",             "Consumer Staples"),
    ("coffee",          "Consumer Staples"),
    ("brewery",         "Consumer Staples"),
    ("distillery",      "Consumer Staples"),
    ("breweries",       "Consumer Staples"),
    ("nutrition",       "Consumer Staples"),
    # Infrastructure additions
    ("aviation",        "Infrastructure"),
    ("airline",         "Infrastructure"),
    ("courier",         "Infrastructure"),
    ("warehouse",       "Infrastructure"),
    ("cold storage",    "Infrastructure"),
    # More Financial Services
    ("leasing",         "Financial Services"),
    ("factoring",       "Financial Services"),
    ("exchang",         "Financial Services"),
    # Telecom additions
    ("wireless",        "Telecom"),
    ("network",         "Telecom"),
    ("broadband",       "Telecom"),
    # Cement / Building Materials
    ("granite",         "Cement"),
    ("marble",          "Cement"),
    ("building material","Cement"),
    # Healthcare additions
    ("health",          "Healthcare"),
    ("dental",          "Healthcare"),
    ("optical",         "Healthcare"),
    ("vision",          "Healthcare"),
    # Common BSE name fragments
    ("infra",           "Infrastructure"),
    ("chem",            "Chemicals"),
    ("plast",           "Chemicals"),
    ("carbide",         "Chemicals"),
    ("acryl",           "Chemicals"),
    ("petrol",          "Energy"),
    ("refiner",         "Energy"),
    ("forge",           "Capital Goods"),
    ("engines",         "Capital Goods"),
    ("pump",            "Capital Goods"),
    ("instruments",     "Capital Goods"),
    ("holdings",        "Financial Services"),
    ("invest",          "Financial Services"),
    ("venture",         "Financial Services"),
    ("agro",            "Consumer Staples"),
    ("wines",           "Consumer Staples"),
    ("spirits",         "Consumer Staples"),
    ("breweries",       "Consumer Staples"),
    ("watch",           "Consumer Discretionary"),
    ("clock",           "Consumer Discretionary"),
    ("sport",           "Consumer Discretionary"),
    ("fitness",         "Consumer Discretionary"),
    ("leisure",         "Consumer Discretionary"),
    # Additional patterns
    ("oxygen",          "Chemicals"),
    ("gelatin",         "Chemicals"),
    ("carbon",          "Chemicals"),
    ("oxygen",          "Chemicals"),
    ("gas ",            "Energy"),
    ("gases",           "Chemicals"),
    ("construction",    "Infrastructure"),
    ("build",           "Infrastructure"),
    ("housin",          "Real Estate"),
    ("cosmetic",        "Consumer Discretionary"),
    ("cooker",          "Consumer Discretionary"),
    ("advisory",        "Financial Services"),
    ("consult",         "Financial Services"),
    ("education",       "Consumer Discretionary"),
    ("school",          "Consumer Discretionary"),
    ("educ",            "Consumer Discretionary"),
    ("syntex",          "Textiles"),
    ("fibre",           "Textiles"),
    ("fiber",           "Textiles"),
    ("plantation",      "Consumer Staples"),
    ("estates",         "Consumer Staples"),
    ("estate",          "Real Estate"),
    ("housing",         "Real Estate"),
    ("realty",          "Real Estate"),
    ("dev ",            "Real Estate"),
    ("developer",       "Real Estate"),
    ("printing",        "Consumer Discretionary"),
    ("stationery",      "Consumer Discretionary"),
    # Textiles additions
    ("mills",           "Textiles"),
    ("silk",            "Textiles"),
    ("nylon",           "Textiles"),
    ("polyte",          "Textiles"),
    ("wooll",           "Textiles"),
    ("loom",            "Textiles"),
    # Capital Goods additions
    ("batteries",       "Capital Goods"),
    ("battery",         "Capital Goods"),
    ("magnets",         "Capital Goods"),
    ("rectif",          "Capital Goods"),
    ("switchboard",     "Capital Goods"),
    ("controls",        "Capital Goods"),
    ("automation",      "Capital Goods"),
    ("robotics",        "Capital Goods"),
    # IT
    ("computer",        "Information Technology"),
    ("microchip",       "Information Technology"),
    ("semicon",         "Information Technology"),
    # Chemicals
    ("kem",             "Chemicals"),
    ("kemi",            "Chemicals"),
    ("coke",            "Chemicals"),
    ("naptha",          "Chemicals"),
    ("soda",            "Chemicals"),
    ("acid",            "Chemicals"),
    ("alkali",          "Chemicals"),
    ("oxide",           "Chemicals"),
    ("nitro",           "Chemicals"),
    ("sulph",           "Chemicals"),
    ("chlor",           "Chemicals"),
    ("fluor",           "Chemicals"),
    # Consumer Staples
    ("soya",            "Consumer Staples"),
    ("soya",            "Consumer Staples"),
    ("rice",            "Consumer Staples"),
    ("wheat",           "Consumer Staples"),
    ("milling",         "Consumer Staples"),
    ("vanaspati",       "Consumer Staples"),
    ("bakery",          "Consumer Staples"),
    ("confection",      "Consumer Staples"),
    # Financial Services
    ("micro finance",   "Financial Services"),
    ("trading company", "Financial Services"),
    ("commerce",        "Financial Services"),
    # Infrastructure
    ("highway",         "Infrastructure"),
    ("port ",           "Infrastructure"),
    ("airports",        "Infrastructure"),
    ("pipeline",        "Infrastructure"),
]


def _name_to_sector(company_name: str) -> str:
    """Classify a stock by company name keywords. Returns 'Unknown' if no match."""
    if not company_name:
        return "Unknown"
    name = company_name.lower()
    for keyword, sector in _NAME_KEYWORDS:
        if keyword in name:
            return sector
    return "Unknown"


def _industry_to_sector(industry: str) -> str:
    """Case-insensitive multi-strategy lookup. Returns 'Unknown' if no match."""
    if not industry:
        return "Unknown"
    key = industry.strip().lower()

    # 1. Exact match
    if key in INDUSTRY_TO_SECTOR:
        return INDUSTRY_TO_SECTOR[key]

    # 2. Pattern contained in key  (pattern is substring of industry string)
    for pattern, sector in INDUSTRY_TO_SECTOR.items():
        if pattern in key:
            return sector

    # 3. Key is substring of a longer pattern  (e.g. key="it" inside "it-software")
    for pattern, sector in INDUSTRY_TO_SECTOR.items():
        if key in pattern and len(key) >= 3:
            return sector

    return "Unknown"


def fetch_nse_stocks() -> pd.DataFrame:
    """
    Fetch NSE stocks from three NSE APIs:
    1. NIFTY TOTAL MARKET (750 stocks — company names + industry from meta)
    2. Sector-specific indices (BANK, IT, PHARMA, etc.) — builds authoritative sector_map
    3. Extended index list (MIDCAP 150, SMALLCAP 250, NEXT 50, etc.) — more stocks + meta
    4. Pre-open ALL (~1975 symbols — fills remaining gaps with live prices)
    """
    logger.info("Fetching NSE stocks...")
    session = _make_nse_session()
    stocks: dict[str, dict] = {}
    sector_map: dict[str, str] = {}  # symbol → confirmed sector from sector index

    # ── Step 1: NIFTY TOTAL MARKET ─────────────────────────────────────────────
    try:
        r = session.get(NSE_TOTAL_MARKET_URL, timeout=15)
        r.raise_for_status()
        data = r.json()
        for item in data.get("data", []):
            sym = item.get("symbol", "").strip()
            if not sym or sym == "NIFTY TOTAL MARKET":
                continue
            meta = item.get("meta", {})
            industry = meta.get("industry", "")
            company = meta.get("companyName", sym)
            sector = (
                SECTOR_OVERRIDES.get(sym)
                or _industry_to_sector(industry)
                or _name_to_sector(company)
            )
            stocks[sym] = {
                "symbol":       sym,
                "company_name": company,
                "isin":         meta.get("isin", ""),
                "industry_raw": industry,
                "sector":       sector,
                "last_price":   item.get("lastPrice"),
                "exchange":     "NSE",
                "yf_symbol":    sym + ".NS",
            }
        logger.info(f"  NSE TOTAL MARKET: {len(stocks)} stocks")
    except Exception as e:
        logger.error(f"  NSE TOTAL MARKET failed: {e}")

    # ── Step 2: Sector indices — builds authoritative sector_map ───────────────
    for index_name, sector_label in NSE_SECTOR_INDICES:
        try:
            url = NSE_INDEX_URL.format(urllib.parse.quote(index_name))
            r = session.get(url, timeout=15)
            if r.status_code != 200:
                continue
            data = r.json()
            count = 0
            for item in data.get("data", []):
                sym = item.get("symbol", "").strip()
                if not sym or sym == index_name:
                    continue
                # Explicit sector label beats generic index membership
                if sector_label:
                    sector_map[sym] = sector_label

                if sym not in stocks:
                    meta = item.get("meta", {})
                    industry = meta.get("industry", "")
                    stocks[sym] = {
                        "symbol":       sym,
                        "company_name": meta.get("companyName", sym),
                        "isin":         meta.get("isin", ""),
                        "industry_raw": industry,
                        "sector":       sector_label or _industry_to_sector(industry) or "Unknown",
                        "last_price":   item.get("lastPrice"),
                        "exchange":     "NSE",
                        "yf_symbol":    sym + ".NS",
                    }
                count += 1
            logger.info(f"  {index_name}: {count} stocks")
            time.sleep(0.3)
        except Exception as e:
            logger.warning(f"  {index_name} failed: {e}")

    # Apply sector_map (SECTOR_OVERRIDES always win)
    for sym, sector in sector_map.items():
        if sym in stocks and sym not in SECTOR_OVERRIDES:
            stocks[sym]["sector"] = sector

    # ── Step 3: Pre-open ALL — fills remaining 1200+ with live prices ──────────
    try:
        r = session.get(NSE_ALL_PREOPEN_URL, timeout=15)
        r.raise_for_status()
        data = r.json()
        new_count = 0
        for item in data.get("data", []):
            meta = item.get("metadata", {})
            sym = meta.get("symbol", "").strip()
            if not sym or sym in stocks:
                continue
            # Try to get industry from detail field
            industry = meta.get("industry", "") or ""
            company = meta.get("companyName", sym) or sym
            sector = (
                SECTOR_OVERRIDES.get(sym)
                or _industry_to_sector(industry)
                or _name_to_sector(company)
            )
            stocks[sym] = {
                "symbol":       sym,
                "company_name": company,
                "isin":         meta.get("isin", "") or "",
                "industry_raw": industry,
                "sector":       sector,
                "last_price":   meta.get("lastPrice"),
                "exchange":     "NSE",
                "yf_symbol":    sym + ".NS",
            }
            new_count += 1
        logger.info(f"  NSE pre-open added {new_count} more → total {len(stocks)} NSE stocks")
    except Exception as e:
        logger.error(f"  NSE pre-open failed: {e}")

    if not stocks:
        return pd.DataFrame(columns=["symbol", "company_name", "isin", "sector",
                                     "exchange", "yf_symbol", "last_price"])

    df = pd.DataFrame(list(stocks.values()))
    logger.info(f"NSE total: {len(df)} stocks")
    return df


def fetch_bse_stocks() -> pd.DataFrame:
    """
    Fetch all active BSE equity stocks (~4800) from BSE's open API.
    """
    logger.info("Fetching BSE stocks...")
    try:
        r = requests.get(BSE_LIST_URL, headers=BSE_HEADERS, timeout=30)
        r.raise_for_status()
        data = r.json()
        if not isinstance(data, list):
            data = data.get("Table", [])

        rows = []
        for item in data:
            scrip_code = str(item.get("SCRIP_CD", "")).strip()
            scrip_id   = str(item.get("scrip_id", "")).strip()
            company    = str(item.get("Issuer_Name") or item.get("Scrip_Name", scrip_code)).strip()
            isin       = str(item.get("ISIN_NUMBER", "")).strip()
            industry   = str(item.get("INDUSTRY") or item.get("Industry", "") or "").strip()
            symbol     = scrip_id if scrip_id else scrip_code

            # Sector: override > industry string > company name keyword
            sector = (
                SECTOR_OVERRIDES.get(symbol)
                or _industry_to_sector(industry)
                or _name_to_sector(company)
            )
            rows.append({
                "symbol":       symbol,
                "bse_code":     scrip_code,
                "company_name": company,
                "isin":         isin,
                "industry_raw": industry,
                "sector":       sector,
                "exchange":     "BSE",
                "yf_symbol":    scrip_code + ".BO",
                "last_price":   None,
            })

        df = pd.DataFrame(rows)
        logger.info(f"BSE total: {len(df)} stocks")
        return df
    except Exception as e:
        logger.error(f"BSE fetch failed: {e}")
        return pd.DataFrame(columns=["symbol", "company_name", "isin", "sector",
                                     "exchange", "yf_symbol", "last_price"])


def merge_and_deduplicate(nse_df: pd.DataFrame, bse_df: pd.DataFrame) -> pd.DataFrame:
    """
    Merge NSE + BSE. Deduplicate by ISIN — prefer NSE ticker when both exist.
    BSE-only stocks inherit sector from the NSE counterpart via ISIN matching.
    """
    # Build ISIN → sector from NSE (authoritative)
    isin_to_sector: dict[str, str] = {}
    for _, row in nse_df.iterrows():
        isin = str(row.get("isin", ""))
        sector = str(row.get("sector", ""))
        if len(isin) == 12 and sector and sector != "Unknown":
            isin_to_sector[isin] = sector

    # Build company-name → sector map (normalised) for BSE fallback
    name_to_sector: dict[str, str] = {}
    for _, row in nse_df.iterrows():
        sector = str(row.get("sector", ""))
        if sector and sector != "Unknown":
            name = str(row.get("company_name", "")).lower()[:20]
            if len(name) > 4:
                name_to_sector[name] = sector

    def bse_sector(row) -> str:
        # 1. Symbol override
        sym = str(row.get("symbol", ""))
        if sym in SECTOR_OVERRIDES:
            return SECTOR_OVERRIDES[sym]
        # 2. Already has a sector from BSE industry field
        s = str(row.get("sector", "Unknown"))
        if s and s != "Unknown":
            return s
        # 3. ISIN propagation from NSE
        isin = str(row.get("isin", ""))
        if len(isin) == 12 and isin in isin_to_sector:
            return isin_to_sector[isin]
        # 4. Company name prefix match from NSE
        bse_name = str(row.get("company_name", "")).lower()[:20]
        if bse_name and bse_name in name_to_sector:
            return name_to_sector[bse_name]
        # 5. Industry field re-try (BSE industry may differ from NSE)
        industry = str(row.get("industry_raw", ""))
        mapped = _industry_to_sector(industry)
        if mapped != "Unknown":
            return mapped
        # 6. Company name keyword matching
        name_mapped = _name_to_sector(str(row.get("company_name", "")))
        if name_mapped != "Unknown":
            return name_mapped
        return "Unknown"

    bse_df = bse_df.copy()
    bse_df["sector"] = bse_df.apply(bse_sector, axis=1)

    combined = pd.concat([nse_df, bse_df], ignore_index=True)

    has_isin  = combined[combined["isin"].str.len() == 12].copy()
    no_isin   = combined[combined["isin"].str.len() != 12].copy()

    has_isin["_priority"] = has_isin["exchange"].map({"NSE": 0, "BSE": 1})
    has_isin = has_isin.sort_values("_priority")
    deduped  = has_isin.drop_duplicates(subset="isin", keep="first")

    result = pd.concat([deduped, no_isin], ignore_index=True)
    result = result.drop(columns=["_priority"], errors="ignore")

    # Final pass — resolve any remaining Unknown
    def resolve_sector(row) -> str:
        sym = str(row.get("symbol", ""))
        if sym in SECTOR_OVERRIDES:
            return SECTOR_OVERRIDES[sym]
        s = str(row.get("sector", ""))
        if s and s != "Unknown":
            return s
        return "Unknown"

    result["sector"]     = result.apply(resolve_sector, axis=1)
    result["market_cap"] = None
    result["pe_ratio"]   = None
    result["beta"]       = None

    known   = result[result["sector"] != "Unknown"]["sector"].count()
    unknown = (result["sector"] == "Unknown").sum()
    logger.info(
        f"Universe: {len(nse_df)} NSE + {len(bse_df)} BSE → {len(result)} unique stocks "
        f"({known} with sector, {unknown} still Unknown)"
    )
    return result.reset_index(drop=True)


def get_stocks_by_sector(df: pd.DataFrame) -> dict[str, list[dict]]:
    sectors: dict[str, list[dict]] = {}
    for _, row in df.iterrows():
        sector = row.get("sector") or "Unknown"
        if sector not in sectors:
            sectors[sector] = []
        sectors[sector].append({
            "symbol":       row["symbol"],
            "company_name": row["company_name"],
            "exchange":     row["exchange"],
            "isin":         row.get("isin", ""),
            "yf_symbol":    row["yf_symbol"],
            "sector":       sector,
            "industry":     row.get("industry_raw", ""),
            "last_price":   row.get("last_price"),
            "market_cap":   row.get("market_cap"),
            "pe_ratio":     row.get("pe_ratio"),
            "beta":         row.get("beta"),
        })
    return {
        sector: sorted(stocks, key=lambda s: s["symbol"])
        for sector, stocks in sorted(sectors.items())
    }


def build_universe(enrich: bool = False) -> tuple[pd.DataFrame, dict[str, list[dict]]]:
    nse_df = fetch_nse_stocks()
    bse_df = fetch_bse_stocks()
    universe_df = merge_and_deduplicate(nse_df, bse_df)
    by_sector = get_stocks_by_sector(universe_df)
    return universe_df, by_sector
