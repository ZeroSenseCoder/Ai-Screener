"""
News service — three layers (all free):
  1. Yahoo Finance news via yfinance (PRIMARY — no key, always works)
  2. RSS feeds — Economic Times, Moneycontrol, Mint, Business Standard
  3. NewsAPI.org (OPTIONAL — 100 req/day free tier, requires key)
"""

import logging
from datetime import datetime, timedelta, timezone
from typing import Optional

import feedparser
import httpx

logger = logging.getLogger(__name__)

RSS_FEEDS = [
    ("Economic Times",    "https://economictimes.indiatimes.com/markets/rss.cms"),
    ("Moneycontrol",      "https://www.moneycontrol.com/rss/marketreports.xml"),
    ("Mint",              "https://www.livemint.com/rss/markets"),
    ("Business Standard", "https://www.business-standard.com/rss/markets-106.rss"),
]

HIGH_KEYWORDS = [
    "quarterly results", "q1", "q2", "q3", "q4", "profit", "earnings",
    "acquisition", "merger", "block deal", "sebi", "rbi", "ipo", "fpo",
    "dividend", "buyback", "rating upgrade", "rating downgrade",
    "insider trading", "fraud", "bankruptcy", "default", "delisting",
    "takeover", "open offer", "rights issue", "bonus", "split",
    "management change", "ceo", "md resign", "promoter pledge",
    "order win", "mega deal", "stake sale",
]

MEDIUM_KEYWORDS = [
    "expansion", "capex", "capacity", "jv", "joint venture", "stake",
    "contract", "guidance", "outlook", "revised", "analyst",
    "target price", "buy", "sell", "hold", "upgrade", "downgrade",
    "supply chain", "export", "import", "anti-dumping",
]


def score_importance(title: str, description: str = "") -> str:
    text = (title + " " + (description or "")).lower()
    if any(kw in text for kw in HIGH_KEYWORDS):
        return "high"
    if any(kw in text for kw in MEDIUM_KEYWORDS):
        return "medium"
    return "low"


def _mentions(text: str, symbol: str, company_name: str) -> bool:
    t = text.lower()
    sym = symbol.lower().replace(".ns", "").replace(".bo", "")
    co  = company_name.lower()
    # Match symbol (exact word) or first 8 chars of company name
    return (
        f" {sym} " in f" {t} "
        or (len(co) > 4 and co[:8] in t)
    )


def fetch_rss_news(symbol: str, company_name: str, days: int = 7) -> list[dict]:
    cutoff = datetime.now(tz=timezone.utc) - timedelta(days=days)
    articles = []
    for feed_name, url in RSS_FEEDS:
        try:
            feed = feedparser.parse(url)
            for entry in feed.entries:
                title = entry.get("title", "")
                desc  = entry.get("summary", "") or entry.get("description", "")
                pub   = entry.get("published_parsed")
                if pub:
                    pub_dt = datetime(*pub[:6], tzinfo=timezone.utc)
                    if pub_dt < cutoff:
                        continue
                if not _mentions(title + " " + desc, symbol, company_name):
                    continue
                importance = score_importance(title, desc)
                articles.append({
                    "title":        title,
                    "url":          entry.get("link", ""),
                    "source":       feed_name,
                    "published_at": entry.get("published", ""),
                    "summary":      desc[:400] if desc else "",
                    "thumbnail":    "",
                    "importance":   importance,
                })
        except Exception as e:
            logger.warning(f"RSS {feed_name}: {e}")
    return articles


async def fetch_newsapi(symbol: str, company_name: str, api_key: str, days: int = 7) -> list[dict]:
    if not api_key:
        return []
    from_date = (datetime.utcnow() - timedelta(days=days)).strftime("%Y-%m-%d")
    try:
        async with httpx.AsyncClient(timeout=10) as client:
            resp = await client.get("https://newsapi.org/v2/everything", params={
                "q":       f"{company_name} OR {symbol}",
                "from":    from_date,
                "sortBy":  "relevancy",
                "language": "en",
                "pageSize": 20,
                "apiKey":  api_key,
            })
            articles = []
            for item in resp.json().get("articles", []):
                title = item.get("title", "")
                desc  = item.get("description", "") or ""
                importance = score_importance(title, desc)
                articles.append({
                    "title":        title,
                    "url":          item.get("url", ""),
                    "source":       item.get("source", {}).get("name", "NewsAPI"),
                    "published_at": item.get("publishedAt", ""),
                    "summary":      desc[:400],
                    "thumbnail":    item.get("urlToImage", "") or "",
                    "importance":   importance,
                })
            return articles
    except Exception as e:
        logger.warning(f"NewsAPI: {e}")
        return []


async def get_stock_news(
    symbol: str,
    company_name: str,
    yf_news: list[dict],         # already fetched by market_data.fetch_yf_news
    api_key: str = "",
    days: int = 7,
) -> list[dict]:
    """
    Merge all three news sources.
    yf_news comes pre-fetched from the stocks endpoint (avoids double yfinance call).
    """
    # Score yfinance news
    scored_yf = []
    for a in yf_news:
        a["importance"] = score_importance(a.get("title", ""), a.get("summary", ""))
        scored_yf.append(a)

    # RSS news
    rss = fetch_rss_news(symbol, company_name, days)

    # Optional NewsAPI
    api_articles = await fetch_newsapi(symbol, company_name, api_key, days) if api_key else []

    # Combine; deduplicate by URL
    seen: set[str] = set()
    combined = []
    for article in scored_yf + rss + api_articles:
        url = article.get("url", "")
        key = url or article.get("title", "")
        if key and key not in seen:
            seen.add(key)
            combined.append(article)

    # Sort: high importance first, then by date desc
    order = {"high": 0, "medium": 1, "low": 2}
    combined.sort(key=lambda x: (order.get(x.get("importance", "low"), 2), ""))

    return combined
