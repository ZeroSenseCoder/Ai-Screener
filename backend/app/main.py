import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.core.state import load_universe
from app.api.v1 import universe, screener, stocks, macro, pro_screener, valuation

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s — %(message)s",
)
logger = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("Starting up — initialising database...")
    from app.db.database import init_db
    await init_db()
    logger.info("Loading stock universe...")
    await load_universe()
    yield
    logger.info("Shutting down.")


app = FastAPI(
    title="Indian Stock Screener API",
    description="NSE + BSE stock screener with technical indicators, news, and macro data",
    version="1.0.0",
    lifespan=lifespan,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://localhost:3000"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(universe.router, prefix="/api/v1")
app.include_router(screener.router, prefix="/api/v1")
app.include_router(stocks.router, prefix="/api/v1")
app.include_router(macro.router, prefix="/api/v1")
app.include_router(pro_screener.router, prefix="/api/v1")
app.include_router(valuation.router, prefix="/api/v1")


@app.get("/api/v1/health")
async def health():
    from app.core.state import app_state
    return {
        "status": "ok",
        "universe_loaded": app_state.universe_df is not None,
        "total_stocks": len(app_state.universe_df) if app_state.universe_df is not None else 0,
        "last_updated": app_state.last_updated,
        "enrichment_status": app_state.enrichment_status,
        "enrichment_progress": app_state.enrichment_progress,
    }
