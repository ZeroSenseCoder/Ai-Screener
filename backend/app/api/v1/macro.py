from fastapi import APIRouter
from app.services.macro_service import (
    fetch_global_indices,
    fetch_forex,
    fetch_commodities,
    fetch_india_vix,
    fetch_fii_dii,
    get_macro_overview,
)

router = APIRouter(prefix="/macro", tags=["Macro"])


@router.get("/overview")
async def macro_overview():
    """All macro data in one payload."""
    return await get_macro_overview()


@router.get("/indices")
async def global_indices():
    return fetch_global_indices()


@router.get("/forex")
async def forex():
    return fetch_forex()


@router.get("/commodities")
async def commodities():
    return fetch_commodities()


@router.get("/vix")
async def india_vix():
    return fetch_india_vix()


@router.get("/fii-dii")
async def fii_dii():
    return await fetch_fii_dii()
