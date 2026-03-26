from pydantic_settings import BaseSettings
from functools import lru_cache


class Settings(BaseSettings):
    news_api_key: str = ""
    database_url: str = "sqlite+aiosqlite:///./fintech.db"
    environment: str = "development"

    class Config:
        env_file = ".env"


@lru_cache
def get_settings() -> Settings:
    return Settings()
