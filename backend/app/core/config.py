from functools import lru_cache
from pathlib import Path
from pydantic_settings import BaseSettings

class Settings(BaseSettings):
  app_name: str = "Non Language Communication Coach"
  database_url: str = f"sqlite:///{Path(__file__).resolve().parent.parent / 'nlc.db'}"
  media_dir: Path = Path(__file__).resolve().parent.parent / "media"
  google_client_id: str = "demo-client"
  google_redirect_uri: str = "http://localhost:5173/auth/callback"

  class Config:
    env_file = ".env"
    env_file_encoding = "utf-8"

@lru_cache
def get_settings() -> Settings:
  settings = Settings()
  settings.media_dir.mkdir(exist_ok=True)
  return settings
