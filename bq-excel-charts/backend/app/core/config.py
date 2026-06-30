"""Configuración del backend (variables de entorno / .env)."""

from __future__ import annotations

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file=".env", env_prefix="BQEC_", extra="ignore")

    # Ruta al modelo semántico YAML.
    semantic_model_path: str = "app/semantic/example_model.yaml"

    # Proyecto de facturación de BigQuery. Si está vacío, el backend funciona en
    # modo "dry-run": compila SQL pero no ejecuta (útil para desarrollar la UX).
    gcp_project: str | None = None

    # Ruta al JSON de la service account. Las credenciales viven SOLO en el
    # backend, nunca en el cliente/add-in.
    google_application_credentials: str | None = None

    # Orígenes permitidos para CORS (el add-in se sirve en localhost en dev).
    cors_origins: list[str] = ["https://localhost:3000", "http://localhost:3000"]


settings = Settings()
