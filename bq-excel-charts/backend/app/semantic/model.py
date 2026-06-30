"""Capa semántica.

Define datasets, dimensiones y métricas una sola vez. La UI consume estas
definiciones para que el usuario nunca escriba SQL: solo elige campos. Además,
es la lista blanca que protege frente a inyección SQL y costes descontrolados
en BigQuery (solo se puede consultar lo que está definido aquí).
"""

from __future__ import annotations

from pathlib import Path

import yaml
from pydantic import BaseModel, Field


class Dimension(BaseModel):
    """Un eje por el que agrupar (categórico o temporal)."""

    name: str = Field(..., description="Identificador estable usado por la API/UI")
    label: str = Field(..., description="Nombre legible mostrado en la UI")
    # Expresión SQL de BigQuery. Puede ser una columna o una transformación
    # controlada por nosotros (p. ej. FORMAT_DATE). Nunca proviene del usuario.
    sql: str = Field(..., description="Expresión SQL de la dimensión")
    type: str = Field(default="string", description="string | date | number")


class Measure(BaseModel):
    """Un valor numérico agregado (lo que se mide)."""

    name: str
    label: str
    sql: str = Field(..., description="Expresión agregada, p. ej. SUM(amount)")
    format: str | None = Field(default=None, description="Formato sugerido, p. ej. '#,##0'")


class Dataset(BaseModel):
    """Un modelo consultable: una tabla/vista de BigQuery más sus campos."""

    name: str
    label: str
    table: str = Field(..., description="Referencia completa `proyecto.dataset.tabla`")
    description: str | None = None
    dimensions: list[Dimension] = Field(default_factory=list)
    measures: list[Measure] = Field(default_factory=list)
    # Límite duro de bytes facturados por consulta (control de coste).
    max_bytes_billed: int | None = Field(default=None)

    def dimension(self, name: str) -> Dimension | None:
        return next((d for d in self.dimensions if d.name == name), None)

    def measure(self, name: str) -> Measure | None:
        return next((m for m in self.measures if m.name == name), None)


class SemanticModel(BaseModel):
    datasets: list[Dataset] = Field(default_factory=list)

    def dataset(self, name: str) -> Dataset | None:
        return next((d for d in self.datasets if d.name == name), None)

    @classmethod
    def from_yaml(cls, path: str | Path) -> "SemanticModel":
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        return cls(**data)
