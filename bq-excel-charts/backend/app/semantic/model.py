"""Capa semántica.

Define datasets, selectores, dimensiones y métricas una sola vez. La UI consume
estas definiciones para construir el formulario (incluidos los SELECTORES EN
CASCADA) sin que el usuario escriba SQL. Además, es la lista blanca que protege
frente a inyección SQL y costes descontrolados en BigQuery.

Conceptos:
- Selector: campo de configuración con opciones (p. ej. "Tipo de activo" →
  Renta fija / Renta variable). Al elegirlo, se aplica como filtro.
- applies_to: una métrica o dimensión puede declarar para qué valores de un
  selector es válida. La UI solo la muestra cuando la selección encaja
  (p. ej. "Duración" solo para Renta fija).
"""

from __future__ import annotations

from pathlib import Path

import yaml
from pydantic import BaseModel, Field

# applies_to es un mapa {nombre_selector: [valores_permitidos]}.
# Una métrica/dimensión está disponible si, para cada selector que menciona,
# el valor elegido está en su lista.
AppliesTo = dict[str, list[str]]


def _matches(applies_to: AppliesTo | None, selections: dict[str, str]) -> bool:
    """¿Es aplicable este campo dada la selección actual del usuario?"""
    if not applies_to:
        return True  # sin restricción → siempre disponible
    for selector_name, allowed in applies_to.items():
        chosen = selections.get(selector_name)
        if chosen is None or chosen not in allowed:
            return False
    return True


class FieldOption(BaseModel):
    value: str
    label: str


class Selector(BaseModel):
    """Campo de configuración con opciones que condiciona el resto del formulario."""

    name: str
    label: str
    column: str = Field(..., description="Columna de BigQuery sobre la que filtra")
    type: str = Field(default="string", description="string | number | date")
    options: list[FieldOption] = Field(default_factory=list)

    @property
    def sql(self) -> str:
        return self.column


class Dimension(BaseModel):
    """Un eje por el que agrupar (categórico o temporal)."""

    name: str
    label: str
    sql: str = Field(..., description="Expresión SQL; nunca proviene del usuario")
    type: str = Field(default="string", description="string | date | number")
    applies_to: AppliesTo | None = None


class Measure(BaseModel):
    """Un valor numérico agregado (lo que se mide)."""

    name: str
    label: str
    sql: str = Field(..., description="Expresión agregada, p. ej. SUM(amount)")
    format: str | None = None
    applies_to: AppliesTo | None = None


class Dataset(BaseModel):
    name: str
    label: str
    table: str = Field(..., description="Referencia completa `proyecto.dataset.tabla`")
    description: str | None = None
    selectors: list[Selector] = Field(default_factory=list)
    dimensions: list[Dimension] = Field(default_factory=list)
    measures: list[Measure] = Field(default_factory=list)
    max_bytes_billed: int | None = None

    def selector(self, name: str) -> Selector | None:
        return next((s for s in self.selectors if s.name == name), None)

    def dimension(self, name: str) -> Dimension | None:
        return next((d for d in self.dimensions if d.name == name), None)

    def measure(self, name: str) -> Measure | None:
        return next((m for m in self.measures if m.name == name), None)

    def filterable(self, name: str) -> Selector | Dimension | None:
        """Campo sobre el que se puede filtrar: un selector o una dimensión."""
        return self.selector(name) or self.dimension(name)

    def available_measures(self, selections: dict[str, str]) -> list[Measure]:
        return [m for m in self.measures if _matches(m.applies_to, selections)]

    def available_dimensions(self, selections: dict[str, str]) -> list[Dimension]:
        return [d for d in self.dimensions if _matches(d.applies_to, selections)]


class SemanticModel(BaseModel):
    datasets: list[Dataset] = Field(default_factory=list)

    def dataset(self, name: str) -> Dataset | None:
        return next((d for d in self.datasets if d.name == name), None)

    @classmethod
    def from_yaml(cls, path: str | Path) -> "SemanticModel":
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        return cls(**data)
