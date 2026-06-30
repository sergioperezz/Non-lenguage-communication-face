"""Contratos de la API entre el add-in y el backend."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field

# Operadores permitidos en filtros. Lista blanca: nada que venga del usuario
# se interpola directamente en el SQL salvo el nombre del operador validado aquí.
FilterOp = Literal["=", "!=", ">", ">=", "<", "<=", "in", "not_in", "contains"]
ChartType = Literal["column", "bar", "line", "area", "pie", "table"]


class Filter(BaseModel):
    field: str = Field(..., description="Nombre de una dimensión del modelo")
    op: FilterOp = "="
    value: Any


class OrderBy(BaseModel):
    field: str
    dir: Literal["asc", "desc"] = "asc"


class ChartSpec(BaseModel):
    """Lo que el usuario configura en la UI. Se traduce a SQL y a un gráfico."""

    dataset: str
    # Valores elegidos en los selectores en cascada, p. ej. {"asset_class": "fixed_income"}.
    # Se aplican como filtros y deciden qué dimensiones/métricas son válidas.
    selections: dict[str, str] = Field(default_factory=dict)
    dimensions: list[str] = Field(default_factory=list)
    measures: list[str] = Field(default_factory=list)
    filters: list[Filter] = Field(default_factory=list)
    order_by: list[OrderBy] = Field(default_factory=list)
    limit: int = Field(default=1000, ge=1, le=50000)
    chart_type: ChartType = "column"


class CompiledQuery(BaseModel):
    sql: str
    # Parámetros nombrados de BigQuery (@p0, @p1...) para ejecutar de forma segura.
    parameters: dict[str, Any] = Field(default_factory=dict)


class QueryResult(BaseModel):
    columns: list[str]
    rows: list[list[Any]]
    total_rows: int
    bytes_processed: int | None = None
    sql: str
