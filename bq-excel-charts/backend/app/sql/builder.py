"""Generador de SQL de BigQuery a partir de un ChartSpec.

Diseño deliberado:
- Solo se referencian campos definidos en el modelo semántico (lista blanca).
  Si el usuario pide un campo desconocido, error claro — nunca se interpola.
- Los VALORES de los filtros van como parámetros nombrados de BigQuery
  (@p0, @p1...), no concatenados en el texto → sin inyección SQL.
- Los NOMBRES de campo se resuelven a expresiones SQL controladas por nosotros.
"""

from __future__ import annotations

from ..schemas import ChartSpec, CompiledQuery, Filter
from ..semantic.model import Dataset, SemanticModel

# Traducción de operadores de la API a SQL.
_SQL_OP = {
    "=": "=",
    "!=": "!=",
    ">": ">",
    ">=": ">=",
    "<": "<",
    "<=": "<=",
}


class SpecError(ValueError):
    """El ChartSpec referencia algo que no existe en el modelo."""


def _resolve_filter(ds: Dataset, f: Filter, param_name: str) -> tuple[str, dict]:
    """Devuelve (fragmento_sql, {param: valor}) para un filtro."""
    dim = ds.dimension(f.field)
    if dim is None:
        raise SpecError(f"Campo de filtro desconocido: '{f.field}'")
    col = dim.sql

    if f.op in _SQL_OP:
        return f"{col} {_SQL_OP[f.op]} @{param_name}", {param_name: f.value}
    if f.op == "contains":
        # LIKE con comodines; el valor sigue siendo un parámetro.
        return f"{col} LIKE CONCAT('%', @{param_name}, '%')", {param_name: f.value}
    if f.op in ("in", "not_in"):
        if not isinstance(f.value, (list, tuple)) or not f.value:
            raise SpecError(f"El operador '{f.op}' requiere una lista no vacía")
        kw = "NOT IN" if f.op == "not_in" else "IN"
        # UNNEST de un array-parámetro: una sola variable, longitud arbitraria.
        return f"{col} {kw} UNNEST(@{param_name})", {param_name: list(f.value)}

    raise SpecError(f"Operador no soportado: '{f.op}'")


def build_query(model: SemanticModel, spec: ChartSpec) -> CompiledQuery:
    ds = model.dataset(spec.dataset)
    if ds is None:
        raise SpecError(f"Dataset desconocido: '{spec.dataset}'")
    if not spec.measures and not spec.dimensions:
        raise SpecError("Selecciona al menos una dimensión o una métrica")

    select_parts: list[str] = []
    group_parts: list[str] = []

    for dim_name in spec.dimensions:
        dim = ds.dimension(dim_name)
        if dim is None:
            raise SpecError(f"Dimensión desconocida: '{dim_name}'")
        select_parts.append(f"{dim.sql} AS {dim.name}")
        # Agrupamos por la expresión (no por el alias) para máxima compatibilidad.
        group_parts.append(dim.sql)

    for measure_name in spec.measures:
        m = ds.measure(measure_name)
        if m is None:
            raise SpecError(f"Métrica desconocida: '{measure_name}'")
        select_parts.append(f"{m.sql} AS {m.name}")

    params: dict = {}
    where_parts: list[str] = []
    for i, f in enumerate(spec.filters):
        fragment, p = _resolve_filter(ds, f, f"p{i}")
        where_parts.append(fragment)
        params.update(p)

    order_parts: list[str] = []
    valid_out = set(spec.dimensions) | set(spec.measures)
    for ob in spec.order_by:
        if ob.field not in valid_out:
            raise SpecError(f"No se puede ordenar por '{ob.field}': no está seleccionado")
        order_parts.append(f"{ob.field} {ob.dir.upper()}")

    sql = f"SELECT {', '.join(select_parts)}\nFROM `{ds.table}`"
    if where_parts:
        sql += "\nWHERE " + " AND ".join(where_parts)
    if group_parts:
        sql += "\nGROUP BY " + ", ".join(group_parts)
    if order_parts:
        sql += "\nORDER BY " + ", ".join(order_parts)
    sql += f"\nLIMIT {spec.limit}"

    return CompiledQuery(sql=sql, parameters=params)
