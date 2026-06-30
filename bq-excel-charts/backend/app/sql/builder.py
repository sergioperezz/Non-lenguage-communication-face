"""Generador de SQL de BigQuery a partir de un ChartSpec.

Diseño deliberado:
- Solo se referencian campos definidos en el modelo semántico (lista blanca).
  Si el usuario pide un campo desconocido, error claro — nunca se interpola.
- Los VALORES de los filtros van como parámetros nombrados de BigQuery
  (@p0, @p1...), no concatenados en el texto → sin inyección SQL.
- Las selecciones en cascada (selectors) se aplican como filtros y, además,
  validan que las dimensiones/métricas elegidas sean aplicables (applies_to).
"""

from __future__ import annotations

from ..schemas import ChartSpec, CompiledQuery, Filter
from ..semantic.model import Dataset, SemanticModel, _matches

_SQL_OP = {"=": "=", "!=": "!=", ">": ">", ">=": ">=", "<": "<", "<=": "<="}


class SpecError(ValueError):
    """El ChartSpec referencia algo que no existe o no es aplicable."""


def _field_sql(ds: Dataset, name: str) -> tuple[str, str]:
    """Devuelve (expresión_sql, tipo) de un campo filtrable (selector o dimensión)."""
    field = ds.filterable(name)
    if field is None:
        raise SpecError(f"Campo de filtro desconocido: '{name}'")
    return field.sql, field.type


def _wrap_value(col: str, type_: str, param: str) -> str:
    """Para columnas de fecha, parsea el parámetro STRING como DATE."""
    if type_ == "date":
        return f"{col} {{op}} DATE(@{param})"
    return f"{col} {{op}} @{param}"


def _resolve_filter(ds: Dataset, f: Filter, param: str) -> tuple[str, dict]:
    col, type_ = _field_sql(ds, f.field)

    if f.op in _SQL_OP:
        template = _wrap_value(col, type_, param).format(op=_SQL_OP[f.op])
        return template, {param: f.value}
    if f.op == "contains":
        return f"{col} LIKE CONCAT('%', @{param}, '%')", {param: f.value}
    if f.op in ("in", "not_in"):
        if not isinstance(f.value, (list, tuple)) or not f.value:
            raise SpecError(f"El operador '{f.op}' requiere una lista no vacía")
        kw = "NOT IN" if f.op == "not_in" else "IN"
        return f"{col} {kw} UNNEST(@{param})", {param: list(f.value)}

    raise SpecError(f"Operador no soportado: '{f.op}'")


def _validate_selections(ds: Dataset, selections: dict[str, str]) -> None:
    for name, value in selections.items():
        sel = ds.selector(name)
        if sel is None:
            raise SpecError(f"Selector desconocido: '{name}'")
        if sel.options and value not in {o.value for o in sel.options}:
            raise SpecError(f"Valor inválido para '{name}': '{value}'")


def build_query(model: SemanticModel, spec: ChartSpec) -> CompiledQuery:
    ds = model.dataset(spec.dataset)
    if ds is None:
        raise SpecError(f"Dataset desconocido: '{spec.dataset}'")
    if not spec.measures and not spec.dimensions:
        raise SpecError("Selecciona al menos una dimensión o una métrica")

    _validate_selections(ds, spec.selections)

    select_parts: list[str] = []
    group_parts: list[str] = []

    for dim_name in spec.dimensions:
        dim = ds.dimension(dim_name)
        if dim is None:
            raise SpecError(f"Dimensión desconocida: '{dim_name}'")
        if not _matches(dim.applies_to, spec.selections):
            raise SpecError(f"La dimensión '{dim_name}' no aplica a la selección actual")
        select_parts.append(f"{dim.sql} AS {dim.name}")
        group_parts.append(dim.sql)

    for measure_name in spec.measures:
        m = ds.measure(measure_name)
        if m is None:
            raise SpecError(f"Métrica desconocida: '{measure_name}'")
        if not _matches(m.applies_to, spec.selections):
            raise SpecError(f"La métrica '{measure_name}' no aplica a la selección actual")
        select_parts.append(f"{m.sql} AS {m.name}")

    params: dict = {}
    where_parts: list[str] = []

    # 1) Selecciones en cascada → filtros de igualdad sobre la columna del selector.
    for i, (name, value) in enumerate(spec.selections.items()):
        sel = ds.selector(name)
        param = f"s{i}"
        where_parts.append(f"{sel.sql} = @{param}")  # sel no es None (validado arriba)
        params[param] = value

    # 2) Filtros explícitos (rango de fechas, etc.).
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
