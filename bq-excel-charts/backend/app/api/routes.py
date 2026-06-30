"""Endpoints de la API."""

from __future__ import annotations

from fastapi import APIRouter, HTTPException

from ..bigquery.client import run_query
from ..core.config import settings
from ..schemas import ChartSpec, CompiledQuery, QueryResult
from ..semantic.model import SemanticModel
from ..sql.builder import SpecError, build_query

router = APIRouter(prefix="/api")

# El modelo semántico se carga una vez al arrancar.
_model = SemanticModel.from_yaml(settings.semantic_model_path)


@router.get("/health")
def health() -> dict:
    return {"status": "ok", "mode": "live" if settings.gcp_project else "dry-run"}


def _dim_json(dim) -> dict:
    return {"name": dim.name, "label": dim.label, "type": dim.type, "applies_to": dim.applies_to}


def _measure_json(m) -> dict:
    return {"name": m.name, "label": m.label, "format": m.format, "applies_to": m.applies_to}


@router.get("/datasets")
def list_datasets() -> dict:
    """Catálogo completo que consume la UI para construir el query builder.

    Incluye selectores (con sus opciones) y el `applies_to` de cada campo, de modo
    que la UI puede hacer la cascada en cliente sin más llamadas.
    """
    return {
        "datasets": [
            {
                "name": d.name,
                "label": d.label,
                "description": d.description,
                "selectors": [
                    {
                        "name": s.name,
                        "label": s.label,
                        "type": s.type,
                        "options": [{"value": o.value, "label": o.label} for o in s.options],
                    }
                    for s in d.selectors
                ],
                "dimensions": [_dim_json(dim) for dim in d.dimensions],
                "measures": [_measure_json(m) for m in d.measures],
            }
            for d in _model.datasets
        ]
    }


@router.post("/available-fields")
def available_fields(body: dict) -> dict:
    """Dada una selección actual, devuelve qué dimensiones/métricas son válidas.

    Alternativa server-side a calcular la cascada en el cliente (útil si en el
    futuro las opciones se vuelven dinámicas o dependen de permisos).
    """
    ds = _model.dataset(body.get("dataset", ""))
    if ds is None:
        raise HTTPException(status_code=404, detail="Dataset desconocido")
    selections = body.get("selections", {}) or {}
    return {
        "dimensions": [_dim_json(d) for d in ds.available_dimensions(selections)],
        "measures": [_measure_json(m) for m in ds.available_measures(selections)],
    }


@router.post("/compile", response_model=CompiledQuery)
def compile_spec(spec: ChartSpec) -> CompiledQuery:
    """Genera el SQL sin ejecutarlo (vista previa de la query)."""
    try:
        return build_query(_model, spec)
    except SpecError as e:
        raise HTTPException(status_code=422, detail=str(e)) from e


@router.post("/query", response_model=QueryResult)
def execute(spec: ChartSpec) -> QueryResult:
    """Compila y ejecuta: devuelve los datos para pintar el gráfico."""
    try:
        compiled = build_query(_model, spec)
    except SpecError as e:
        raise HTTPException(status_code=422, detail=str(e)) from e

    ds = _model.dataset(spec.dataset)
    return run_query(compiled, max_bytes_billed=ds.max_bytes_billed if ds else None)
