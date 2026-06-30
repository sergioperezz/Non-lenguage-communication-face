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


@router.get("/datasets")
def list_datasets() -> dict:
    """Catálogo que consume la UI para construir el query builder."""
    return {
        "datasets": [
            {
                "name": d.name,
                "label": d.label,
                "description": d.description,
                "dimensions": [
                    {"name": dim.name, "label": dim.label, "type": dim.type}
                    for dim in d.dimensions
                ],
                "measures": [
                    {"name": m.name, "label": m.label, "format": m.format}
                    for m in d.measures
                ],
            }
            for d in _model.datasets
        ]
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
