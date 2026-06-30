"""Cliente de BigQuery con modo dry-run.

Si no hay proyecto/credenciales configurados, devuelve datos simulados para que
el equipo pueda desarrollar y demostrar la UX sin tocar BigQuery todavía.
"""

from __future__ import annotations

from typing import Any

from ..core.config import settings
from ..schemas import CompiledQuery, QueryResult


def _to_bq_parameters(parameters: dict[str, Any]):
    """Convierte el dict de parámetros en Scalar/ArrayQueryParameter de BigQuery."""
    from google.cloud import bigquery  # import perezoso: solo si se ejecuta de verdad

    def _scalar_type(value: Any) -> str:
        if isinstance(value, bool):
            return "BOOL"
        if isinstance(value, int):
            return "INT64"
        if isinstance(value, float):
            return "FLOAT64"
        return "STRING"

    bq_params = []
    for name, value in parameters.items():
        if isinstance(value, list):
            elem_type = _scalar_type(value[0]) if value else "STRING"
            bq_params.append(bigquery.ArrayQueryParameter(name, elem_type, value))
        else:
            bq_params.append(bigquery.ScalarQueryParameter(name, _scalar_type(value), value))
    return bq_params


def _mock_result(query: CompiledQuery) -> QueryResult:
    """Resultado simulado para modo dry-run."""
    return QueryResult(
        columns=["dimension", "value"],
        rows=[["Ejemplo A", 1200], ["Ejemplo B", 980], ["Ejemplo C", 1530]],
        total_rows=3,
        bytes_processed=0,
        sql=query.sql,
    )


def run_query(query: CompiledQuery, max_bytes_billed: int | None = None) -> QueryResult:
    if not settings.gcp_project:
        return _mock_result(query)

    from google.cloud import bigquery

    client = bigquery.Client(project=settings.gcp_project)
    job_config = bigquery.QueryJobConfig(
        query_parameters=_to_bq_parameters(query.parameters),
        maximum_bytes_billed=max_bytes_billed,
    )
    job = client.query(query.sql, job_config=job_config)
    rows = job.result()
    columns = [field.name for field in rows.schema]
    data = [[row[c] for c in columns] for row in rows]
    return QueryResult(
        columns=columns,
        rows=data,
        total_rows=rows.total_rows or len(data),
        bytes_processed=job.total_bytes_processed,
        sql=query.sql,
    )
