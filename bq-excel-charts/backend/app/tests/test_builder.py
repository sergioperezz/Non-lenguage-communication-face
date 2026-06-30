import pytest

from app.schemas import ChartSpec, Filter, OrderBy
from app.semantic.model import SemanticModel
from app.sql.builder import SpecError, build_query

MODEL = SemanticModel.from_yaml("app/semantic/example_model.yaml")


def test_basic_group_by():
    spec = ChartSpec(dataset="portfolio", dimensions=["sector"], measures=["market_value"])
    q = build_query(MODEL, spec)
    assert "SUM(market_value) AS market_value" in q.sql
    assert "sector AS sector" in q.sql
    assert "GROUP BY sector" in q.sql
    assert "LIMIT 1000" in q.sql


def test_selection_becomes_filter():
    spec = ChartSpec(
        dataset="portfolio",
        selections={"asset_class": "fixed_income"},
        dimensions=["sector"],
        measures=["duration"],
    )
    q = build_query(MODEL, spec)
    assert "asset_class = @s0" in q.sql
    assert q.parameters["s0"] == "fixed_income"


def test_measure_requires_matching_selection():
    # "duration" solo aplica a renta fija; sin esa selección debe fallar.
    spec = ChartSpec(dataset="portfolio", dimensions=["sector"], measures=["duration"])
    with pytest.raises(SpecError):
        build_query(MODEL, spec)


def test_measure_not_applicable_for_wrong_class():
    spec = ChartSpec(
        dataset="portfolio",
        selections={"asset_class": "equity"},
        measures=["duration"],  # duración no aplica a renta variable
    )
    with pytest.raises(SpecError):
        build_query(MODEL, spec)


def test_invalid_selection_value_rejected():
    spec = ChartSpec(
        dataset="portfolio",
        selections={"asset_class": "crypto"},  # no está en las opciones
        measures=["market_value"],
    )
    with pytest.raises(SpecError):
        build_query(MODEL, spec)


def test_date_range_uses_date_cast():
    spec = ChartSpec(
        dataset="portfolio",
        dimensions=["date"],
        measures=["market_value"],
        filters=[
            Filter(field="date", op=">=", value="2025-01-01"),
            Filter(field="date", op="<=", value="2025-12-31"),
        ],
    )
    q = build_query(MODEL, spec)
    assert "as_of_date >= DATE(@p0)" in q.sql
    assert "as_of_date <= DATE(@p1)" in q.sql
    assert q.parameters == {"p0": "2025-01-01", "p1": "2025-12-31"}
    # Anti-inyección: las fechas no aparecen literales en el SQL.
    assert "2025-01-01" not in q.sql.replace("@p0", "")


def test_available_measures_cascade():
    ds = MODEL.dataset("portfolio")
    fi = {m.name for m in ds.available_measures({"asset_class": "fixed_income"})}
    eq = {m.name for m in ds.available_measures({"asset_class": "equity"})}
    assert "duration" in fi and "duration" not in eq
    assert "pe_ratio" in eq and "pe_ratio" not in fi
    assert "market_value" in fi and "market_value" in eq  # siempre disponible


def test_order_by_must_be_selected():
    spec = ChartSpec(
        dataset="portfolio",
        dimensions=["sector"],
        order_by=[OrderBy(field="market_value", dir="desc")],
    )
    with pytest.raises(SpecError):
        build_query(MODEL, spec)
