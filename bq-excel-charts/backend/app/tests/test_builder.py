from app.schemas import ChartSpec, Filter, OrderBy
from app.semantic.model import SemanticModel
from app.sql.builder import SpecError, build_query

MODEL = SemanticModel.from_yaml("app/semantic/example_model.yaml")


def test_basic_group_by():
    spec = ChartSpec(dataset="sales", dimensions=["region"], measures=["revenue"])
    q = build_query(MODEL, spec)
    assert "SUM(amount) AS revenue" in q.sql
    assert "region AS region" in q.sql
    assert "GROUP BY region" in q.sql
    assert "LIMIT 1000" in q.sql


def test_filters_are_parametrized():
    spec = ChartSpec(
        dataset="sales",
        dimensions=["region"],
        measures=["orders"],
        filters=[Filter(field="year", op="=", value=2025)],
    )
    q = build_query(MODEL, spec)
    assert "@p0" in q.sql
    assert q.parameters == {"p0": 2025}
    # El valor nunca aparece literal en el SQL (anti-inyección).
    assert "2025" not in q.sql


def test_in_operator_uses_unnest():
    spec = ChartSpec(
        dataset="sales",
        measures=["revenue"],
        filters=[Filter(field="region", op="in", value=["EU", "US"])],
    )
    q = build_query(MODEL, spec)
    assert "UNNEST(@p0)" in q.sql
    assert q.parameters["p0"] == ["EU", "US"]


def test_order_by_must_be_selected():
    spec = ChartSpec(
        dataset="sales",
        dimensions=["region"],
        order_by=[OrderBy(field="revenue", dir="desc")],
    )
    try:
        build_query(MODEL, spec)
        assert False, "debería fallar: revenue no está seleccionado"
    except SpecError:
        pass


def test_unknown_field_rejected():
    spec = ChartSpec(dataset="sales", dimensions=["nope"])
    try:
        build_query(MODEL, spec)
        assert False
    except SpecError:
        pass
