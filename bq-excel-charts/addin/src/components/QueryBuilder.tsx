// Corazón de la UX: constructor visual con SELECTORES EN CASCADA.
// El usuario elige p. ej. "Tipo de activo" → Renta fija, y solo entonces
// aparecen las métricas válidas (Duración, TIR...). Todo se deriva del modelo
// semántico (campo `applies_to`); aquí no hay reglas hardcodeadas.

import { useEffect, useMemo, useState } from "react";
import {
  api,
  fieldApplies,
  type ChartSpec,
  type ChartType,
  type Dataset,
  type Field,
} from "../api";
import { insertChart } from "../office";

const CHART_TYPES: ChartType[] = ["column", "bar", "line", "area", "pie", "table"];

function toggle(list: string[], item: string): string[] {
  return list.includes(item) ? list.filter((x) => x !== item) : [...list, item];
}

export function QueryBuilder() {
  const [datasets, setDatasets] = useState<Dataset[]>([]);
  const [datasetName, setDatasetName] = useState("");
  const [selections, setSelections] = useState<Record<string, string>>({});
  const [dimensions, setDimensions] = useState<string[]>([]);
  const [measures, setMeasures] = useState<string[]>([]);
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [chartType, setChartType] = useState<ChartType>("column");
  const [sql, setSql] = useState("");
  const [status, setStatus] = useState("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api
      .datasets()
      .then((d) => {
        setDatasets(d);
        if (d[0]) setDatasetName(d[0].name);
      })
      .catch((e) => setStatus(`Error cargando datasets: ${e.message}`));
  }, []);

  const dataset = useMemo(
    () => datasets.find((d) => d.name === datasetName),
    [datasets, datasetName],
  );

  // Al cambiar de dataset, reiniciamos toda la configuración.
  useEffect(() => {
    setSelections({});
    setDimensions([]);
    setMeasures([]);
  }, [datasetName]);

  // Campos visibles según la selección actual (la cascada).
  const visibleDimensions: Field[] = useMemo(
    () => (dataset ? dataset.dimensions.filter((d) => fieldApplies(d, selections)) : []),
    [dataset, selections],
  );
  const visibleMeasures: Field[] = useMemo(
    () => (dataset ? dataset.measures.filter((m) => fieldApplies(m, selections)) : []),
    [dataset, selections],
  );

  // Si una selección oculta un campo ya elegido, lo quitamos.
  useEffect(() => {
    const okDim = new Set(visibleDimensions.map((d) => d.name));
    const okMea = new Set(visibleMeasures.map((m) => m.name));
    setDimensions((prev) => prev.filter((n) => okDim.has(n)));
    setMeasures((prev) => prev.filter((n) => okMea.has(n)));
  }, [visibleDimensions, visibleMeasures]);

  const dateField = useMemo(
    () => visibleDimensions.find((d) => d.type === "date"),
    [visibleDimensions],
  );

  function setSelection(name: string, value: string) {
    setSelections((prev) => {
      const next = { ...prev };
      if (value) next[name] = value;
      else delete next[name];
      return next;
    });
  }

  const spec: ChartSpec = useMemo(() => {
    const filters: ChartSpec["filters"] = [];
    if (dateField && dateFrom) filters.push({ field: dateField.name, op: ">=", value: dateFrom });
    if (dateField && dateTo) filters.push({ field: dateField.name, op: "<=", value: dateTo });
    return {
      dataset: datasetName,
      selections,
      dimensions,
      measures,
      filters,
      order_by: [],
      limit: 1000,
      chart_type: chartType,
    };
  }, [datasetName, selections, dimensions, measures, dateField, dateFrom, dateTo, chartType]);

  const canRun = datasetName && (dimensions.length > 0 || measures.length > 0);

  async function preview() {
    try {
      setStatus("");
      setSql((await api.compile(spec)).sql);
    } catch (e) {
      setStatus((e as Error).message);
    }
  }

  async function insert() {
    setBusy(true);
    setStatus("Ejecutando consulta…");
    try {
      const result = await api.query(spec);
      await insertChart(result, chartType, dataset?.label ?? "Gráfico");
      setStatus(`Insertado · ${result.total_rows} filas`);
    } catch (e) {
      setStatus(`Error: ${(e as Error).message}`);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
      <label>
        <strong>Dataset</strong>
        <select value={datasetName} onChange={(e) => setDatasetName(e.target.value)} style={{ width: "100%" }}>
          {datasets.map((d) => (
            <option key={d.name} value={d.name}>{d.label}</option>
          ))}
        </select>
      </label>

      {dataset && (
        <>
          {/* Selectores en cascada (p. ej. Tipo de activo) */}
          {dataset.selectors.map((sel) => (
            <label key={sel.name}>
              <strong>{sel.label}</strong>
              <select
                value={selections[sel.name] ?? ""}
                onChange={(e) => setSelection(sel.name, e.target.value)}
                style={{ width: "100%" }}
              >
                <option value="">(todos)</option>
                {sel.options.map((o) => (
                  <option key={o.value} value={o.value}>{o.label}</option>
                ))}
              </select>
            </label>
          ))}

          <fieldset>
            <legend><strong>Dimensiones</strong></legend>
            {visibleDimensions.map((dim) => (
              <label key={dim.name} style={{ display: "block" }}>
                <input
                  type="checkbox"
                  checked={dimensions.includes(dim.name)}
                  onChange={() => setDimensions((d) => toggle(d, dim.name))}
                />
                {dim.label}
              </label>
            ))}
          </fieldset>

          <fieldset>
            <legend><strong>Métricas</strong></legend>
            {visibleMeasures.map((m) => (
              <label key={m.name} style={{ display: "block" }}>
                <input
                  type="checkbox"
                  checked={measures.includes(m.name)}
                  onChange={() => setMeasures((s) => toggle(s, m.name))}
                />
                {m.label}
              </label>
            ))}
          </fieldset>

          {/* Rango de fechas (solo si el dataset tiene una dimensión de fecha) */}
          {dateField && (
            <fieldset>
              <legend><strong>Fechas ({dateField.label})</strong></legend>
              <label>Desde <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} /></label>
              <label style={{ marginLeft: 8 }}>Hasta <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} /></label>
            </fieldset>
          )}

          <label>
            <strong>Tipo de gráfico</strong>
            <select value={chartType} onChange={(e) => setChartType(e.target.value as ChartType)} style={{ width: "100%" }}>
              {CHART_TYPES.map((t) => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>
          </label>

          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={preview} disabled={!canRun}>Ver SQL</button>
            <button onClick={insert} disabled={!canRun || busy}>
              {busy ? "Insertando…" : "Insertar en Excel"}
            </button>
          </div>

          {sql && (
            <pre style={{ background: "#f4f4f4", padding: 8, fontSize: 12, overflowX: "auto" }}>{sql}</pre>
          )}
          {status && <div style={{ fontSize: 13 }}>{status}</div>}
        </>
      )}
    </div>
  );
}
