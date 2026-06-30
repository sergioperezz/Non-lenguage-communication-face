// Corazón de la UX: constructor visual. El usuario elige dataset, dimensiones,
// métricas y tipo de gráfico — nunca escribe SQL. Cada cambio puede previsualizar
// el SQL generado y, al pulsar "Insertar", ejecuta y dibuja el gráfico en Excel.

import { useEffect, useMemo, useState } from "react";
import { api, type ChartSpec, type ChartType, type Dataset } from "../api";
import { insertChart } from "../office";

const CHART_TYPES: ChartType[] = ["column", "bar", "line", "area", "pie", "table"];

function toggle(list: string[], item: string): string[] {
  return list.includes(item) ? list.filter((x) => x !== item) : [...list, item];
}

export function QueryBuilder() {
  const [datasets, setDatasets] = useState<Dataset[]>([]);
  const [datasetName, setDatasetName] = useState<string>("");
  const [dimensions, setDimensions] = useState<string[]>([]);
  const [measures, setMeasures] = useState<string[]>([]);
  const [chartType, setChartType] = useState<ChartType>("column");
  const [sql, setSql] = useState<string>("");
  const [status, setStatus] = useState<string>("");
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    api.datasets().then((d) => {
      setDatasets(d);
      if (d[0]) setDatasetName(d[0].name);
    }).catch((e) => setStatus(`Error cargando datasets: ${e.message}`));
  }, []);

  const dataset = useMemo(
    () => datasets.find((d) => d.name === datasetName),
    [datasets, datasetName],
  );

  // Al cambiar de dataset, reiniciamos selección.
  useEffect(() => {
    setDimensions([]);
    setMeasures([]);
  }, [datasetName]);

  const spec: ChartSpec = {
    dataset: datasetName,
    dimensions,
    measures,
    filters: [],
    order_by: [],
    limit: 1000,
    chart_type: chartType,
  };

  const canRun = datasetName && (dimensions.length > 0 || measures.length > 0);

  async function preview() {
    try {
      setStatus("");
      const { sql } = await api.compile(spec);
      setSql(sql);
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
        <select
          value={datasetName}
          onChange={(e) => setDatasetName(e.target.value)}
          style={{ width: "100%" }}
        >
          {datasets.map((d) => (
            <option key={d.name} value={d.name}>{d.label}</option>
          ))}
        </select>
      </label>

      {dataset && (
        <>
          <fieldset>
            <legend><strong>Dimensiones</strong></legend>
            {dataset.dimensions.map((dim) => (
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
            {dataset.measures.map((m) => (
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

          <label>
            <strong>Tipo de gráfico</strong>
            <select
              value={chartType}
              onChange={(e) => setChartType(e.target.value as ChartType)}
              style={{ width: "100%" }}
            >
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
            <pre style={{ background: "#f4f4f4", padding: 8, fontSize: 12, overflowX: "auto" }}>
              {sql}
            </pre>
          )}
          {status && <div style={{ fontSize: 13 }}>{status}</div>}
        </>
      )}
    </div>
  );
}
