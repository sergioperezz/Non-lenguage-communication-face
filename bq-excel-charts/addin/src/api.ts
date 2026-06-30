// Cliente del backend. El add-in nunca habla con BigQuery directamente:
// siempre pasa por el backend, que es quien tiene las credenciales.

const BASE = "http://localhost:8000/api";

export type Field = { name: string; label: string; type?: string; format?: string | null };
export type Dataset = {
  name: string;
  label: string;
  description?: string | null;
  dimensions: Field[];
  measures: Field[];
};

export type ChartType = "column" | "bar" | "line" | "area" | "pie" | "table";

export type ChartSpec = {
  dataset: string;
  dimensions: string[];
  measures: string[];
  filters: { field: string; op: string; value: unknown }[];
  order_by: { field: string; dir: "asc" | "desc" }[];
  limit: number;
  chart_type: ChartType;
};

export type QueryResult = {
  columns: string[];
  rows: unknown[][];
  total_rows: number;
  bytes_processed: number | null;
  sql: string;
};

async function post<T>(path: string, body: unknown): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const detail = await res.json().catch(() => ({}));
    throw new Error(detail.detail ?? `Error ${res.status}`);
  }
  return res.json() as Promise<T>;
}

export const api = {
  datasets: async (): Promise<Dataset[]> => {
    const res = await fetch(`${BASE}/datasets`);
    return (await res.json()).datasets;
  },
  compile: (spec: ChartSpec) => post<{ sql: string }>("/compile", spec),
  query: (spec: ChartSpec) => post<QueryResult>("/query", spec),
};
