// Cliente del backend. El add-in nunca habla con BigQuery directamente:
// siempre pasa por el backend, que es quien tiene las credenciales.

const BASE = "http://localhost:8000/api";

// applies_to: {selector: [valoresPermitidos]}. Si está presente, el campo solo
// se muestra cuando la selección actual encaja.
export type AppliesTo = Record<string, string[]> | null;

export type Field = {
  name: string;
  label: string;
  type?: string;
  format?: string | null;
  applies_to?: AppliesTo;
};

export type SelectorOption = { value: string; label: string };
export type Selector = {
  name: string;
  label: string;
  type: string;
  options: SelectorOption[];
};

export type Dataset = {
  name: string;
  label: string;
  description?: string | null;
  selectors: Selector[];
  dimensions: Field[];
  measures: Field[];
};

// ¿Es aplicable un campo dada la selección actual? (cascada en cliente)
export function fieldApplies(field: Field, selections: Record<string, string>): boolean {
  if (!field.applies_to) return true;
  return Object.entries(field.applies_to).every(
    ([sel, allowed]) => selections[sel] !== undefined && allowed.includes(selections[sel]),
  );
}

export type ChartType = "column" | "bar" | "line" | "area" | "pie" | "table";

export type ChartSpec = {
  dataset: string;
  selections: Record<string, string>;
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
