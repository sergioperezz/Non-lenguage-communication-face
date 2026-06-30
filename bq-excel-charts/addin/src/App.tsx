import { QueryBuilder } from "./components/QueryBuilder";

export function App() {
  return (
    <main style={{ fontFamily: "Segoe UI, sans-serif", padding: 16 }}>
      <h2 style={{ marginTop: 0 }}>BQ Charts</h2>
      <p style={{ fontSize: 13, color: "#555" }}>
        Construye gráficos desde BigQuery sin escribir SQL.
      </p>
      <QueryBuilder />
    </main>
  );
}
