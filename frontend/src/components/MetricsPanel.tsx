import { useMemo } from "react";
import { SessionMetrics } from "../types/metrics";

interface MetricsPanelProps {
  metrics: SessionMetrics | null;
}

const formatPercentage = (value: number) => `${Math.round(value * 100)}%`;

const MetricsPanel = ({ metrics }: MetricsPanelProps) => {
  const entries = useMemo(() => {
    if (!metrics) return [];
    return [
      metrics.voice,
      metrics.prosody,
      metrics.pronunciation,
      metrics.speed,
      metrics.facialExpression,
      metrics.bodyLanguage,
      metrics.fillerWords,
      metrics.overall
    ];
  }, [metrics]);

  return (
    <section className="card">
      <header>
        <h2>Panel de métricas en tiempo real</h2>
        <p>Visualiza indicadores relevantes según el contexto seleccionado.</p>
      </header>
      {!metrics ? (
        <p>No hay métricas disponibles todavía.</p>
      ) : (
        <div className="metrics-grid">
          {entries.map((metric) => (
            <article key={metric.label} className="metric-card">
              <header>
                <h3>{metric.label}</h3>
                {metric.trend && <span className={`trend ${metric.trend}`}>{metric.trend}</span>}
              </header>
              <p className="value">{formatPercentage(metric.value)}</p>
              {metric.description && <p className="description">{metric.description}</p>}
            </article>
          ))}
        </div>
      )}
    </section>
  );
};

export default MetricsPanel;
