import { SessionSummary as SessionSummaryType } from "../types/metrics";

interface SessionSummaryProps {
  summary: SessionSummaryType | null;
  onRefresh: () => void;
  isLoading: boolean;
}

const SessionSummary = ({ summary, onRefresh, isLoading }: SessionSummaryProps) => (
  <section className="card">
    <header>
      <h2>Resumen post-sesión</h2>
      <button onClick={onRefresh} type="button" disabled={isLoading}>
        {isLoading ? "Actualizando..." : "Actualizar"}
      </button>
    </header>
    {!summary ? (
      <p>No hay resumen disponible todavía.</p>
    ) : (
      <div className="summary-content">
        <p>{summary.generalComment}</p>
        <ul>
          {summary.recommendations.map((recommendation) => (
            <li key={recommendation.title}>
              <h3>{recommendation.title}</h3>
              <p>{recommendation.description}</p>
              <ul>
                {recommendation.suggestedExercises.map((exercise) => (
                  <li key={exercise}>{exercise}</li>
                ))}
              </ul>
            </li>
          ))}
        </ul>
      </div>
    )}
  </section>
);

export default SessionSummary;
