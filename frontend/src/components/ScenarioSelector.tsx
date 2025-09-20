import { Scenario } from "../types/metrics";

interface ScenarioSelectorProps {
  selected: Scenario;
  onSelect: (scenario: Scenario) => void;
}

const SCENARIOS: { id: Scenario; label: string; description: string }[] = [
  {
    id: "interview",
    label: "Entrevista",
    description: "Simula entrevistas laborales con preguntas dinámicas."
  },
  {
    id: "presentation",
    label: "Presentación",
    description: "Entrena exposiciones con retroalimentación sobre ritmo y claridad."
  },
  {
    id: "date",
    label: "Cita",
    description: "Practica conversaciones informales cuidando lenguaje corporal."
  }
];

const ScenarioSelector = ({ selected, onSelect }: ScenarioSelectorProps) => (
  <section className="card">
    <header>
      <h2>Escenario</h2>
      <p>Elige el contexto para adaptar las métricas relevantes.</p>
    </header>
    <div className="scenario-grid">
      {SCENARIOS.map((scenario) => (
        <button
          key={scenario.id}
          onClick={() => onSelect(scenario.id)}
          className={`scenario-card ${selected === scenario.id ? "selected" : ""}`}
          type="button"
        >
          <h3>{scenario.label}</h3>
          <p>{scenario.description}</p>
        </button>
      ))}
    </div>
  </section>
);

export default ScenarioSelector;
