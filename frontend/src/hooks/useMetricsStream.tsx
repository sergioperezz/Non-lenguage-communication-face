import {
  createContext,
  PropsWithChildren,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState
} from "react";
import { fetchMetrics as fetchMetricsApi, fetchSummary as fetchSummaryApi } from "../services/api";
import { SessionMetrics, SessionSummary, Scenario } from "../types/metrics";
import { useAuthStore } from "./useAuthStore";

interface MetricsContextValue {
  currentMetrics: SessionMetrics | null;
  summary: SessionSummary | null;
  isLoading: boolean;
  refreshSummary: () => Promise<void>;
  setScenario: (scenario: Scenario) => void;
  scenario: Scenario;
}

const MetricsContext = createContext<MetricsContextValue | undefined>(undefined);

const DEFAULT_METRICS: SessionMetrics = {
  scenario: "interview",
  voice: { label: "Claridad de voz", value: 0.75 },
  prosody: { label: "Entonación", value: 0.68 },
  pronunciation: { label: "Pronunciación", value: 0.7 },
  speed: { label: "Velocidad", value: 0.62 },
  facialExpression: { label: "Expresión facial", value: 0.8 },
  bodyLanguage: { label: "Lenguaje corporal", value: 0.65 },
  fillerWords: { label: "Muletillas", value: 0.55 },
  overall: { label: "Puntaje global", value: 0.7 },
  timestamp: new Date().toISOString()
};

const DEFAULT_SUMMARY: SessionSummary = {
  generalComment:
    "Buen desempeño general. Trabaja en modular la entonación y reducir muletillas.",
  recommendations: [
    {
      title: "Ejercicios de prosodia",
      description: "Practica variaciones de tono en lecturas en voz alta.",
      suggestedExercises: [
        "Lectura con cambios de emoción",
        "Ejercicio de preguntas con distintas intenciones"
      ]
    },
    {
      title: "Control de muletillas",
      description: "Haz pausas conscientes antes de responder.",
      suggestedExercises: [
        "Responder preguntas con silencios de 2 segundos",
        "Grabar respuestas y autoevaluarlas"
      ]
    }
  ]
};

export function MetricsProvider({ children }: PropsWithChildren) {
  const { token } = useAuthStore();
  const [currentMetrics, setCurrentMetrics] = useState<SessionMetrics | null>(
    DEFAULT_METRICS
  );
  const [summary, setSummary] = useState<SessionSummary | null>(DEFAULT_SUMMARY);
  const [isLoading, setIsLoading] = useState(false);
  const [scenario, setScenarioState] = useState<Scenario>("interview");
  const intervalRef = useRef<number | null>(null);

  const fetchMetrics = async (activeScenario: Scenario) => {
    if (!token) {
      setCurrentMetrics((prev) => ({
        ...(prev ?? DEFAULT_METRICS),
        scenario: activeScenario,
        timestamp: new Date().toISOString()
      }));
      return;
    }

    try {
      const response = await fetchMetricsApi(activeScenario, token);
      setCurrentMetrics(response);
    } catch (error) {
      console.warn("Falling back to default metrics", error);
      setCurrentMetrics((prev) => ({
        ...(prev ?? DEFAULT_METRICS),
        scenario: activeScenario,
        timestamp: new Date().toISOString()
      }));
    }
  };

  const refreshSummary = async () => {
    if (!token) {
      setSummary(DEFAULT_SUMMARY);
      return;
    }
    try {
      setIsLoading(true);
      const response = await fetchSummaryApi(token);
      setSummary(response);
    } catch (error) {
      console.warn("Falling back to default summary", error);
      setSummary(DEFAULT_SUMMARY);
    } finally {
      setIsLoading(false);
    }
  };

  const setScenario = (newScenario: Scenario) => {
    setScenarioState(newScenario);
    void fetchMetrics(newScenario);
  };

  useEffect(() => {
    intervalRef.current = window.setInterval(() => {
      void fetchMetrics(scenario);
    }, 5000);
    return () => {
      if (intervalRef.current) {
        window.clearInterval(intervalRef.current);
      }
    };
  }, [scenario, token]);

  useEffect(() => {
    void fetchMetrics(scenario);
  }, [scenario, token]);

  const value = useMemo(
    () => ({ currentMetrics, summary, isLoading, refreshSummary, setScenario, scenario }),
    [currentMetrics, summary, isLoading, scenario]
  );

  return <MetricsContext.Provider value={value}>{children}</MetricsContext.Provider>;
}

export const useMetrics = (): MetricsContextValue => {
  const context = useContext(MetricsContext);
  if (!context) {
    throw new Error("useMetrics must be used within a MetricsProvider");
  }
  return context;
};
