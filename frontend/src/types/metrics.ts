export type Scenario = "interview" | "presentation" | "date";

export interface MetricScore {
  label: string;
  value: number;
  trend?: "up" | "down" | "steady";
  description?: string;
}

export interface SessionMetrics {
  scenario: Scenario;
  voice: MetricScore;
  prosody: MetricScore;
  pronunciation: MetricScore;
  speed: MetricScore;
  facialExpression: MetricScore;
  bodyLanguage: MetricScore;
  fillerWords: MetricScore;
  overall: MetricScore;
  timestamp: string;
}

export interface Recommendation {
  title: string;
  description: string;
  suggestedExercises: string[];
}

export interface SessionSummary {
  generalComment: string;
  recommendations: Recommendation[];
}
