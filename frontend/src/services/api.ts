import axios from "axios";
import { SessionMetrics, SessionSummary, Scenario } from "../types/metrics";

const api = axios.create({
  baseURL: "/api"
});

export const requestGoogleAuthUrl = async () => {
  const response = await api.get<{ url: string }>("/auth/google/url");
  return response.data.url;
};

export const exchangeGoogleCode = async (code: string) => {
  const response = await api.post<{ token: string; user: { id: string; name: string; email: string; avatarUrl?: string } }>(
    "/auth/google/callback",
    { code }
  );
  return response.data;
};

export const fetchMetrics = async (
  scenario: Scenario,
  token: string
): Promise<SessionMetrics> => {
  const response = await api.get<SessionMetrics>(`/sessions/current?scenario=${scenario}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  return response.data;
};

export const fetchSummary = async (token: string): Promise<SessionSummary> => {
  const response = await api.get<SessionSummary>("/sessions/summary", {
    headers: { Authorization: `Bearer ${token}` }
  });
  return response.data;
};

export default api;
