import { describe, expect, it } from "vitest";
import { render, screen } from "@testing-library/react";
import { BrowserRouter } from "react-router-dom";
import App from "./App";
import { MetricsProvider } from "./hooks/useMetricsStream";
import { useAuthStore } from "./hooks/useAuthStore";

const renderApp = () =>
  render(
    <BrowserRouter>
      <MetricsProvider>
        <App />
      </MetricsProvider>
    </BrowserRouter>
  );

describe("App", () => {
  it("muestra la página de aterrizaje por defecto", () => {
    useAuthStore.getState().logout();
    renderApp();
    expect(screen.getByText(/Entrenador de Comunicación No Verbal/i)).toBeInTheDocument();
  });
});
