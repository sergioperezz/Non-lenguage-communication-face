import { useEffect, useState } from "react";
import AvatarSceneViewer from "../components/AvatarSceneViewer";
import CaptureControls from "../components/CaptureControls";
import MetricsPanel from "../components/MetricsPanel";
import ScenarioSelector from "../components/ScenarioSelector";
import SessionSummary from "../components/SessionSummary";
import { useMetrics } from "../hooks/useMetricsStream";
import { useAuthStore } from "../hooks/useAuthStore";

const Dashboard = () => {
  const { currentMetrics, summary, refreshSummary, isLoading, setScenario, scenario } =
    useMetrics();
  const { user, logout } = useAuthStore();
  const [avatar, setAvatar] = useState("Profesional");
  const [environment, setEnvironment] = useState("Oficina");
  const [recordedBlob, setRecordedBlob] = useState<Blob | null>(null);

  useEffect(() => {
    if (recordedBlob) {
      console.info("Blob registrado listo para enviar", recordedBlob.size);
    }
  }, [recordedBlob]);

  return (
    <main className="dashboard">
      <header className="dashboard-header">
        <div>
          <h1>Hola, {user?.name ?? "Entrenador"}</h1>
          <p>Gestiona tu sesión y observa el progreso en tiempo real.</p>
        </div>
        <button onClick={logout} type="button" className="secondary">
          Cerrar sesión
        </button>
      </header>
      <div className="grid">
        <ScenarioSelector selected={scenario} onSelect={setScenario} />
        <AvatarSceneViewer
          avatar={avatar}
          environment={environment}
          onAvatarChange={setAvatar}
          onEnvironmentChange={setEnvironment}
        />
        <CaptureControls onRecordingFinished={setRecordedBlob} />
        <MetricsPanel metrics={currentMetrics} />
        <SessionSummary
          summary={summary}
          onRefresh={() => void refreshSummary()}
          isLoading={isLoading}
        />
      </div>
    </main>
  );
};

export default Dashboard;
