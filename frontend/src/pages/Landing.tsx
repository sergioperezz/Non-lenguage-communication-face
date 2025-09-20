import { useEffect } from "react";
import { useNavigate } from "react-router-dom";
import GoogleAuthButton from "../components/GoogleAuthButton";
import { useAuthStore } from "../hooks/useAuthStore";

const Landing = () => {
  const navigate = useNavigate();
  const { login, isAuthenticated } = useAuthStore();

  useEffect(() => {
    if (isAuthenticated) {
      navigate("/dashboard", { replace: true });
    }
  }, [isAuthenticated, navigate]);

  const handleMockLogin = () => {
    login(
      {
        id: "demo-user",
        name: "Demo Coach",
        email: "demo@example.com"
      },
      "demo-token"
    );
  };

  return (
    <main className="landing">
      <section className="hero">
        <h1>Entrenador de Comunicación No Verbal</h1>
        <p>
          Practica entrevistas, presentaciones y conversaciones sociales con métricas en
          tiempo real y recomendaciones personalizadas.
        </p>
        <div className="auth-options">
          <GoogleAuthButton />
          <button type="button" onClick={handleMockLogin} className="secondary">
            Probar demo sin cuenta
          </button>
        </div>
      </section>
      <section className="features">
        <article>
          <h2>Métricas multi-modal</h2>
          <p>Analizamos voz, entonación, pronunciación, gestos y contacto visual.</p>
        </article>
        <article>
          <h2>Escenarios personalizables</h2>
          <p>Ajusta avatar, entorno y panel de métricas según tus objetivos.</p>
        </article>
        <article>
          <h2>Recomendaciones accionables</h2>
          <p>Recibe ejercicios específicos para mejorar cada habilidad.</p>
        </article>
      </section>
    </main>
  );
};

export default Landing;
