import { useState } from "react";
import { requestGoogleAuthUrl } from "../services/api";

interface GoogleAuthButtonProps {
  onAuthStarted?: () => void;
}

const GoogleAuthButton = ({ onAuthStarted }: GoogleAuthButtonProps) => {
  const [isLoading, setIsLoading] = useState(false);

  const handleClick = async () => {
    setIsLoading(true);
    onAuthStarted?.();
    try {
      const url = await requestGoogleAuthUrl();
      window.location.href = url;
    } catch (error) {
      console.error("No se pudo iniciar la autenticación con Google", error);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <button
      className="google-auth-button"
      onClick={handleClick}
      disabled={isLoading}
      aria-label="Autenticarse con Google"
    >
      {isLoading ? "Conectando..." : "Continuar con Google"}
    </button>
  );
};

export default GoogleAuthButton;
