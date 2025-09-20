import { Canvas } from "@react-three/fiber";
import { OrbitControls, PerspectiveCamera } from "@react-three/drei";
import { Suspense } from "react";

interface AvatarSceneViewerProps {
  avatar: string;
  environment: string;
  onAvatarChange: (avatar: string) => void;
  onEnvironmentChange: (environment: string) => void;
}

const AVATARS = ["Profesional", "Estudiante", "Casual"];
const ENVIRONMENTS = ["Oficina", "Auditorio", "Cafetería"];

const PlaceholderAvatar = () => (
  <mesh>
    <boxGeometry args={[1, 2, 1]} />
    <meshStandardMaterial color="#6c63ff" />
  </mesh>
);

const PlaceholderEnvironment = () => (
  <mesh rotation={[-Math.PI / 2, 0, 0]} position={[0, -1, 0]}>
    <planeGeometry args={[20, 20]} />
    <meshStandardMaterial color="#f5f5f5" />
  </mesh>
);

const AvatarSceneViewer = ({
  avatar,
  environment,
  onAvatarChange,
  onEnvironmentChange
}: AvatarSceneViewerProps) => (
  <section className="card">
    <header>
      <h2>Avatar y escenario 3D</h2>
      <p>Personaliza la práctica seleccionando un avatar y entorno virtual.</p>
    </header>
    <div className="avatar-layout">
      <div className="avatar-selectors">
        <label htmlFor="avatar-select">Avatar</label>
        <select
          id="avatar-select"
          value={avatar}
          onChange={(event) => onAvatarChange(event.target.value)}
        >
          {AVATARS.map((option) => (
            <option key={option}>{option}</option>
          ))}
        </select>
        <label htmlFor="environment-select">Escenario</label>
        <select
          id="environment-select"
          value={environment}
          onChange={(event) => onEnvironmentChange(event.target.value)}
        >
          {ENVIRONMENTS.map((option) => (
            <option key={option}>{option}</option>
          ))}
        </select>
      </div>
      <div className="viewer">
        <Canvas style={{ height: 320 }} shadows>
          <ambientLight intensity={0.5} />
          <spotLight position={[5, 10, 5]} angle={0.2} penumbra={0.5} castShadow />
          <Suspense fallback={null}>
            <PerspectiveCamera makeDefault position={[4, 2, 4]} />
            <PlaceholderEnvironment />
            <PlaceholderAvatar />
          </Suspense>
          <OrbitControls enablePan enableZoom enableRotate />
        </Canvas>
      </div>
    </div>
  </section>
);

export default AvatarSceneViewer;
