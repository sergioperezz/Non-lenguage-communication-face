# Arquitectura del entrenador de comunicación

## Visión general

El sistema se divide en dos capas principales:

- **Frontend (React + Vite + Three.js)**: interfaz progresiva centrada en la práctica de sesiones con captura multimedia local, selección de escenarios y visualización de métricas en tiempo real.
- **Backend (FastAPI + SQLite)**: API local que procesa métricas de voz y lenguaje corporal, almacena sesiones, gestiona autenticación simulada con Google y produce recomendaciones personalizadas.

## Componentes clave

| Componente | Descripción |
| --- | --- |
| `frontend/src/hooks/useMetricsStream.tsx` | Orquesta la obtención periódica de métricas desde la API, administra el escenario activo y expone el resumen post-sesión. |
| `frontend/src/components/*` | Bloques reutilizables para autenticación, selección de escenarios, visualización 3D con Three.js y control de grabación multimedia. |
| `backend/app/services/metrics.py` | Fusiona el análisis de prosodia y lenguaje corporal para generar métricas normalizadas. |
| `backend/app/services/recommendations.py` | Construye comentarios y ejercicios adaptativos según escenario y métricas. |
| `backend/app/api/routes.py` | Expone endpoints REST para autenticación, sesiones, historial y carga de grabaciones. |
| `backend/app/core/database.py` | Inicializa la base de datos SQLite y gestiona sesiones SQLModel. |

## Extensibilidad

- **Nuevas métricas**: añadir funciones en `backend/app/services` y extender `SessionMetrics` en `backend/app/schemas.py`. El frontend consumirá los nuevos campos automáticamente al mapear el panel de métricas.
- **Escenarios adicionales**: actualizar las constantes de escenarios en el frontend y añadir plantillas en `FEEDBACK_TEMPLATES` dentro de `recommendations.py`.
- **Integración de modelos reales**: sustituir la lógica aleatoria en `prosody.py` y `facial.py` por llamadas a modelos externos (por ejemplo, PyTorch, MediaPipe) manteniendo la forma del retorno `MetricScore`.
- **Persistencia remota**: reemplazar SQLite por PostgreSQL actualizando la cadena de conexión en `Settings.database_url`.
