# Entrenador de Comunicación No Verbal

Suite local para practicar habilidades de comunicación no verbal combinando escenarios interactivos en 3D, captura multimedia y análisis multimodal.

## Visión

El proyecto ayuda a profesionales y estudiantes a mejorar entrevistas, presentaciones y conversaciones sociales. El frontend guía las sesiones con avatares y métricas en vivo, mientras que el backend procesa audio/vídeo para generar recomendaciones personalizadas.

## Estructura del repositorio

```
frontend/  # Aplicación React + Vite + Three.js
backend/   # API FastAPI + SQLite para análisis y persistencia
docs/      # Documentación técnica y guías de extensión
```

## Dependencias principales

### Frontend
- React 18 + Vite
- Three.js + @react-three/fiber para escenas 3D
- Zustand para estado global
- Axios para comunicación con la API
- Vitest + Testing Library para pruebas

### Backend
- FastAPI + Uvicorn
- SQLModel + SQLite para persistencia
- Pydantic v2 y pydantic-settings para validación
- Pytest para pruebas

## Requisitos previos

- Node.js >= 18
- npm >= 9
- Python >= 3.11
- Entorno virtual recomendado (`python -m venv .venv`)

## Instalación

### Frontend
```bash
cd frontend
npm install
```

### Backend
```bash
cd backend
python -m venv .venv
source .venv/bin/activate  # En Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

## Ejecución

### Backend (FastAPI)
```bash
cd backend
source .venv/bin/activate
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

### Frontend (Vite)
```bash
cd frontend
npm run dev
```

La SPA consumirá la API en `http://localhost:8000/api`. Ajusta las variables de entorno en `backend/.env` si necesitas cambiar puertos u orígenes permitidos.

## Pruebas

### Frontend
```bash
cd frontend
npm test
```

### Backend
```bash
cd backend
source .venv/bin/activate
pytest
```

## Lint y formateo

### Frontend
```bash
cd frontend
npm run lint
npm run format
```

### Backend
Usa `ruff` o `black` si deseas integrarlos (no incluidos por defecto). El código sigue pep8/black y se puede ampliar con un `pyproject.toml` personalizado.

## API principal

| Método | Endpoint | Descripción |
| --- | --- | --- |
| `GET` | `/api/health` | Estado del backend. |
| `GET` | `/api/auth/google/url` | Devuelve la URL simulada de autenticación con Google. |
| `POST` | `/api/auth/google/callback` | Intercambia un código simulado por token y perfil. |
| `GET` | `/api/sessions/current` | Devuelve las métricas actuales del escenario seleccionado. |
| `GET` | `/api/sessions/summary` | Último resumen generado. |
| `GET` | `/api/sessions/history` | Historial de métricas guardadas. |
| `POST` | `/api/sessions/upload` | Sube audio/vídeo, analiza y guarda métricas. |
| `POST` | `/api/sessions/register` | Registra métricas calculadas externamente. |

## Extensión futura

- Integrar modelos reales de prosodia (p. ej. librosa + modelos de voz) y análisis facial (MediaPipe o OpenVINO) en `backend/app/services`.
- Añadir escenarios adicionales ampliando `frontend/src/components/ScenarioSelector.tsx` y las plantillas en `backend/app/services/recommendations.py`.
- Expandir el panel de métricas con configuraciones personalizadas almacenadas por usuario.
- Automatizar despliegues con contenedores Docker y orquestadores locales.

## Mantenimiento

1. Actualiza dependencias con `npm update` (frontend) o `pip-tools`/`pip` (backend).
2. Ejecuta pruebas y linters antes de cada commit.
3. Documenta cambios relevantes en `docs/` y en este README si impactan la arquitectura o los flujos principales.
