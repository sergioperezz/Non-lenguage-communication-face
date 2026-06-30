# BQ Excel Charts

Herramienta para construir **gráficos y tablas en Excel** a partir de datos de
**BigQuery**, sin escribir SQL, y copiarlos a **PowerPoint**. Pensada para una
gestora de activos (informes de fondos/carteras).

## Dos enfoques explorados

| | **A. Excel nativo** ⭐ (elegido) | **B. Office Add-in + FastAPI** (alternativa) |
|---|---|---|
| Dónde vive | En Excel: desplegables + macro VBA | Panel lateral (add-in) + backend |
| Datos | Power Query con tu **driver ODBC** de BigQuery | Backend con capa semántica → BigQuery |
| Gobierno (banca) | **Baja fricción**: instalar un driver / habilitar macros | Alta: desplegar add-in en M365, hostear backend, SSO |
| Carpeta | [`excel-nativo/`](excel-nativo/) | [`backend/`](backend/) + [`addin/`](addin/) |

Para **CaixaBank AM** seguimos el **enfoque A** (menos obstáculos de IT/Seguridad).
El enfoque B se conserva como referencia por si IT prefiere esa vía.

## Fases (enfoque A, Excel nativo)

- **Fase 1 — hecho:** panel interactivo nativo (Fondo → Grupo → Métrica →
  Dimensión → Filtro → Periodo → Benchmark → Gráfico) con **datos de ejemplo**, y
  prototipo de **copia a PowerPoint como EMF**. Ver [`excel-nativo/`](excel-nativo/).
- **Fase 2:** datos **reales** de BigQuery vía Power Query + ODBC (modo "Generar
  con datos reales").
- **Fase 3:** automatizar el pegado a PowerPoint (plantilla corporativa, varias
  diapositivas, presets de casos).

## Estructura

```
excel-nativo/   Enfoque A: panel en Excel (xlsx + macro VBA + docs)  ← actual
backend/        Enfoque B: API FastAPI (capa semántica, SQL, BigQuery)
addin/          Enfoque B: Office Add-in (React + Office.js)
```

Empieza por **[`excel-nativo/README.md`](excel-nativo/README.md)**.

---

## Enfoque B — arranque rápido (solo si exploras la vía add-in)

### Backend (funciona en modo dry-run sin BigQuery)

```bash
cd backend
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000   # docs en http://localhost:8000/docs
pytest                                       # tests del generador de SQL
```

Configuración en `backend/.env` (prefijo `BQEC_`): `BQEC_GCP_PROJECT` (vacío =
dry-run), `BQEC_GOOGLE_APPLICATION_CREDENTIALS`, `BQEC_SEMANTIC_MODEL_PATH`. El
modelo semántico (`backend/app/semantic/example_model.yaml`) define datasets,
dimensiones y métricas; las credenciales viven **solo en el backend**.

### Add-in

```bash
cd addin
npm install
npm run dev        # https://localhost:3000 (Office exige HTTPS)
```

Luego "sideload" de `addin/manifest.xml` en Excel (genera antes un GUID nuevo en
`<Id>`).

### Seguridad (enfoque B)

- Credenciales de BigQuery **solo en el backend**, nunca en el cliente.
- Valores de filtro como **parámetros de BigQuery** (anti-inyección).
- Solo se consultan campos definidos en la capa semántica (lista blanca).
