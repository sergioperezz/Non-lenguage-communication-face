# BQ Excel Charts

Herramienta para construir **gráficos y tablas en Excel** a partir de datos de
**BigQuery**, sin escribir SQL, y (Fase 2) pegarlos automáticamente en
**PowerPoint**.

## Idea central

El usuario trabaja en un **panel lateral dentro de Excel** (Office Add-in).
Elige un dataset, arrastra dimensiones y métricas, escoge el tipo de gráfico y
pulsa "Insertar". Por debajo:

1. El add-in envía esa selección al **backend**.
2. El backend la traduce a **SQL de BigQuery** usando una **capa semántica**
   (el usuario nunca escribe SQL).
3. Ejecuta la consulta en BigQuery (las credenciales viven solo en el backend).
4. El add-in escribe los datos en una hoja y crea un **gráfico nativo de Excel**.

```
Office Add-in (React + Office.js)  ──HTTPS──►  Backend (FastAPI)  ──►  BigQuery
        panel en Excel                       capa semántica + SQL
```

## Por qué este diseño

- **Office Add-in**: mejor UX (dentro de Excel) y el mismo framework sirve para
  PowerPoint en la Fase 2.
- **Capa semántica**: el usuario no escribe SQL → mejor UX, y es la lista blanca
  que evita inyección SQL y controla costes de BigQuery.
- **Gráficos nativos** (no imágenes): editables y fáciles de pegar en PowerPoint.

## Estructura

```
backend/   API FastAPI: capa semántica, generador de SQL, cliente BigQuery
addin/     Office Add-in: React + Office.js (query builder)
```

## Arranque rápido

### Backend (funciona en modo dry-run sin BigQuery)

```bash
cd backend
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
# Sin BQEC_GCP_PROJECT → datos simulados. Doc interactiva en http://localhost:8000/docs
pytest    # ejecuta los tests del generador de SQL
```

### Add-in

```bash
cd addin
npm install
npm run dev        # sirve en https://localhost:3000 (Office exige HTTPS)
```

Luego haz "sideload" de `addin/manifest.xml` en Excel
(genera antes un GUID nuevo en `<Id>` y ajusta las URLs si cambian).

## Configuración

Copia `.env.example` a `backend/.env`. Variables (prefijo `BQEC_`):

| Variable | Descripción |
|---|---|
| `BQEC_GCP_PROJECT` | Proyecto de facturación. Vacío = modo dry-run. |
| `BQEC_GOOGLE_APPLICATION_CREDENTIALS` | Ruta al JSON de la service account. |
| `BQEC_SEMANTIC_MODEL_PATH` | Ruta al modelo semántico YAML. |

El modelo semántico (`backend/app/semantic/example_model.yaml`) define qué
datasets, dimensiones y métricas ve el usuario. Adáptalo a tus tablas reales.

### Selectores en cascada

El formulario se genera solo a partir del YAML. Para tener campos dependientes
(p. ej. eliges *Renta fija* y aparecen *Duración* / *TIR*) usas dos piezas:

```yaml
selectors:                      # desplegables de configuración
  - name: asset_class
    label: Tipo de activo
    column: asset_class
    options:
      - { value: fixed_income, label: Renta fija }
      - { value: equity, label: Renta variable }

measures:
  - name: duration
    label: Duración
    sql: AVG(duration)
    applies_to:                 # solo visible si asset_class = renta fija
      asset_class: [fixed_income]
```

- **`selectors`** → desplegables. Lo elegido se aplica como filtro (`WHERE`).
- **`applies_to`** en una métrica/dimensión → la UI solo la muestra cuando la
  selección encaja, y el backend rechaza combinaciones inválidas.
- Añadir un nuevo tipo de activo o métrica = editar el YAML, sin tocar código.

Las dimensiones de tipo `date` activan automáticamente un **rango de fechas**
en la UI (se traduce a `WHERE columna >= DATE(@p)`).

## Roadmap

- **Fase 1 (este esqueleto):** query builder + generación de SQL + gráficos en Excel.
- **Fase 2:** pegar los gráficos en PowerPoint (mismo add-in, host `Presentation`;
  copiar el gráfico como objeto/imagen a la diapositiva).
- **Mejoras futuras:** filtros en la UI, guardado de "plantillas" de gráfico,
  caché de consultas, OAuth por usuario, refresco de datos vinculados.

## Seguridad

- Las credenciales de BigQuery están **solo en el backend**, nunca en el cliente.
- Los valores de los filtros se pasan como **parámetros de BigQuery** (anti-inyección).
- Solo se pueden consultar campos definidos en la capa semántica (lista blanca).
