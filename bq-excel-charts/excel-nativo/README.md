# Panel de gráficos en Excel (Fase 1)

Motor de gráficos **en Excel nativo** (sin backend ni add-in; solo Excel + una
macro). El usuario combina parámetros y el gráfico **se actualiza solo**; un clic
lo copia a PowerPoint como **EMF**. Hoy con **datos de ejemplo**; en la Fase 2
esos datos vendrán de **BigQuery** (por tu driver ODBC).

## Archivos

| Archivo | Qué es |
|---|---|
| `Panel_Generico.xlsx` | El libro listo para usar (desplegables + gráfico) |
| `macro_Panel.vba` | Macro: cascadas, multi-entidad, tipo/estilo de gráfico, rango dinámico, ejes y **copia a PowerPoint (EMF)** |
| `generar_panel.py` | Script que regenera el `.xlsx` (si cambias datos/estructura) |
| `mapa_casos_uso.html` / `.md` | Mapa de casos de uso por familias |
| `catalogo_graficos.md` | Qué gráficos usar en un AM, EMF a PowerPoint y notas de UX |

## El modelo

| Parámetro | Celda | Opciones |
|---|---|---|
| **Tipo de entidad** | B3 | Fondo · Cartera · Indice |
| **Entidad 1** (cascada) | B4 | la entidad principal de ese tipo |
| **Entidad 2** (opcional) | B5 | otra entidad **para comparar** (vacía = no se usa) |
| **Entidad 3** (opcional) | B6 | otra entidad para comparar |
| **Grupo de métrica** | B7 | Rendimiento · Riesgo · Composicion · Costes · Liquidez · Valoracion |
| **Métrica** (cascada del grupo) | B8 | p. ej. Riesgo → Duración, TIR, Spread, Volatilidad, Beta |
| **Dimensión / eje X** | B9 | *Tiempo*: Mensual, Trimestral, Semestral, Anual · *Composición*: Activo, Geografia, Industria, Sector, Divisa, Rating |
| **Filtro: tipo de activo** | B10 | Todos · RF · RV |
| **Periodo** | B11 | MTD · YTD · 1M–6M · 1A–6A |
| **Benchmark** | B12 | Con benchmark · Sin benchmark (de la Entidad 1) |
| **Estilo benchmark** | B13 | Barras · Líneas · **Puntos** |
| **Tipo de gráfico** | B14 | Columnas · Barras · Líneas · Área · Circular · Anillo · Radar · **Apiladas** · **100% apiladas** |

> Nota: los tokens **Indice**, **Composicion**, **Valoracion**, **Geografia** van
> **sin tilde** en los desplegables (Excel los usa para nombres internos vía `INDIRECT`).

**Comparar varias entidades**: cada Entidad seleccionada (B4/B5/B6) es una serie
del gráfico. Deja B5/B6 vacías si no quieres comparar.

**Ejemplo (tu caso)** — *Rentabilidad de Cartera RF Gobierno vs Cartera RF Crédito*:
`B3=Cartera`, `B4=Cartera RF Gobierno`, `B5=Cartera RF Crédito`, `B7=Rendimiento`,
`B8=Rentabilidad`, `B9=Trimestral`, `B12=Sin benchmark`, `B14=Columnas`.

## Cómo funciona por dentro

El **Panel** se apoya en tres hojas (además, Tablas/Sectorial/Comparativa son
generadores aparte con sus propias hojas de datos):
- **Panel**: los desplegables + el gráfico.
- **Datos**: formato largo `Entidad | TipoActivo | Metrica | EjeXTipo | EjeXValor | Serie | Valor`.
- **Listas**: rangos con nombre para cascadas y categorías.

Mecanismo:
1. **Entidad** cascadea del tipo (B3→B4/B5/B6) vía `INDIRECT("Ent_"&B3)`; la
   **métrica** cascadea del grupo (B7→B8) vía `INDIRECT("Grupo_"&B7)`.
2. La tabla de resultados (D:H) se calcula con **`SUMIFS`**: D=categoría,
   E/F/G = una columna por entidad, H = benchmark de la Entidad 1.
3. El **filtro** (B10) usa comodín en `SUMIFS`; el **Periodo** (B11) acota buckets.
4. La **macro** enlaza una serie por entidad no vacía + el benchmark, con **rango
   dinámico** (solo categorías activas), aplica tipo/estilo, título y ejes.

## Puesta en marcha (una vez)

1. Abre `Panel_Generico.xlsx`.
2. Pega `macro_Panel.vba`: **Parte A** en la hoja *Panel* (clic derecho → *Ver
   código*) y **Parte B** en un *Módulo* (Insertar → Módulo).
3. **Guardar como** → *.xlsm* → **Habilitar macros**.

### Exportar a PowerPoint (EMF)

`CopiarAPowerPoint` copia el gráfico como **EMF vectorial** y lo pega en
PowerPoint. Un clic: Insertar → una Forma → *Asignar macro* → `CopiarAPowerPoint`.

## Cómo adaptarlo

- Edita la hoja **Datos**/**Listas**, o regenera con `generar_panel.py`
  (ENTIDADES, GRUPOS, DIMS, CATEGORIAS, PARAMS). Tras regenerar, repega la macro.

## Hoja "Tablas" (generador configurable — con macro)

La hoja **Tablas** genera tablas formateadas desde los mismos datos. Es
**configurable por macro** (`FormatearTablas`, PARTE C): al cambiar un desplegable
o abrir la pestaña se ajustan columnas, decimales y estilo. Selectores:

| Celda | Parámetro | Opciones |
|---|---|---|
| B3 | Tipo de entidad | Fondo · Cartera · Indice |
| B4 | Métrica | cualquiera |
| B5 | Dimensión (columnas) | tiempo o composición |
| B6 | Serie | Cartera · Benchmark |
| B7 | Filtro tipo de activo | Todos · RF · RV |
| B8 | Periodo | MTD · YTD · 1M–6M · 1A–6A |
| B9 | **Estilo** | Mapa de calor · Barras de datos · Signos +/- · Sin formato |
| B10 | **Decimales** | 0 · 1 · 2 |

- Filas = entidades del tipo; columnas = categorías de la dimensión (las que
  entren según el **periodo**; las vacías se **ocultan solas**).
- El **estilo** cambia el formato condicional (mapa de calor, barras de datos,
  verde/rojo por signo, o sin formato) y **Decimales** el número de decimales.
- Valores con `SUMIFS` (incluye filtro de tipo de activo). Cubre "Mapa de Calor
  RF" y "Comportamiento de mercados" de los decks.

## Hoja "Sectorial" (apilado + eje secundario)

Reproduce la **"Distribución sectorial"** de los decks: columnas **apiladas** por
componentes (Large/Mid/Small Cap), **más** una serie de **puntos** con la
**Contribución al Tracking Error** en un **eje secundario**. Se puede **agrupar
por** varias dimensiones: selector de **Entidad** (B3) y **Agrupar por** (B4:
Sector · Industria · Geografia · Divisa · Rating · Activo). La tabla se calcula
con `SUMIFS` (hoja `DatosSector`) y el combo está construido con openpyxl —
**sin macro**.

## Hoja "Comparativa" (eje X = carteras + columnas *diff*)

Reproduce el caso **"Comportamiento de las Carteras"** (DPM Global Managers) de
los decks: el **eje X es la lista de carteras** (no tiempo ni composición) y por
cada cartera hay **columnas agrupadas** de varias versiones (**Gestionada / IMP /
Benchmark**), más una tabla con columnas **"diff"** (cartera − benchmark) en
**verde/rojo** según el signo. Es un **corte transversal**: aquí las entidades son
las *categorías*, al revés que en el Panel (donde son *series*). Selectores: B3
(tipo entidad) y B4 (métrica); tabla con `SUMIFS` (hoja `DatosComp`) — **sin macro**.

## Pendiente de implementar (siguiente paso)

- **Dispersión (riesgo-retorno)**: necesita un modo con **dos métricas** (X e Y)
  y un punto por entidad. Es un modo aparte; se hará con pruebas en tu Excel.
- Nota: en el panel, **Apiladas / 100% apiladas** apilan las **series
  seleccionadas** (las entidades). La composición apilada por *componentes*
  (sub-industrias/caps) es la **hoja Sectorial** de arriba.

## Fase 2 (BigQuery por ODBC) — PARTE D de la macro

La **PARTE D** (módulo estándar nuevo) convierte los parámetros del Panel en una
consulta al **modelo homologado** (esquema en estrella real), la ejecuta por
**ODBC (ADODB)** y vuelca el resultado en la hoja **`BQ_Resultado`**.

**Mapeo (según el diccionario de datos homologado):**

| Métrica del Panel | Tabla | Columna |
|---|---|---|
| Rentabilidad / Rentab. acum. | `CAM_TX_PERFORMANCE_FIGURES_PD` | `TWR_<periodo>` (fondo) y `TWR_<periodo>_BMK` (benchmark) |
| Volatilidad | `CAM_TX_PERFORMANCE_FIGURES_PD` | `VOL_1Y_260` |
| Beta | `CAM_TX_PERFORMANCE_FIGURES_PD` | `BETA` |
| Duración / TIR / Spread | `CAM_TX_RISK_FIG_AGG_PD` | `PK_VARIABLE_TARGET` + `PK_CRITERIO_AGREGACION` *(pendiente)* |
| Peso / Composición | `CAM_TM_PORTFOLIOS_PD` / `CAM_TX_BENCHMARK_COMP_PD` | `COMPONENT`, `WEIGHT` *(pendiente)* |

- Identidad y tipo de la entidad: `CAM_TM_PORTFOLIOS_PD` (`PORTFOLIO_NAME`,
  `PTF_PORTFOLIO_TYPE`, `PK_PORTFOLIO_ID`, `PK_FECHA_DATOS`).
- **Periodo → columna** (no es un filtro de fecha): `MTD→MTD`, `YTD→YTD`, `1M→1M`,
  `1A→1Y`, `3A→3Y`, `5A→5Y`. Los periodos `2M/3M/4M/5M/6M` y `2A/4A/6A` **no
  tienen columna directa** en el DWH (habría que calcularlos de una serie).
- La consulta es de **corte transversal**: un valor por portfolio en la **última
  fecha** (`PK_FECHA_DATOS = MAX(...)`), fondo y benchmark vía `UNION ALL`.
- **Vista previa en vivo** en el Panel (celda **A41**) al cambiar un parámetro;
  **`VerSQL`** la muestra; **`RefrescarDatos`** conecta y trae los datos a
  `BQ_Resultado`.
- **CONFIG** (arriba de la PARTE D): `BQ_CONN` (`DSN=...` o `Driver={...}`) y
  `BQ_DATASET` (`proyecto.dataset`).

> **Pendiente de acordar** para cerrar el mapeo: (1) columna de valor de
> `CAM_TX_RISK_FIG_AGG_PD` y las etiquetas de `PK_CRITERIO_AGREGACION` para las
> dimensiones (sector/rating/…); (2) si la selección de entidad va por
> `PORTFOLIO_NAME` o por `PK_PORTFOLIO_ID`; (3) alinear la lista de periodos del
> Panel a los horizontes reales del DWH.

## Limitaciones del *mock* (se resuelven en la Fase 2)

- El filtro **Todos** en una entidad con RF y RV **suma** ambos (aproximación).
- Filtrar a un tipo de activo que la entidad no tiene muestra vacío (esperado).
- Solo hay **3 años** de datos, así que 3A–6A muestran lo mismo.
