# Panel interactivo en Excel (Fase 1)

Gráfico que **se actualiza solo** cuando el usuario cambia tres parámetros, sin
backend ni add-in. Solo Excel (+ una macro pequeña para el tipo de gráfico).

## Archivos

| Archivo | Qué es |
|---|---|
| `Panel_Interactivo.xlsx` | El libro listo para usar (desplegables + gráfico) |
| `macro_Panel.vba` | Macro: cambia tipo de gráfico, cuida la cascada y **copia a PowerPoint como EMF** |
| `catalogo_graficos.md` | Qué gráficos usar en un AM, EMF a PowerPoint y notas de UX |
| `generar_panel.py` | Script que regenera el `.xlsx` (por si cambias los datos) |

Tipos de gráfico disponibles en el desplegable: **Columnas, Barras, Líneas,
Área, Circular, Anillo, Radar**. Los multi-serie (apiladas, combo, dispersión
riesgo/retorno, waterfall de atribución) se añaden en la Fase 2. Ver
`catalogo_graficos.md`.

---

## Versión genérica (recomendada): `Panel_Generico.xlsx`

En lugar de un caso fijo, este libro es un **motor guiado por datos**: cualquier
combinación se construye al vuelo combinando **5 parámetros**, sin pre-generar
gráficos. Es la forma de cubrir "todos los casos".

| Parámetro | Celda | Opciones |
|---|---|---|
| Tipo de activo | B3 | RF / RV |
| Métrica (cascada) | B4 | Rentabilidad, TIR, TER, Duración, Spread, Liquidez, Peso · (RV: PER, DividendYield) |
| **Dimensión / eje X** (cascada) | B5 | **Tiempo**: Mensual, Trimestral, Semestral, Anual · **Composición**: Activo, Geografia, Industria, Sector, Divisa, Rating |
| **Periodo** (ventana) | B6 | MTD, YTD, 3M, 6M, 1A, 3A |
| **Comparación** | B7 | Cartera + Benchmark, Solo cartera, Solo benchmark |
| Tipo de gráfico | B8 | Columnas, Barras, Líneas, Área, Circular, Anillo, Radar |

**Tu ejemplo** ("duración por trimestre, barras + benchmark línea") = `B3=RF`,
`B4=Duración`, `B5=Trimestral`, `B6=3A`, `B7=Cartera + Benchmark`, `B8=Columnas`.
La macro detecta "Columnas + ambas series" y dibuja el **benchmark como línea**
(combo).

**El Periodo (B6)** acota cuántos buckets recientes se muestran en las
dimensiones de tiempo (p. ej. `1A` + `Mensual` = últimos 12 meses); en las de
composición se ignora (muestra todas las categorías). Es una aproximación del
*mock*: en la Fase 2, el periodo es un filtro de fechas exacto en la query.

Cómo funciona por dentro:
- La hoja **Datos** está en formato largo:
  `TipoActivo | Metrica | EjeXTipo | EjeXValor | Serie | Valor`.
- La tabla de resultados (D2:F8) se calcula con **SUMIFS** según los 5 parámetros;
  las categorías del eje X salen de la dimensión elegida (`INDIRECT("Cat_"&B5)`).
- El gráfico está enlazado a esa tabla → se actualiza solo; la macro ajusta el
  tipo y el combo.

Instalación: igual que arriba, pero pegando `macro_Panel_v2.vba`.

Para añadir una dimensión nueva (p. ej. "Divisa") o una métrica: basta con
añadir filas en **Datos**, su lista en **Listas** y un rango con nombre. En la
Fase 2 todo esto lo aportará la consulta a BigQuery automáticamente.

### Cobertura completa de los datos de ejemplo

Los datos de ejemplo cubren **todas** las combinaciones seleccionables
(activo × métrica × dimensión × serie), así que en "modo diseño" **cualquier**
gráfico que elijas se ve al instante, sin huecos vacíos (valores ficticios pero
formato real). Incluye la métrica **Liquidez** y la dimensión **Mes** (36 meses,
3 años) para comprobar que escala. Las combinaciones imposibles (p. ej. Rating en
RV) no aparecen como opción gracias a la cascada.

> El "modo real" (botón *Generar con datos reales*) de la Fase 2 sustituirá esos
> valores ficticios por los de BigQuery para la combinación elegida.

## Cómo funciona

Tres hojas:
- **Panel**: lo que usa el analista. Tres desplegables y el gráfico.
- **Datos**: tabla de ejemplo en formato largo (`TipoActivo | Metrica | Segmento | Valor`).
  En la **Fase 2** esta hoja se sustituye por una consulta a BigQuery.
- **Listas**: las métricas de cada tipo de activo (para el desplegable en cascada).

El mecanismo:
1. El usuario elige **Tipo de activo** (B3), **Métrica** (B4) y **Tipo de gráfico** (B5).
2. La métrica es un **desplegable en cascada**: sus opciones dependen del tipo de
   activo (vía `INDIRECT` y los rangos con nombre `RF` / `RV`).
3. La tabla de resultados (D3:E8) usa **`SUMIFS`** para traer los valores del
   tipo+métrica elegidos.
4. El **gráfico está enlazado a esa tabla** → al cambiar un desplegable, las
   fórmulas recalculan y el gráfico **se redibuja automáticamente**.
5. La **macro** detecta el cambio y: ajusta la métrica si quedó inválida al
   cambiar de RF a RV, y cambia el gráfico entre **verticales/horizontales**.

## Puesta en marcha (una vez)

1. Abre `Panel_Interactivo.xlsx`.
2. Clic derecho en la pestaña **Panel** → **Ver código** → pega el contenido de
   `macro_Panel.vba`.
3. **Guardar como** → tipo *Libro de Excel habilitado para macros (.xlsm)*.
4. Al reabrir, **Habilitar macros**.

A partir de ahí: cambias un desplegable y el gráfico cambia solo. Cero fricción.

### Exportar a PowerPoint (EMF)

La macro incluye `CopiarAPowerPoint`: copia el gráfico como **EMF vectorial**
(nítido, sin vínculos que se rompan) y lo pega en PowerPoint. Para usarla con un
clic: Insertar → una Forma/botón → clic derecho → *Asignar macro* →
`CopiarAPowerPoint`. (Pega la PARTE B de la macro en un **módulo estándar**.)

## Cómo adaptarlo a tus datos / métricas

- **Pocos cambios**: edita la hoja **Datos** (y **Listas** si añades métricas).
- **Regenerar desde cero**: edita el diccionario `DATOS` en `generar_panel.py` y
  ejecútalo. (Tras regenerar hay que volver a pegar la macro y guardar como .xlsm.)

Métricas de ejemplo:
- **RF** (renta fija): Duration, TIR, Spread
- **RV** (renta variable): Rentabilidad, PER, DividendYield

## Puente a la Fase 2 (BigQuery)

El diseño está pensado para que la Fase 2 sea un cambio pequeño:
- Sustituyes la hoja **Datos** por una **consulta de Power Query** contra
  BigQuery (con tu driver ODBC), parametrizada por las celdas B3/B4.
- Descomentas la línea `ThisWorkbook.RefreshAll` de la macro para que, al cambiar
  un parámetro, se relance la consulta.
- **El gráfico y los desplegables no cambian**: siguen enlazados a la misma tabla.
