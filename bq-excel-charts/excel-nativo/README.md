# Panel de gráficos en Excel (Fase 1)

Motor de gráficos **en Excel nativo** (sin backend ni add-in; solo Excel + una
macro). El usuario combina parámetros y el gráfico **se actualiza solo**; un clic
lo copia a PowerPoint como **EMF**. Hoy con **datos de ejemplo**; en la Fase 2
esos datos vendrán de **BigQuery** (por tu driver ODBC).

## Archivos

| Archivo | Qué es |
|---|---|
| `Panel_Generico.xlsx` | El libro listo para usar (desplegables + gráfico) |
| `macro_Panel.vba` | Macro: cascada métrica, tipo de gráfico, combo barras+línea y **copia a PowerPoint (EMF)** |
| `generar_panel.py` | Script que regenera el `.xlsx` (si cambias datos/estructura) |
| `mapa_casos_uso.html` / `.md` | Mapa de casos de uso por familias |
| `catalogo_graficos.md` | Qué gráficos usar en un AM, EMF a PowerPoint y notas de UX |

## El modelo: 8 parámetros que cubren cualquier caso

No se pre-genera un gráfico por caso: un **motor guiado por datos** combina los
parámetros y construye cualquier gráfico al vuelo.

| Parámetro | Celda | Opciones |
|---|---|---|
| **Fondo / Cartera** | B3 | RF Privada A · Bolsa Europa · Mixto Moderado |
| **Grupo de métrica** | B4 | Rendimiento · Riesgo · Composicion · Costes · Liquidez · Valoracion |
| **Métrica** (cascada del grupo) | B5 | p. ej. Riesgo → Duración, TIR, Spread, Volatilidad, Beta |
| **Dimensión / eje X** | B6 | *Tiempo*: Mensual, Trimestral, Semestral, Anual · *Composición*: Activo, Geografia, Industria, Sector, Divisa, Rating |
| **Filtro: tipo de activo** | B7 | Todos · RF · RV *(útil en fondos mixtos)* |
| **Periodo** | B8 | MTD · YTD · 3M · 6M · 1A · 3A |
| **Benchmark** | B9 | Con benchmark · Sin benchmark |
| **Tipo de gráfico** | B10 | Columnas · Barras · Líneas · Área · Circular · Anillo · Radar |

> Nota: los tokens **Composicion**, **Valoracion** y **Geografia** van **sin tilde** en
> los desplegables (Excel los usa para resolver nombres internos vía `INDIRECT`).

**Ejemplo** ("duración por trimestre, barras + benchmark en línea") = `B3=RF Privada A`,
`B4=Riesgo`, `B5=Duración`, `B6=Trimestral`, `B7=Todos`, `B8=3A`,
`B9=Con benchmark`, `B10=Columnas`. La macro detecta "Columnas + Con benchmark" y
dibuja el benchmark como **línea** (combo).

## Cómo funciona por dentro

Tres hojas:
- **Panel**: lo que usa el analista (los 8 desplegables + el gráfico).
- **Datos**: tabla en formato largo
  `Fondo | TipoActivo | Metrica | EjeXTipo | EjeXValor | Serie | Valor`.
  En la **Fase 2** esta hoja la rellena la consulta a BigQuery.
- **Listas**: rangos con nombre para las cascadas y categorías.

Mecanismo:
1. La **métrica cascadea del grupo** (B4→B5) vía `INDIRECT("Grupo_"&B4)` → ningún
   desplegable kilométrico.
2. La tabla de resultados (D:F) se calcula con **`SUMIFS`** según los parámetros;
   las categorías del eje X salen de la dimensión elegida (`INDIRECT("Cat_"&B6)`).
3. El **filtro de tipo de activo** (B7) usa un comodín en `SUMIFS` (`Todos`=`*`).
4. El **Periodo** (B8) acota cuántos buckets recientes se muestran en dimensiones
   de tiempo (tabla Periodo×Granularidad); en composición se ignora.
5. El **gráfico está enlazado a la tabla** → al cambiar un desplegable, recalcula
   y se redibuja solo. La **macro** ajusta el tipo de gráfico y el combo.

## Puesta en marcha (una vez)

1. Abre `Panel_Generico.xlsx`.
2. Pega `macro_Panel.vba`: **Parte A** en la hoja *Panel* (clic derecho en la
   pestaña → *Ver código*) y **Parte B** en un *Módulo* estándar (Insertar → Módulo).
3. **Guardar como** → *.xlsm* (libro habilitado para macros).
4. Al reabrir, **Habilitar macros**.

Cambias un desplegable y el gráfico cambia solo.

### Exportar a PowerPoint (EMF)

La macro `CopiarAPowerPoint` copia el gráfico como **EMF vectorial** (nítido, sin
vínculos que se rompan) y lo pega en PowerPoint. Para un clic: Insertar → una
Forma/botón → *Asignar macro* → `CopiarAPowerPoint`.

## Cobertura completa de los datos de ejemplo

Los datos de ejemplo cubren **todas** las combinaciones seleccionables, así que en
modo diseño **cualquier** gráfico que elijas se ve al instante, sin huecos
(valores ficticios, formato real). En la Fase 2, el "modo real" sustituirá esos
valores por los de BigQuery para la combinación elegida.

## Cómo adaptarlo

- **Pocos cambios**: edita la hoja **Datos** (y **Listas** si añades métricas).
- **Regenerar**: edita los diccionarios de `generar_panel.py` (FONDOS, GRUPOS,
  DIMS, CATEGORIAS, PARAMS) y ejecútalo. Tras regenerar, vuelve a pegar la macro y
  guarda como `.xlsm`.

## Puente a la Fase 2 (BigQuery)

Diseñado para que la Fase 2 sea un cambio pequeño:
- Sustituyes la hoja **Datos** por una **consulta de Power Query** contra BigQuery
  (driver ODBC), con las mismas columnas y parametrizada por las celdas del panel.
- Descomentas `ThisWorkbook.RefreshAll` en la macro para refrescar al cambiar el
  ámbito de datos (fondo/periodo).
- **El gráfico y los desplegables no cambian**: siguen enlazados a la misma tabla.

## Limitaciones del *mock* (se resuelven en la Fase 2)

- En un fondo **mixto** con filtro **Todos**, los valores se **suman** entre RF y
  RV (simplificación); la agregación correcta (media ponderada) es de la Fase 2.
- El **Periodo** es una aproximación (buckets recientes); en la Fase 2 será un
  filtro de fechas exacto en la query.
- Todas las dimensiones aparecen para todos los fondos (cobertura del *mock*); con
  datos reales solo existirán las combinaciones que de verdad tengan datos.
