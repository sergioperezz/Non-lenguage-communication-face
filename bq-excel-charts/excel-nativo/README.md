# Panel de gráficos en Excel (Fase 1)

Motor de gráficos **en Excel nativo** (sin backend ni add-in; solo Excel + una
macro). El usuario combina parámetros y el gráfico **se actualiza solo**; un clic
lo copia a PowerPoint como **EMF**. Hoy con **datos de ejemplo**; en la Fase 2
esos datos vendrán de **BigQuery** (por tu driver ODBC).

## Archivos

| Archivo | Qué es |
|---|---|
| `Panel_Generico.xlsx` | El libro listo para usar (desplegables + gráfico) |
| `macro_Panel.vba` | Macro: cascadas, tipo/estilo de gráfico, rango dinámico, ejes y **copia a PowerPoint (EMF)** |
| `generar_panel.py` | Script que regenera el `.xlsx` (si cambias datos/estructura) |
| `mapa_casos_uso.html` / `.md` | Mapa de casos de uso por familias |
| `catalogo_graficos.md` | Qué gráficos usar en un AM, EMF a PowerPoint y notas de UX |

## El modelo: 10 parámetros

| Parámetro | Celda | Opciones |
|---|---|---|
| **Tipo de entidad** | B3 | Fondo · Cartera · Indice |
| **Entidad** (cascada del tipo) | B4 | los fondos / carteras / índices de ese tipo |
| **Grupo de métrica** | B5 | Rendimiento · Riesgo · Composicion · Costes · Liquidez · Valoracion |
| **Métrica** (cascada del grupo) | B6 | p. ej. Riesgo → Duración, TIR, Spread, Volatilidad, Beta |
| **Dimensión / eje X** | B7 | *Tiempo*: Mensual, Trimestral, Semestral, Anual · *Composición*: Activo, Geografia, Industria, Sector, Divisa, Rating |
| **Filtro: tipo de activo** | B8 | Todos · RF · RV |
| **Periodo** | B9 | MTD · YTD · 3M · 6M · 1A · 3A |
| **Benchmark** | B10 | Con benchmark · Sin benchmark |
| **Estilo benchmark** | B11 | Barras · Líneas |
| **Tipo de gráfico** | B12 | Columnas · Barras · Líneas · Área · Circular · Anillo · Radar |

> Nota: los tokens **Indice**, **Composicion**, **Valoracion**, **Geografia** van
> **sin tilde** en los desplegables (Excel los usa para resolver nombres internos
> vía `INDIRECT`).

**Ejemplo** ("duración por trimestre, cartera en barras + benchmark en línea") =
`B3=Fondo`, `B4=RF Privada A`, `B5=Riesgo`, `B6=Duración`, `B7=Trimestral`,
`B8=Todos`, `B9=3A`, `B10=Con benchmark`, `B11=Líneas`, `B12=Columnas`.

## Cómo funciona por dentro

Tres hojas:
- **Panel**: lo que usa el analista (los 10 desplegables + el gráfico).
- **Datos**: tabla en formato largo
  `Entidad | TipoActivo | Metrica | EjeXTipo | EjeXValor | Serie | Valor`.
  En la **Fase 2** la rellena la consulta a BigQuery.
- **Listas**: rangos con nombre para las cascadas y categorías.

Mecanismo:
1. **Entidad** cascadea del tipo (B3→B4) vía `INDIRECT("Ent_"&B3)`; la **métrica**
   cascadea del grupo (B5→B6) vía `INDIRECT("Grupo_"&B5)`.
2. La tabla de resultados (D:F) se calcula con **`SUMIFS`** según los parámetros;
   las categorías del eje X salen de la dimensión elegida (`INDIRECT("Cat_"&B7)`).
3. El **filtro de tipo de activo** (B8) usa un comodín en `SUMIFS` (`Todos`=`*`).
4. El **Periodo** (B9) acota los buckets recientes en dimensiones de tiempo.
5. El **gráfico** se enlaza solo a las **categorías visibles** (rango dinámico), y
   la **macro** aplica tipo de gráfico, el **estilo del benchmark** (barras o
   línea), reconstruye series según Benchmark, y pone título y títulos de eje.

## Puesta en marcha (una vez)

1. Abre `Panel_Generico.xlsx`.
2. Pega `macro_Panel.vba`: **Parte A** en la hoja *Panel* (clic derecho en la
   pestaña → *Ver código*) y **Parte B** en un *Módulo* estándar (Insertar → Módulo).
3. **Guardar como** → *.xlsm*.
4. Al reabrir, **Habilitar macros**.

### Exportar a PowerPoint (EMF)

La macro `CopiarAPowerPoint` copia el gráfico como **EMF vectorial** y lo pega en
PowerPoint. Para un clic: Insertar → una Forma/botón → *Asignar macro* →
`CopiarAPowerPoint`.

## Cómo adaptarlo

- **Pocos cambios**: edita la hoja **Datos** (y **Listas** si añades entidades/métricas).
- **Regenerar**: edita los diccionarios de `generar_panel.py` (ENTIDADES, GRUPOS,
  DIMS, CATEGORIAS, PARAMS) y ejecútalo. Tras regenerar, vuelve a pegar la macro y
  guarda como `.xlsm`.
- ⚠️ Si cambias el nº máximo de categorías (`MAX_CATS`), la macro usa un rango
  dinámico (celda B14), así que no hay que tocarla.

## Puente a la Fase 2 (BigQuery)

- Sustituyes la hoja **Datos** por una **consulta de Power Query** contra BigQuery
  (driver ODBC), con las mismas columnas y parametrizada por las celdas del panel.
- Descomentas `ThisWorkbook.RefreshAll` en la macro para refrescar al cambiar el
  ámbito de datos.
- **El gráfico y los desplegables no cambian**.

## Limitaciones del *mock* (se resuelven en la Fase 2)

- El filtro **Todos** en una entidad con RF y RV **suma** ambos (simplificación);
  la agregación correcta (media ponderada) es de la Fase 2.
- Filtrar a un tipo de activo que la entidad no tiene (p. ej. RF en un índice de
  renta variable) muestra vacío — es el comportamiento esperado.
- El **Periodo** es una aproximación; en la Fase 2 será un filtro de fechas exacto.
