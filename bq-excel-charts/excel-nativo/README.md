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
| **Estilo benchmark** | B13 | Barras · Líneas |
| **Tipo de gráfico** | B14 | Columnas · Barras · Líneas · Área · Circular · Anillo · Radar |

> Nota: los tokens **Indice**, **Composicion**, **Valoracion**, **Geografia** van
> **sin tilde** en los desplegables (Excel los usa para nombres internos vía `INDIRECT`).

**Comparar varias entidades**: cada Entidad seleccionada (B4/B5/B6) es una serie
del gráfico. Deja B5/B6 vacías si no quieres comparar.

**Ejemplo (tu caso)** — *Rentabilidad de Cartera RF Gobierno vs Cartera RF Crédito*:
`B3=Cartera`, `B4=Cartera RF Gobierno`, `B5=Cartera RF Crédito`, `B7=Rendimiento`,
`B8=Rentabilidad`, `B9=Trimestral`, `B12=Sin benchmark`, `B14=Columnas`.

## Cómo funciona por dentro

Tres hojas:
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

## Puente a la Fase 2 (BigQuery)

- Sustituyes **Datos** por una consulta de Power Query (ODBC) con las mismas
  columnas. Descomentas `ThisWorkbook.RefreshAll` en la macro. El resto no cambia.

## Limitaciones del *mock* (se resuelven en la Fase 2)

- El filtro **Todos** en una entidad con RF y RV **suma** ambos (aproximación).
- Filtrar a un tipo de activo que la entidad no tiene muestra vacío (esperado).
- Solo hay **3 años** de datos, así que 3A–6A muestran lo mismo.
