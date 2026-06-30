# Catálogo de gráficos para una gestora + copia a PowerPoint (EMF)

## 1. Qué gráficos tienen sentido en un AM (y cuándo)

No se trata de generar "todas" las combinaciones posibles (serían cientos y la
mayoría no aportan), sino de un **set curado** y elegir el adecuado al *tipo de
dato*. La clave es la **forma del dato**, no el gusto.

### Disponibles en el panel (1–2 series: cartera y, opcionalmente, benchmark)

| Gráfico | Cuándo usarlo | Tipo Excel (VBA) |
|---|---|---|
| **Columnas** (vertical) | Comparar una métrica entre pocas categorías (sector, región) | `xlColumnClustered` |
| **Barras** (horizontal) | Rankings o etiquetas largas (emisores, fondos) | `xlBarClustered` |
| **Líneas** | Evolución temporal (TIR, NAV, rentabilidad acumulada) | `xlLineMarkers` |
| **Área** | Evolución con énfasis en magnitud | `xlArea` |
| **Circular / Anillo** | Reparto de un total en un momento (asignación) — con pocas categorías | `xlPie` / `xlDoughnut` |
| **Radar** | Perfil multi-eje (p. ej. estilo/factores) | `xlRadarMarkers` |

> El panel ya incluye estos siete en el desplegable. Cambiar entre ellos es
> instantáneo (solo cambia `.ChartType` de un único gráfico).

> El **combo cartera + benchmark** (barras + línea) **ya está** en el panel: se
> activa con *Columnas + Con benchmark*.

### Más tipos multi-serie (llegan con datos reales en la Fase 2)

| Gráfico | Cuándo usarlo | Nota técnica |
|---|---|---|
| **Columnas apiladas / 100%** | Composición que suma al total (peso por sector, asset allocation en el tiempo) | `xlColumnStacked` / `xlColumnStacked100` (≥2 series) |
| **Dispersión XY** | Riesgo/retorno (volatilidad vs rentabilidad), duración vs TIR | `xlXYScatter` (series X,Y) |
| **Cascada / Waterfall** | Atribución de resultados, *bridge* de P&L o de rentabilidad | Gráfico moderno; mejor vía **plantilla .crtx** |

**Recomendación práctica:** estos se añaden cuando la consulta a BigQuery pueda
devolver varias series. Para *waterfall* lo más robusto es guardar una
**plantilla de gráfico corporativa (.crtx)** y aplicarla.

## 2. Cómo evitar el "todas las combinaciones" sin perder flexibilidad

En lugar de pre-generar N gráficos estáticos:
- **Un único gráfico** enlazado a la tabla de resultados; el **dato** lo cambian
  los desplegables (fondo, grupo, métrica, dimensión, filtro…) y el **tipo** lo
  cambia la macro.
- Así, "todas las combinaciones" = (fondos) × (métricas) × (dimensiones) ×
  (7 tipos de gráfico) **se generan al vuelo**, sin duplicar objetos ni inflar el
  libro.
- Si quieres formato perfecto por tipo, aplica una **plantilla .crtx** por tipo
  (corporativa) en `AplicarTipoGrafico` — formato consistente y sin reconstruir.

## 3. Copia a PowerPoint como EMF (que NO se rompe)

### Por qué EMF y no pegar el gráfico nativo

| Forma de pegar | Resultado | Riesgo |
|---|---|---|
| Gráfico nativo de Excel | Editable, pero lleva un **libro vinculado** | "No se puede editar", **vínculos rotos**, pesa más |
| Imagen **PNG/bitmap** | Estable, pero **se pixela** al escalar | Mala calidad en pantalla/impresión |
| **EMF (metafile vectorial)** ✅ | Imagen **vectorial**, nítida a cualquier tamaño, **sin vínculo** | Ninguno relevante (no editable, que para deck de cliente es lo deseable) |

Para presentaciones de cliente en una gestora, **EMF es el estándar**: pixel
perfect, escalable, sin vínculos que se rompan al mover el archivo, y más fácil
para Compliance (es una imagen fija, no datos vivos).

### Cómo se hace (ya implementado en la macro `CopiarAPowerPoint`)

1. `Chart.CopyPicture Appearance:=xlScreen, Format:=xlPicture`
   → en Windows copia el gráfico como **metafile (EMF)**, vectorial.
2. Se conecta a PowerPoint (lo abre si hace falta; *late binding*, sin
   configurar referencias).
3. `Slide.Shapes.PasteSpecial(DataType:=2)` (`ppPasteEnhancedMetafile`)
   → pega como **EMF**.
4. Centra la imagen en la diapositiva.

> **Nota de plataforma:** EMF es de **Windows**. En Mac, PowerPoint no soporta
> EMF (usa PDF/PICT); si hay usuarios Mac, se adapta el paso de pegado.

## 4. UX óptima y rapidez

- **Cambio instantáneo:** se cambia `.ChartType` de un solo gráfico (no se
  reconstruye) y se envuelve en `Application.ScreenUpdating = False` para que no
  parpadee.
- **Sin estados imposibles:** al cambiar el grupo de métrica, la macro corrige la
  métrica si dejó de ser válida para ese grupo.
- **Datos acotados (Fase 2):** usa rangos acotados en `SUMIFS`/consultas, no
  columnas enteras, para que el recálculo sea rápido con datos reales.
- **Botón de exportar:** asigna `CopiarAPowerPoint` a una forma/botón en el panel
  → un clic y el gráfico está en la diapositiva como EMF.
- **Formato corporativo:** plantilla `.crtx` única para que todos los gráficos
  salgan con la identidad visual de la gestora.
