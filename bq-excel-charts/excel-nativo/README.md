# Panel interactivo en Excel (Fase 1)

Gráfico que **se actualiza solo** cuando el usuario cambia tres parámetros, sin
backend ni add-in. Solo Excel (+ una macro pequeña para el tipo de gráfico).

## Archivos

| Archivo | Qué es |
|---|---|
| `Panel_Interactivo.xlsx` | El libro listo para usar (desplegables + gráfico) |
| `macro_Panel.vba` | Macro que cambia el tipo de gráfico y cuida la cascada |
| `generar_panel.py` | Script que regenera el `.xlsx` (por si cambias los datos) |

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
