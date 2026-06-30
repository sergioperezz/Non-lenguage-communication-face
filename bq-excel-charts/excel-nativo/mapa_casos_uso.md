# Mapa de casos de uso — Panel de gráficos

> Cada caso = una combinación de **6 parámetros** = una diapositiva.
> Estas tablas son **ejemplos representativos**, no plantillas cerradas: el motor
> admite cualquier combinación (decenas de miles), no solo las listadas.

## El espacio de parámetros

| Parámetro | Valores |
|---|---|
| **Tipo de activo** | RF · RV |
| **Métrica** | Rentabilidad · TIR · TER · Duración · Spread · Liquidez · Peso · (RV: PER · DividendYield) |
| **Dimensión (eje X)** | *Tiempo*: Mensual · Trimestral · Semestral · Anual — *Composición*: Activo · Geografía · Industria · Sector · Divisa · Rating |
| **Periodo** | MTD · YTD · 3M · 6M · 1A · 3A |
| **Comparación** | Cartera · Benchmark · Ambos |
| **Gráfico** | Columnas · Barras · Líneas · Área · Circular · Anillo · Radar |

Número teórico de combinaciones (orden de magnitud): 2 × ~7 × ~10 × 6 × 3 × 7 ≈ **17.000+**. De ahí la importancia del motor genérico: no se generan a mano.

---

## ① Rendimiento

| Caso | Activo | Métrica | Dimensión | Periodo | Comparación | Gráfico |
|---|---|---|---|---|---|---|
| Evolución de rentabilidad | RV | Rentabilidad | Mensual | 1A | Cartera | Líneas |
| Rentabilidad vs benchmark | RV | Rentabilidad | Trimestral | 3A | Ambos | Columnas (combo) |
| Rentabilidad YTD por geografía | RV | Rentabilidad | Geografía | YTD | Cartera | Barras |
| Rentabilidad anual | RV | Rentabilidad | Anual | 3A | Ambos | Columnas |

## ② Riesgo (renta fija)

| Caso | Activo | Métrica | Dimensión | Periodo | Comparación | Gráfico |
|---|---|---|---|---|---|---|
| Evolución de duración (vs bench) | RF | Duración | Mensual | 1A | Ambos | Columnas (combo) |
| TIR por rating | RF | TIR | Rating | 3A | Cartera | Barras |
| Spread por sector | RF | Spread | Sector | 3A | Cartera | Columnas |
| Evolución de TIR | RF | TIR | Trimestral | 3A | Cartera | Líneas |

## ③ Composición / Asset allocation

| Caso | Activo | Métrica | Dimensión | Periodo | Comparación | Gráfico |
|---|---|---|---|---|---|---|
| Peso por tipo de activo | RF/RV | Peso | Activo | 3A | Cartera | Anillo |
| Peso por sector | RV | Peso | Sector | 3A | Cartera | Barras |
| Peso por geografía | RV | Peso | Geografía | 3A | Cartera | Circular |
| Peso por divisa | RF/RV | Peso | Divisa | 3A | Cartera | Anillo |
| Peso por rating | RF | Peso | Rating | 3A | Cartera | Columnas |

## ④ Costes

| Caso | Activo | Métrica | Dimensión | Periodo | Comparación | Gráfico |
|---|---|---|---|---|---|---|
| Evolución del TER | RF/RV | TER | Trimestral | 3A | Cartera | Líneas |
| TER vs benchmark/categoría | RF/RV | TER | Anual | 3A | Ambos | Columnas |

## ⑤ Liquidez

| Caso | Activo | Métrica | Dimensión | Periodo | Comparación | Gráfico |
|---|---|---|---|---|---|---|
| Evolución de la liquidez | RF/RV | Liquidez | Mensual | 3A | Cartera | Área |
| Liquidez por tipo de activo | RF/RV | Liquidez | Activo | MTD | Cartera | Columnas |

## ⑥ Valoración (renta variable)

| Caso | Activo | Métrica | Dimensión | Periodo | Comparación | Gráfico |
|---|---|---|---|---|---|---|
| PER por sector | RV | PER | Sector | 3A | Ambos | Barras |
| Rentabilidad por dividendo | RV | DividendYield | Geografía | 3A | Cartera | Columnas |

---

## Cómo se traduce a la herramienta

- **Hoy (Fase 1):** cualquiera de estos casos se construye en el panel con datos
  de ejemplo (cobertura completa) → eliges parámetros y se ve al instante.
- **Fase 2:** los mismos casos tiran de **datos reales de BigQuery**; el periodo
  pasa a ser un filtro de fechas exacto.
- **Fase 3:** un clic copia el gráfico a PowerPoint como **EMF** (vectorial).

Los casos recurrentes pueden guardarse como **presets** (configuraciones
aprobadas) para el informe trimestral, reservando el modo ad-hoc para lo nuevo.
