# Mapa de casos de uso — Panel de gráficos

> Cada caso = una combinación de parámetros = una diapositiva.
> Modelo: **Tipo entidad → Entidad → Grupo → Métrica → Dimensión → (Filtro) → Periodo → Benchmark (+estilo) → Gráfico**.
> Las tablas son **ejemplos**, no plantillas cerradas: el motor admite cualquier combinación.

## El espacio de parámetros

| Parámetro | Valores |
|---|---|
| **Tipo de entidad** | Fondo · Cartera · Indice |
| **Entidad 1-3** (cascada) | 1 principal + hasta 2 para **comparar varias a la vez** |
| **Grupo de métrica** | Rendimiento · Riesgo · Composicion · Costes · Liquidez · Valoracion |
| **Métrica** (cascada del grupo) | Rend.: Rentabilidad, Rentab. acum. · Riesgo: Duración, TIR, Spread, Volatilidad, Beta · Comp.: Peso · Costes: TER · Liquidez: Liquidez · Valor.: PER, DividendYield |
| **Dimensión (eje X)** | *Tiempo*: Mensual, Trimestral, Semestral, Anual — *Composición*: Activo, Geografia, Industria, Sector, Divisa, Rating |
| **Filtro: tipo de activo** | Todos · RF · RV *(útil en fondos mixtos)* |
| **Periodo** | MTD · YTD · 1M–6M · 1A–6A |
| **Benchmark** | Con benchmark · Sin benchmark |
| **Estilo benchmark** | Barras · Líneas |
| **Gráfico** | Columnas · Barras · Líneas · Área · Circular · Anillo · Radar |

---

## ① Rendimiento

| Caso | Fondo | Métrica | Dimensión | Filtro | Periodo | Benchmark | Gráfico |
|---|---|---|---|---|---|---|---|
| Evolución de rentabilidad | Bolsa Europa | Rentabilidad | Mensual | Todos | 1A | Sin | Líneas |
| Rentabilidad vs benchmark | Bolsa Europa | Rentabilidad | Trimestral | Todos | 3A | Con | Columnas (combo) |
| Rentabilidad YTD por geografía | Bolsa Europa | Rentabilidad | Geografia | Todos | YTD | Sin | Barras |
| Rentab. acumulada | Mixto Moderado | Rentab. acum. | Mensual | Todos | 3A | Con | Líneas |

## ② Riesgo

| Caso | Fondo | Métrica | Dimensión | Filtro | Periodo | Benchmark | Gráfico |
|---|---|---|---|---|---|---|---|
| Duración cartera vs benchmark | RF Privada A | Duración | Trimestral | Todos | 3A | Con | Columnas (combo) |
| Duración solo de la RF (fondo mixto) | Mixto Moderado | Duración | Trimestral | **RF** | 3A | Sin | Columnas |
| TIR por rating | RF Privada A | TIR | Rating | RF | 3A | Sin | Barras |
| Spread por sector | RF Privada A | Spread | Sector | RF | 3A | Sin | Columnas |
| Volatilidad (12m) | Bolsa Europa | Volatilidad | Mensual | Todos | 1A | Con | Líneas |

## ③ Composición / Asset allocation

| Caso | Fondo | Métrica | Dimensión | Filtro | Periodo | Benchmark | Gráfico |
|---|---|---|---|---|---|---|---|
| Distribución por tipo de activo | Mixto Moderado | Peso | Activo | Todos | 3A | Sin | Anillo |
| Distribución geográfica | Mixto Moderado | Peso | Geografia | Todos | 3A | Sin | Circular |
| Distribución sectorial (RV) | Mixto Moderado | Peso | Sector | **RV** | 3A | Sin | Barras |
| Exposición por divisa | RF Privada A | Peso | Divisa | Todos | 3A | Sin | Anillo |
| Calidad crediticia (rating) | RF Privada A | Peso | Rating | RF | 3A | Sin | Columnas |

## ④ Costes

| Caso | Fondo | Métrica | Dimensión | Filtro | Periodo | Benchmark | Gráfico |
|---|---|---|---|---|---|---|---|
| Evolución del TER | Bolsa Europa | TER | Trimestral | Todos | 3A | Sin | Líneas |
| TER vs categoría | Bolsa Europa | TER | Anual | Todos | 3A | Con | Columnas |

## ⑤ Liquidez

| Caso | Fondo | Métrica | Dimensión | Filtro | Periodo | Benchmark | Gráfico |
|---|---|---|---|---|---|---|---|
| Evolución de la liquidez | Mixto Moderado | Liquidez | Mensual | Todos | 3A | Sin | Área |
| Liquidez por tipo de activo | Mixto Moderado | Liquidez | Activo | Todos | MTD | Sin | Columnas |

## ⑥ Valoración

| Caso | Fondo | Métrica | Dimensión | Filtro | Periodo | Benchmark | Gráfico |
|---|---|---|---|---|---|---|---|
| PER por sector | Bolsa Europa | PER | Sector | RV | 3A | Con | Barras |
| Rentabilidad por dividendo | Bolsa Europa | DividendYield | Geografia | RV | 3A | Sin | Columnas |

---

## Cómo se traduce a la herramienta

- **Hoy (Fase 1):** cualquiera de estos casos se arma en el panel con datos de
  ejemplo (cobertura completa) → eliges parámetros y se ve al instante.
- **Fase 2:** los mismos casos tiran de **datos reales de BigQuery**; el periodo
  y el filtro de tipo de activo pasan a ser condiciones exactas de la query.
- **Fase 3:** un clic copia el gráfico a PowerPoint como **EMF** (vectorial).

Los casos recurrentes pueden guardarse como **presets** (configuraciones
aprobadas) para el informe trimestral; el modo ad-hoc queda para lo nuevo.

> Nota del *mock*: el filtro "Todos" suma los tipos de activo presentes en un
> fondo mixto (simplificación); la agregación correcta (media ponderada) es de la
> Fase 2 con datos reales.
