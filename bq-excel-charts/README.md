# BQ Excel Charts

Herramienta para construir **gráficos y tablas en Excel** a partir de datos de
**BigQuery**, sin escribir SQL, y copiarlos a **PowerPoint**. Pensada para una
gestora de activos (informes de fondos/carteras).

## Enfoque: Excel nativo

Todo vive **en Excel**: desplegables + una macro VBA. Sin backend, sin add-in.
Es el camino de **menor fricción** para un entorno corporativo de banca (no
requiere desplegar aplicaciones ni hostear servidores; basta con habilitar
macros y, en la Fase 2, el driver ODBC de BigQuery que ya tienes).

El usuario combina parámetros (Fondo → Grupo de métrica → Métrica → Dimensión →
Filtro → Periodo → Benchmark → Gráfico) y el gráfico **se actualiza solo**; un
clic lo copia a PowerPoint como **EMF** (vectorial, sin vínculos que se rompan).

## Estructura

```
excel-nativo/   Panel en Excel: Panel_Generico.xlsx + macro VBA + generador + docs
```

Toda la documentación y los archivos están en
**[`excel-nativo/`](excel-nativo/)** — empieza por su
[`README.md`](excel-nativo/README.md).

## Fases

- **Fase 1 — hecho:** panel interactivo nativo con **datos de ejemplo** y
  prototipo de **copia a PowerPoint (EMF)**.
- **Fase 2:** datos **reales** de BigQuery vía Power Query + driver ODBC (modo
  "Generar con datos reales").
- **Fase 3:** automatizar el pegado a PowerPoint (plantilla corporativa, varias
  diapositivas, presets de casos).
