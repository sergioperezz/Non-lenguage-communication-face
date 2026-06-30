"""Genera el panel GENÉRICO (Fase 1, v2).

La idea clave: NO se pre-genera un gráfico por caso. Hay UNA tabla de datos en
formato largo y CINCO parámetros que el usuario combina. Cualquier combinación
(activo × métrica × dimensión × comparación × tipo de gráfico) se construye al
vuelo con fórmulas + una macro.

Parámetros (celdas del Panel):
  B3 Tipo de activo   (RF / RV)                      -> filtro
  B4 Métrica          (cascada según B3)             -> qué se mide
  B5 Eje X/dimensión  (Trimestre/Sector/Region/...)  -> cómo se desglosa
  B6 Comparación      (Cartera+Benchmark/Solo...)    -> serie(s)
  B7 Tipo de gráfico  (Columnas/Barras/Líneas/...)   -> representación

En la Fase 2, la hoja "Datos" se sustituye por una consulta a BigQuery (ODBC),
con las MISMAS columnas, y el resto no cambia.
"""

from __future__ import annotations

import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# --- Definición del modelo de ejemplo --------------------------------------
# Liquidez aplica a ambos tipos de activo.
METRICAS = {
    "RF": ["Duration", "TIR", "Spread", "Liquidez"],
    "RV": ["Rentabilidad", "PER", "DividendYield", "Liquidez"],
}

# Qué dimensiones (eje X) aplican a cada tipo de activo.
EJES = {
    "RF": ["Mes", "Trimestre", "Sector", "Region", "Rating"],
    "RV": ["Mes", "Trimestre", "Sector", "Region"],
}


def _meses(n: int, y0: int, m0: int) -> list[str]:
    out, y, m = [], y0, m0
    for _ in range(n):
        out.append(f"{y}-{m:02d}")
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out


# Categorías ordenadas de cada dimensión (3 años para Mes/Trimestre).
CATEGORIAS = {
    "Mes": _meses(36, 2023, 7),  # 2023-07 .. 2026-06 (36 meses)
    "Trimestre": (
        ["2023-T3", "2023-T4"]
        + [f"{y}-T{q}" for y in (2024, 2025) for q in (1, 2, 3, 4)]
        + ["2026-T1", "2026-T2"]
    ),  # 12 trimestres
    "Sector": ["Financiero", "Industrial", "Tecnología", "Consumo", "Energía", "Salud"],
    "Region": ["Europa", "EEUU", "Asia", "Emergentes", "Global"],
    "Rating": ["AAA", "AA", "A", "BBB", "BB"],
}
MAX_CATS = max(len(c) for c in CATEGORIAS.values())  # filas de la tabla de resultados

# Valores base por métrica: (base, paso, factor del benchmark).
PARAMS = {
    "Duration": (5.0, 0.30, 0.92),
    "TIR": (3.5, 0.20, 1.05),
    "Spread": (120.0, 15.0, 0.85),
    "Rentabilidad": (9.0, 0.60, 0.90),
    "PER": (15.0, 0.80, 1.04),
    "DividendYield": (2.8, 0.20, 0.95),
    "Liquidez": (5.0, 0.10, 1.02),
}

AZUL = "FF0072CE"
GRIS = "FFF2F2F2"
BOLD = Font(bold=True)
BOLD_WHITE = Font(bold=True, color="FFFFFFFF")


def _valor(metrica: str, idx: int) -> float:
    """Valor ficticio plausible: oscila alrededor de la base (no una rampa fea)."""
    base, paso, _ = PARAMS[metrica]
    val = base + paso * (2.0 * math.sin(idx * 0.6) + (idx % 4) - 1.5)
    return round(val, 2)


def build() -> Workbook:
    wb = Workbook()

    # === Hoja Datos: TipoActivo | Metrica | EjeXTipo | EjeXValor | Serie | Valor ===
    ws_d = wb.active
    ws_d.title = "Datos"
    ws_d.append(["TipoActivo", "Metrica", "EjeXTipo", "EjeXValor", "Serie", "Valor"])
    for c in ws_d[1]:
        c.font = BOLD
    for tipo, metricas in METRICAS.items():
        for met in metricas:
            factor = PARAMS[met][2]
            for ejx in EJES[tipo]:
                for idx, cat in enumerate(CATEGORIAS[ejx]):
                    cartera = _valor(met, idx)
                    bench = round(cartera * factor, 2)
                    ws_d.append([tipo, met, ejx, cat, "Cartera", cartera])
                    ws_d.append([tipo, met, ejx, cat, "Benchmark", bench])

    # === Hoja Listas: cascadas (métricas, ejes) y categorías por dimensión ===
    ws_l = wb.create_sheet("Listas")

    def _col(letter: str, header: str, items: list[str]) -> str:
        ws_l[f"{letter}1"] = header
        ws_l[f"{letter}1"].font = BOLD
        for i, v in enumerate(items, start=2):
            ws_l[f"{letter}{i}"] = v
        return f"Listas!${letter}$2:${letter}${1 + len(items)}"

    ref_rf = _col("A", "RF", METRICAS["RF"])
    ref_rv = _col("B", "RV", METRICAS["RV"])
    ref_eje_rf = _col("C", "Eje_RF", EJES["RF"])
    ref_eje_rv = _col("D", "Eje_RV", EJES["RV"])
    ref_cat = {
        "Trimestre": _col("E", "Cat_Trimestre", CATEGORIAS["Trimestre"]),
        "Sector": _col("F", "Cat_Sector", CATEGORIAS["Sector"]),
        "Region": _col("G", "Cat_Region", CATEGORIAS["Region"]),
        "Rating": _col("H", "Cat_Rating", CATEGORIAS["Rating"]),
    }

    def _name(name: str, ref: str) -> None:
        wb.defined_names.add(DefinedName(name, attr_text=ref))

    _name("RF", ref_rf)
    _name("RV", ref_rv)
    _name("Eje_RF", ref_eje_rf)
    _name("Eje_RV", ref_eje_rv)
    _name("Cat_Trimestre", ref_cat["Trimestre"])
    _name("Cat_Sector", ref_cat["Sector"])
    _name("Cat_Region", ref_cat["Region"])
    _name("Cat_Rating", ref_cat["Rating"])

    # === Hoja Panel ===
    ws = wb.create_sheet("Panel", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Panel genérico de gráficos"
    ws["A1"].font = Font(bold=True, size=14)

    controles = {
        "A3": "Tipo de activo",
        "A4": "Métrica",
        "A5": "Eje X (dimensión)",
        "A6": "Comparación",
        "A7": "Tipo de gráfico",
    }
    for celda, txt in controles.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    ws["B3"], ws["B4"], ws["B5"] = "RF", "Duration", "Trimestre"
    ws["B6"], ws["B7"] = "Cartera + Benchmark", "Columnas"
    for celda in ("B3", "B4", "B5", "B6", "B7"):
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)

    # Desplegables (validación de datos)
    dvs = [
        ("B3", '"RF,RV"'),
        ("B4", "=INDIRECT($B$3)"),                       # métrica en cascada
        ("B5", '=INDIRECT("Eje_"&$B$3)'),                # ejes válidos por activo
        ("B6", '"Cartera + Benchmark,Solo cartera,Solo benchmark"'),
        ("B7", '"Columnas,Barras,Líneas,Área,Circular,Anillo,Radar"'),
    ]
    for celda, formula in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[celda])

    # Título dinámico
    ws.merge_cells("A9:C9")
    ws["A9"] = '="Métrica: "&B4&"   ·   Por: "&B5&"   ·   "&B6'
    ws["A9"].font = Font(bold=True, size=12, color=AZUL[2:])
    ws["A9"].alignment = Alignment(horizontal="left")

    # Tabla de resultados: Categoría | Cartera | Benchmark (calculada al vuelo)
    ws["D2"], ws["E2"], ws["F2"] = "Categoría", "Cartera", "Benchmark"
    for celda in ("D2", "E2", "F2"):
        ws[celda].font = BOLD_WHITE
        ws[celda].fill = PatternFill("solid", fgColor=AZUL)

    for i in range(MAX_CATS):
        r = 3 + i
        # Categoría: la i-ésima de la dimensión elegida (vacío si no existe).
        ws.cell(row=r, column=4,
                value=f'=IFERROR(INDEX(INDIRECT("Cat_"&$B$5),{i + 1}),"")')
        # Cartera
        ws.cell(row=r, column=5, value=(
            f'=IF(OR($B$6="Solo benchmark",$D{r}=""),"",'
            f'SUMIFS(Datos!$F:$F,Datos!$A:$A,$B$3,Datos!$B:$B,$B$4,'
            f'Datos!$C:$C,$B$5,Datos!$D:$D,$D{r},Datos!$E:$E,"Cartera"))'
        ))
        # Benchmark
        ws.cell(row=r, column=6, value=(
            f'=IF(OR($B$6="Solo cartera",$D{r}=""),"",'
            f'SUMIFS(Datos!$F:$F,Datos!$A:$A,$B$3,Datos!$B:$B,$B$4,'
            f'Datos!$C:$C,$B$5,Datos!$D:$D,$D{r},Datos!$E:$E,"Benchmark"))'
        ))

    # Gráfico enlazado a la tabla (la macro ajusta tipo y el combo barras+línea)
    chart = BarChart()
    chart.type = "col"
    chart.title = "Cartera vs Benchmark"
    chart.height = 8.5
    chart.width = 16
    last = 2 + MAX_CATS
    data = Reference(ws, min_col=5, max_col=6, min_row=2, max_row=last)  # E2:F8 con cabeceras
    cats = Reference(ws, min_col=4, min_row=3, max_row=last)            # D3:D8
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "H2")

    for col, w in {"A": 18, "B": 20, "D": 14, "E": 12, "F": 12}.items():
        ws.column_dimensions[col].width = w

    return wb


if __name__ == "__main__":
    out = Path(__file__).parent / "Panel_Generico.xlsx"
    build().save(out)
    print(f"Generado: {out}  (max categorías por dimensión: {MAX_CATS})")
