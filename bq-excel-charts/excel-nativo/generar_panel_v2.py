"""Genera el panel GENÉRICO (Fase 1) — modelo Fondo -> Grupo -> Métrica.

Motor guiado por datos: una tabla en formato largo + parámetros que el usuario
combina. Jerarquía:

  Fondo / Cartera   -> ¿qué producto?            (incluye uno MIXTO)
  Grupo de métrica  -> familia (Rendimiento...)  -> agrupa para no saturar
  Métrica           -> cascada del grupo
  Dimensión (eje X) -> tiempo o composición
  Filtro Tipo activo-> opcional (Todos/RF/RV); útil en fondos mixtos
  Periodo           -> ventana (MTD/YTD/3M/6M/1A/3A)
  Comparación       -> Cartera/Benchmark/Ambos
  Tipo de gráfico   -> representación

Datos de ejemplo de COBERTURA COMPLETA: cualquier combinación se ve al instante.
Nota: el filtro "Todos" suma tipos (mock); la agregación correcta es de la Fase 2.
"""

from __future__ import annotations

import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation


# === Modelo ================================================================
FONDOS = {                       # fondo -> tipos de activo que contiene
    "RF Privada A": ["RF"],
    "Bolsa Europa": ["RV"],
    "Mixto Moderado": ["RF", "RV"],
}

# Métricas agrupadas por familia (tokens de grupo SIN acentos: se usan en nombres).
GRUPOS = {
    "Rendimiento": ["Rentabilidad", "Rentab. acum."],
    "Riesgo": ["Duración", "TIR", "Spread", "Volatilidad", "Beta"],
    "Composicion": ["Peso"],
    "Costes": ["TER"],
    "Liquidez": ["Liquidez"],
    "Valoracion": ["PER", "DividendYield"],
}
ALL_METRICS = [m for ms in GRUPOS.values() for m in ms]

DIM_TIEMPO = ["Mensual", "Trimestral", "Semestral", "Anual"]
DIM_COMP = ["Activo", "Geografia", "Industria", "Sector", "Divisa", "Rating"]
DIMS = DIM_TIEMPO + DIM_COMP


def _meses(n: int, y0: int, m0: int) -> list[str]:
    out, y, m = [], y0, m0
    for _ in range(n):
        out.append(f"{y}-{m:02d}")
        m += 1
        if m > 12:
            m, y = 1, y + 1
    return out


CATEGORIAS = {
    "Mensual": _meses(36, 2023, 7),
    "Trimestral": (["2023-T3", "2023-T4"]
                   + [f"{y}-T{q}" for y in (2024, 2025) for q in (1, 2, 3, 4)]
                   + ["2026-T1", "2026-T2"]),
    "Semestral": ["2023-S2", "2024-S1", "2024-S2", "2025-S1", "2025-S2", "2026-S1"],
    "Anual": ["2024", "2025", "2026"],
    "Activo": ["Deuda Pública", "Deuda Privada", "Acciones", "Liquidez", "Derivados"],
    "Geografia": ["Europa", "Norteamérica", "Asia-Pacífico", "Emergentes", "Latam"],
    "Industria": ["Financiero", "Tecnología", "Salud", "Energía", "Consumo",
                  "Industrial", "Utilities"],
    "Sector": ["Financiero", "Industrial", "Tecnología", "Consumo", "Energía", "Salud"],
    "Divisa": ["EUR", "USD", "GBP", "JPY", "CHF", "Otras"],
    "Rating": ["AAA", "AA", "A", "BBB", "BB", "B"],
}
MAX_CATS = max(len(c) for c in CATEGORIAS.values())

PARAMS = {                       # (base, paso, factor benchmark)
    "Rentabilidad": (9.0, 0.60, 0.90),
    "Rentab. acum.": (12.0, 1.20, 0.90),
    "Duración": (5.0, 0.30, 0.92),
    "TIR": (3.5, 0.20, 1.05),
    "Spread": (120.0, 15.0, 0.85),
    "Volatilidad": (12.0, 0.80, 1.10),
    "Beta": (1.0, 0.05, 1.00),
    "Peso": (18.0, 3.0, 1.00),
    "TER": (0.85, 0.03, 1.00),
    "Liquidez": (5.0, 0.10, 1.02),
    "PER": (15.0, 0.80, 1.04),
    "DividendYield": (2.8, 0.20, 0.95),
}

PERIODOS = ["MTD", "YTD", "3M", "6M", "1A", "3A"]
TABLA_N = {
    "MTD": [1, 1, 1, 1], "YTD": [6, 2, 1, 1], "3M": [3, 1, 1, 1],
    "6M": [6, 2, 1, 1], "1A": [12, 4, 2, 1], "3A": [36, 12, 6, 3],
}

AZUL, GRIS = "FF0072CE", "FFF2F2F2"
BOLD = Font(bold=True)
BOLD_WHITE = Font(bold=True, color="FFFFFFFF")


def _valor(metrica: str, idx: int) -> float:
    base, paso, _ = PARAMS[metrica]
    return round(base + paso * (2.0 * math.sin(idx * 0.6) + (idx % 4) - 1.5), 2)


def build() -> Workbook:
    wb = Workbook()

    # === Datos: Fondo|TipoActivo|Metrica|EjeXTipo|EjeXValor|Serie|Valor ===
    ws_d = wb.active
    ws_d.title = "Datos"
    ws_d.append(["Fondo", "TipoActivo", "Metrica", "EjeXTipo", "EjeXValor", "Serie", "Valor"])
    for c in ws_d[1]:
        c.font = BOLD
    for fondo, tipos in FONDOS.items():
        for tipo in tipos:
            for met in ALL_METRICS:
                factor = PARAMS[met][2]
                for dim in DIMS:
                    for idx, cat in enumerate(CATEGORIAS[dim]):
                        v = _valor(met, idx)
                        ws_d.append([fondo, tipo, met, dim, cat, "Cartera", v])
                        ws_d.append([fondo, tipo, met, dim, cat, "Benchmark", round(v * factor, 2)])
    n = ws_d.max_row

    # === Listas ===
    ws_l = wb.create_sheet("Listas")

    def col(letter, header, items):
        ws_l[f"{letter}1"] = header
        ws_l[f"{letter}1"].font = BOLD
        for i, v in enumerate(items, start=2):
            ws_l[f"{letter}{i}"] = v
        return f"Listas!${letter}$2:${letter}${1 + len(items)}"

    def name(nm, ref):
        wb.defined_names.add(DefinedName(nm, attr_text=ref))

    grp_cols = ["A", "B", "C", "D", "E", "F"]
    for letter, g in zip(grp_cols, GRUPOS):
        name(f"Grupo_{g}", col(letter, f"Grupo_{g}", GRUPOS[g]))

    cat_cols = ["G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
    for letter, dim in zip(cat_cols, CATEGORIAS):
        name(f"Cat_{dim}", col(letter, f"Cat_{dim}", CATEGORIAS[dim]))

    name("Periodos", col("R", "Periodos", PERIODOS))
    name("DimTiempo", col("S", "DimTiempo", DIM_TIEMPO))
    ws_l["T1"], ws_l["U1"], ws_l["V1"], ws_l["W1"] = ("N_Men", "N_Tri", "N_Sem", "N_Anu")
    for r, p in enumerate(PERIODOS, start=2):
        for cidx, val in enumerate(TABLA_N[p]):
            ws_l.cell(row=r, column=20 + cidx, value=val)  # 20 = T
    name("TablaN", f"Listas!$T$2:$W${1 + len(PERIODOS)}")

    # === Panel ===
    ws = wb.create_sheet("Panel", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Panel genérico de gráficos"
    ws["A1"].font = Font(bold=True, size=14)

    etiquetas = {
        "A3": "Fondo / Cartera", "A4": "Grupo de métrica", "A5": "Métrica",
        "A6": "Dimensión (eje X)", "A7": "Filtro: tipo de activo", "A8": "Periodo",
        "A9": "Benchmark", "A10": "Tipo de gráfico",
    }
    for celda, txt in etiquetas.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    defaults = {
        "B3": "Mixto Moderado", "B4": "Riesgo", "B5": "Duración", "B6": "Trimestral",
        "B7": "Todos", "B8": "3A", "B9": "Con benchmark", "B10": "Columnas",
    }
    for celda, val in defaults.items():
        ws[celda] = val
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)

    fondos_lst = ",".join(FONDOS)
    grupos_lst = ",".join(GRUPOS)
    dims_lst = ",".join(DIMS)
    dvs = [
        ("B3", f'"{fondos_lst}"'),
        ("B4", f'"{grupos_lst}"'),
        ("B5", '=INDIRECT("Grupo_"&$B$4)'),
        ("B6", f'"{dims_lst}"'),
        ("B7", '"Todos,RF,RV"'),
        ("B8", '"MTD,YTD,3M,6M,1A,3A"'),
        ("B9", '"Con benchmark,Sin benchmark"'),
        ("B10", '"Columnas,Barras,Líneas,Área,Circular,Anillo,Radar"'),
    ]
    for celda, formula in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[celda])

    # Helpers
    ws["A12"] = "Buckets visibles"
    ws["A12"].font = Font(italic=True, size=9)
    ws["B12"] = ('=IFERROR(INDEX(TablaN,MATCH($B$8,Periodos,0),MATCH($B$6,DimTiempo,0)),'
                 'COUNTA(INDIRECT("Cat_"&$B$6)))')
    ws["A13"] = "Criterio tipo activo"
    ws["A13"].font = Font(italic=True, size=9)
    ws["B13"] = '=IF($B$7="Todos","*",$B$7)'   # "*" = todos en SUMIFS

    # Título dinámico
    ws.merge_cells("A15:C15")
    ws["A15"] = ('="Fondo: "&B3&"  ·  "&B5&" por "&B6&"  ·  "&B8'
                 '&IF(B9="Con benchmark","  ·  vs benchmark","")')
    ws["A15"].font = Font(bold=True, size=12, color=AZUL[2:])
    ws["A15"].alignment = Alignment(horizontal="left")

    # Tabla de resultados
    ws["D2"], ws["E2"], ws["F2"] = "Categoría", "Cartera", "Benchmark"
    for celda in ("D2", "E2", "F2"):
        ws[celda].font = BOLD_WHITE
        ws[celda].fill = PatternFill("solid", fgColor=AZUL)

    def sumifs(serie):
        return (
            f'SUMIFS(Datos!$G$2:$G${n},Datos!$A$2:$A${n},$B$3,'
            f'Datos!$B$2:$B${n},$B$13,Datos!$C$2:$C${n},$B$5,'
            f'Datos!$D$2:$D${n},$B$6,Datos!$E$2:$E${n},$D{{r}},'
            f'Datos!$F$2:$F${n},"{serie}")'
        )

    for i in range(MAX_CATS):
        r = 3 + i
        ws.cell(row=r, column=4, value=(
            f'=IF((ROW()-2)>$B$12,"",INDEX(INDIRECT("Cat_"&$B$6),'
            f'COUNTA(INDIRECT("Cat_"&$B$6))-$B$12+(ROW()-2)))'
        ))
        ws.cell(row=r, column=5, value=(
            f'=IF($D{r}="","",{sumifs("Cartera").format(r=r)})'      # cartera siempre
        ))
        ws.cell(row=r, column=6, value=(
            f'=IF(OR($B$9="Sin benchmark",$D{r}=""),"",{sumifs("Benchmark").format(r=r)})'
        ))

    # Gráfico
    chart = BarChart()
    chart.type = "col"
    chart.title = "Cartera vs Benchmark"
    chart.height = 8.5
    chart.width = 18
    last = 2 + MAX_CATS
    chart.add_data(Reference(ws, min_col=5, max_col=6, min_row=2, max_row=last),
                   titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=4, min_row=3, max_row=last))
    ws.add_chart(chart, "H2")

    for c, w in {"A": 20, "B": 22, "D": 16, "E": 12, "F": 12}.items():
        ws.column_dimensions[c].width = w

    return wb


if __name__ == "__main__":
    out = Path(__file__).parent / "Panel_Generico.xlsx"
    build().save(out)
    print(f"Generado: {out}  (máx categorías: {MAX_CATS})")
