"""Genera el panel GENÉRICO (Fase 1) — modelo Entidad -> Grupo -> Métrica.

Motor guiado por datos. Jerarquía:

  Tipo de entidad   -> Fondo / Cartera / Indice
  Entidad           -> cascada del tipo (el fondo/cartera/índice concreto)
  Grupo de métrica  -> familia (Rendimiento...)  -> agrupa para no saturar
  Métrica           -> cascada del grupo
  Dimensión (eje X) -> tiempo o composición
  Filtro Tipo activo-> opcional (Todos/RF/RV)
  Periodo           -> ventana (MTD/YTD/3M/6M/1A/3A)
  Benchmark         -> Con / Sin
  Estilo benchmark  -> Barras / Líneas
  Tipo de gráfico   -> representación

Datos de ejemplo (ficticios). En la Fase 2 vienen de BigQuery por ODBC.
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
# Tipo de entidad -> {entidad: [tipos de activo que contiene]}. Tokens de tipo
# SIN acentos ("Indice") porque se usan para nombres definidos vía INDIRECT.
ENTIDADES = {
    "Fondo": {
        "RF Privada A": ["RF"],
        "Bolsa Europa": ["RV"],
        "Mixto Moderado": ["RF", "RV"],
    },
    "Cartera": {
        "Cartera Conservadora": ["RF", "RV"],
        "Cartera Dinamica": ["RF", "RV"],
    },
    "Indice": {
        "Euro Stoxx 50": ["RV"],
        "Bloomberg Agg": ["RF"],
        "MSCI World": ["RV"],
    },
}

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


def _meses(n, y0, m0):
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
# OJO: la tabla de resultados ocupa las filas 3..(2+MAX_CATS). Con MAX_CATS=36 -> 3:38.
# El gráfico usa un rango dinámico (la macro lo ajusta a las categorías visibles).

PARAMS = {
    "Rentabilidad": (9.0, 0.60, 0.90), "Rentab. acum.": (12.0, 1.20, 0.90),
    "Duración": (5.0, 0.30, 0.92), "TIR": (3.5, 0.20, 1.05), "Spread": (120.0, 15.0, 0.85),
    "Volatilidad": (12.0, 0.80, 1.10), "Beta": (1.0, 0.05, 1.00), "Peso": (18.0, 3.0, 1.00),
    "TER": (0.85, 0.03, 1.00), "Liquidez": (5.0, 0.10, 1.02),
    "PER": (15.0, 0.80, 1.04), "DividendYield": (2.8, 0.20, 0.95),
}

PERIODOS = ["MTD", "YTD",
            "1M", "2M", "3M", "4M", "5M", "6M",
            "1A", "2A", "3A", "4A", "5A", "6A"]
# Meses que abarca cada periodo (MTD ~1, YTD = meses del año en curso).
_MESES_PERIODO = {"MTD": 1, "YTD": 6, "1M": 1, "2M": 2, "3M": 3, "4M": 4, "5M": 5,
                  "6M": 6, "1A": 12, "2A": 24, "3A": 36, "4A": 48, "5A": 60, "6A": 72}
# Buckets disponibles por granularidad (hay 3 años de datos de ejemplo).
_COUNTS = [36, 12, 6, 3]  # Mensual, Trimestral, Semestral, Anual


def _n_row(meses):
    raw = [meses, math.ceil(meses / 3), math.ceil(meses / 6), math.ceil(meses / 12)]
    return [min(r, c) for r, c in zip(raw, _COUNTS)]  # cap a lo disponible


TABLA_N = {p: _n_row(_MESES_PERIODO[p]) for p in PERIODOS}

AZUL, GRIS = "FF0072CE", "FFF2F2F2"
BOLD = Font(bold=True)
BOLD_WHITE = Font(bold=True, color="FFFFFFFF")


def _valor(metrica, idx):
    base, paso, _ = PARAMS[metrica]
    return round(base + paso * (2.0 * math.sin(idx * 0.6) + (idx % 4) - 1.5), 2)


def build() -> Workbook:
    wb = Workbook()

    # === Datos: Entidad|TipoActivo|Metrica|EjeXTipo|EjeXValor|Serie|Valor ===
    ws_d = wb.active
    ws_d.title = "Datos"
    ws_d.append(["Entidad", "TipoActivo", "Metrica", "EjeXTipo", "EjeXValor", "Serie", "Valor"])
    for c in ws_d[1]:
        c.font = BOLD
    for tipo_ent, entidades in ENTIDADES.items():
        for entidad, tipos in entidades.items():
            for tipo in tipos:
                for met in ALL_METRICS:
                    factor = PARAMS[met][2]
                    for dim in DIMS:
                        for idx, cat in enumerate(CATEGORIAS[dim]):
                            v = _valor(met, idx)
                            ws_d.append([entidad, tipo, met, dim, cat, "Cartera", v])
                            ws_d.append([entidad, tipo, met, dim, cat, "Benchmark", round(v * factor, 2)])
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

    # Entidades por tipo (cascada)
    ent_cols = {"Fondo": "A", "Cartera": "B", "Indice": "C"}
    for tipo_ent, letter in ent_cols.items():
        name(f"Ent_{tipo_ent}", col(letter, f"Ent_{tipo_ent}", list(ENTIDADES[tipo_ent])))

    grp_cols = ["D", "E", "F", "G", "H", "I"]
    for letter, g in zip(grp_cols, GRUPOS):
        name(f"Grupo_{g}", col(letter, f"Grupo_{g}", GRUPOS[g]))

    cat_cols = ["K", "L", "M", "N", "O", "P", "Q", "R", "S", "T"]
    for letter, dim in zip(cat_cols, CATEGORIAS):
        name(f"Cat_{dim}", col(letter, f"Cat_{dim}", CATEGORIAS[dim]))

    name("Periodos", col("V", "Periodos", PERIODOS))
    name("DimTiempo", col("W", "DimTiempo", DIM_TIEMPO))
    ws_l["Y1"], ws_l["Z1"], ws_l["AA1"], ws_l["AB1"] = ("N_Men", "N_Tri", "N_Sem", "N_Anu")
    for r, p in enumerate(PERIODOS, start=2):
        for cidx, val in enumerate(TABLA_N[p]):
            ws_l.cell(row=r, column=25 + cidx, value=val)  # 25 = Y
    name("TablaN", f"Listas!$Y$2:$AB${1 + len(PERIODOS)}")

    # === Panel ===
    ws = wb.create_sheet("Panel", 0)
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Panel genérico de gráficos"
    ws["A1"].font = Font(bold=True, size=14)

    etiquetas = {
        "A3": "Tipo de entidad", "A4": "Entidad", "A5": "Grupo de métrica",
        "A6": "Métrica", "A7": "Dimensión (eje X)", "A8": "Filtro: tipo de activo",
        "A9": "Periodo", "A10": "Benchmark", "A11": "Estilo benchmark",
        "A12": "Tipo de gráfico",
    }
    for celda, txt in etiquetas.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    defaults = {
        "B3": "Fondo", "B4": "Mixto Moderado", "B5": "Riesgo", "B6": "Duración",
        "B7": "Trimestral", "B8": "Todos", "B9": "3A", "B10": "Con benchmark",
        "B11": "Líneas", "B12": "Columnas",
    }
    for celda, val in defaults.items():
        ws[celda] = val
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)

    grupos_lst = ",".join(GRUPOS)
    dims_lst = ",".join(DIMS)
    periodos_lst = ",".join(PERIODOS)
    dvs = [
        ("B3", '"Fondo,Cartera,Indice"'),
        ("B4", '=INDIRECT("Ent_"&$B$3)'),
        ("B5", f'"{grupos_lst}"'),
        ("B6", '=INDIRECT("Grupo_"&$B$5)'),
        ("B7", f'"{dims_lst}"'),
        ("B8", '"Todos,RF,RV"'),
        ("B9", f'"{periodos_lst}"'),
        ("B10", '"Con benchmark,Sin benchmark"'),
        ("B11", '"Barras,Líneas"'),
        ("B12", '"Columnas,Barras,Líneas,Área,Circular,Anillo,Radar"'),
    ]
    for celda, formula in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[celda])

    # Helpers
    ws["A14"] = "Buckets visibles"
    ws["A14"].font = Font(italic=True, size=9)
    ws["B14"] = ('=IFERROR(INDEX(TablaN,MATCH($B$9,Periodos,0),MATCH($B$7,DimTiempo,0)),'
                 'COUNTA(INDIRECT("Cat_"&$B$7)))')
    ws["A15"] = "Criterio tipo activo"
    ws["A15"].font = Font(italic=True, size=9)
    ws["B15"] = '=IF($B$8="Todos","*",$B$8)'

    # Título dinámico
    ws.merge_cells("A17:C17")
    ws["A17"] = ('="Entidad: "&B4&"  ·  "&B6&" por "&B7&"  ·  "&B9'
                 '&IF(B10="Con benchmark","  ·  vs benchmark","")')
    ws["A17"].font = Font(bold=True, size=12, color=AZUL[2:])
    ws["A17"].alignment = Alignment(horizontal="left")

    # Tabla de resultados
    ws["D2"], ws["E2"], ws["F2"] = "Categoría", "Cartera", "Benchmark"
    for celda in ("D2", "E2", "F2"):
        ws[celda].font = BOLD_WHITE
        ws[celda].fill = PatternFill("solid", fgColor=AZUL)

    def sumifs(serie):
        return (
            f'SUMIFS(Datos!$G$2:$G${n},Datos!$A$2:$A${n},$B$4,'
            f'Datos!$B$2:$B${n},$B$15,Datos!$C$2:$C${n},$B$6,'
            f'Datos!$D$2:$D${n},$B$7,Datos!$E$2:$E${n},$D{{r}},'
            f'Datos!$F$2:$F${n},"{serie}")'
        )

    for i in range(MAX_CATS):
        r = 3 + i
        ws.cell(row=r, column=4, value=(
            f'=IF((ROW()-2)>$B$14,"",INDEX(INDIRECT("Cat_"&$B$7),'
            f'COUNTA(INDIRECT("Cat_"&$B$7))-$B$14+(ROW()-2)))'
        ))
        ws.cell(row=r, column=5, value=(
            f'=IF($D{r}="","",{sumifs("Cartera").format(r=r)})'
        ))
        ws.cell(row=r, column=6, value=(
            f'=IF(OR($B$10="Sin benchmark",$D{r}=""),"",{sumifs("Benchmark").format(r=r)})'
        ))
        ws.cell(row=r, column=5).number_format = "#,##0.00"
        ws.cell(row=r, column=6).number_format = "#,##0.00"

    # Gráfico (rango inicial = categorías visibles por defecto; la macro lo hace dinámico)
    vis0 = TABLA_N["3A"][DIM_TIEMPO.index("Trimestral")]  # default Trimestral+3A = 12
    last = 2 + vis0
    chart = BarChart()
    chart.type = "col"
    chart.title = "Cartera vs Benchmark"
    chart.height = 8.5
    chart.width = 18
    chart.y_axis.title = "Duración"
    chart.x_axis.title = "Trimestral"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
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
