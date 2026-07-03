"""Genera el panel GENÉRICO (Fase 1) — con comparación de varias entidades.

Motor guiado por datos. Jerarquía:

  Tipo de entidad   -> Fondo / Cartera / Indice
  Entidad 1         -> cascada del tipo (principal)
  Entidad 2, 3      -> opcionales, para comparar varias a la vez
  Grupo de métrica  -> familia; Métrica -> cascada del grupo
  Dimensión (eje X) -> tiempo o composición
  Filtro Tipo activo-> Todos/RF/RV
  Periodo           -> MTD/YTD/1M..6M/1A..6A
  Benchmark         -> Con/Sin (de la Entidad 1) · Estilo Barras/Líneas
  Tipo de gráfico   -> representación

Cada entidad seleccionada es una serie -> permite "cartera A vs cartera B".
Datos de ejemplo (ficticios). En la Fase 2 vienen de BigQuery por ODBC.
"""

from __future__ import annotations

import math
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation


# === Modelo ================================================================
ENTIDADES = {
    "Fondo": {
        "RF Privada A": ["RF"],
        "Bolsa Europa": ["RV"],
        "Mixto Moderado": ["RF", "RV"],
        "Monetario Euro": ["RF"],
    },
    "Cartera": {
        "Cartera RF Gobierno": ["RF"],
        "Cartera RF Crédito": ["RF"],
        "Cartera Conservadora": ["RF", "RV"],
        "Cartera Dinamica": ["RF", "RV"],
    },
    "Indice": {
        "Euro Stoxx 50": ["RV"],
        "Bloomberg Agg": ["RF"],
        "MSCI World": ["RV"],
        "S&P 500": ["RV"],
    },
}

GRUPOS = {
    "Rendimiento": ["Rentabilidad", "Rentab. acum."],
    "Riesgo": ["Duración Modificada", "Duración Efectiva", "TIR", "Spread",
               "Volatilidad", "Beta"],
    "Composicion": ["Peso"],
    "Costes": ["TER", "TER Look-through"],
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

PARAMS = {
    "Rentabilidad": (9.0, 0.60, 0.90), "Rentab. acum.": (12.0, 1.20, 0.90),
    "Duración Modificada": (5.0, 0.30, 0.92), "Duración Efectiva": (5.2, 0.30, 0.92),
    "TIR": (3.5, 0.20, 1.05), "Spread": (120.0, 15.0, 0.85),
    "Volatilidad": (12.0, 0.80, 1.10), "Beta": (1.0, 0.05, 1.00), "Peso": (18.0, 3.0, 1.00),
    "TER": (0.85, 0.03, 1.00), "TER Look-through": (0.95, 0.03, 1.00),
    "Liquidez": (5.0, 0.10, 1.02),
    "PER": (15.0, 0.80, 1.04), "DividendYield": (2.8, 0.20, 0.95),
}

PERIODOS = ["MTD", "YTD", "1M", "2M", "3M", "4M", "5M", "6M",
            "1A", "2A", "3A", "4A", "5A", "6A"]
_MESES_PERIODO = {"MTD": 1, "YTD": 6, "1M": 1, "2M": 2, "3M": 3, "4M": 4, "5M": 5,
                  "6M": 6, "1A": 12, "2A": 24, "3A": 36, "4A": 48, "5A": 60, "6A": 72}
_COUNTS = [36, 12, 6, 3]


def _n_row(meses):
    raw = [meses, math.ceil(meses / 3), math.ceil(meses / 6), math.ceil(meses / 12)]
    return [min(r, c) for r, c in zip(raw, _COUNTS)]


TABLA_N = {p: _n_row(_MESES_PERIODO[p]) for p in PERIODOS}

# Índice de la entidad (columna de datos) y offset por serie/entidad para variar
# los valores ficticios entre entidades (para que la comparación no salga igual).
_ENT_LIST = [e for ents in ENTIDADES.values() for e in ents]

AZUL, GRIS = "FF0072CE", "FFF2F2F2"
BOLD = Font(bold=True)
BOLD_WHITE = Font(bold=True, color="FFFFFFFF")


def _valor(metrica, idx, ent_off=0):
    base, paso, _ = PARAMS[metrica]
    return round(base + paso * (2.0 * math.sin((idx + ent_off) * 0.6) + (idx % 4) - 1.5)
                 + ent_off * paso * 0.4, 2)


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
            ent_off = _ENT_LIST.index(entidad)              # varía entre entidades
            for tipo in tipos:
                for met in ALL_METRICS:
                    factor = PARAMS[met][2]
                    for dim in DIMS:
                        for idx, cat in enumerate(CATEGORIAS[dim]):
                            v = _valor(met, idx, ent_off)
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

    for tipo_ent, letter in {"Fondo": "A", "Cartera": "B", "Indice": "C"}.items():
        name(f"Ent_{tipo_ent}", col(letter, f"Ent_{tipo_ent}", list(ENTIDADES[tipo_ent])))
    # Listas para los comparadores opcionales: "(ninguna)" + entidades.
    for tipo_ent, letter in {"Fondo": "AD", "Cartera": "AE", "Indice": "AF"}.items():
        name(f"Ent_{tipo_ent}_opt",
             col(letter, f"Ent_{tipo_ent}_opt", ["(ninguna)"] + list(ENTIDADES[tipo_ent])))

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
        "A3": "Tipo de entidad", "A4": "Entidad 1", "A5": "Entidad 2 (opc.)",
        "A6": "Entidad 3 (opc.)", "A7": "Grupo de métrica", "A8": "Métrica",
        "A9": "Dimensión (eje X)", "A10": "Filtro: tipo de activo", "A11": "Periodo",
        "A12": "Benchmark", "A13": "Estilo benchmark", "A14": "Tipo de gráfico",
    }
    for celda, txt in etiquetas.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    defaults = {
        "B3": "Fondo", "B4": "RF Privada A", "B5": "(ninguna)", "B6": "(ninguna)", "B7": "Riesgo",
        "B8": "Duración Modificada", "B9": "Trimestral", "B10": "Todos", "B11": "3A",
        "B12": "Con benchmark", "B13": "Líneas", "B14": "Columnas",
    }
    for celda, val in defaults.items():
        ws[celda] = val
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)

    grupos_lst = ",".join(GRUPOS)
    dims_lst = ",".join(DIMS)
    periodos_lst = ",".join(PERIODOS)
    dvs = [
        ("B3", '"Fondo,Cartera,Indice"', False),
        ("B4", '=INDIRECT("Ent_"&$B$3)', False),
        ("B5", '=INDIRECT("Ent_"&$B$3&"_opt")', True),   # opcional: incluye "(ninguna)"
        ("B6", '=INDIRECT("Ent_"&$B$3&"_opt")', True),   # opcional
        ("B7", f'"{grupos_lst}"', False),
        ("B8", '=INDIRECT("Grupo_"&$B$7)', False),
        ("B9", f'"{dims_lst}"', False),
        ("B10", '"Todos,RF,RV"', False),
        ("B11", f'"{periodos_lst}"', False),
        ("B12", '"Con benchmark,Sin benchmark"', False),
        ("B13", '"Barras,Líneas,Puntos"', False),
        ("B14", '"Columnas,Barras,Líneas,Área,Circular,Anillo,Radar,Apiladas,100% apiladas"', False),
    ]
    for celda, formula, blank in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=blank)
        ws.add_data_validation(dv)
        dv.add(ws[celda])

    # Helpers
    ws["A16"] = "Buckets visibles"
    ws["A16"].font = Font(italic=True, size=9)
    ws["B16"] = ('=IFERROR(INDEX(TablaN,MATCH($B$11,Periodos,0),MATCH($B$9,DimTiempo,0)),'
                 'COUNTA(INDIRECT("Cat_"&$B$9)))')
    ws["A17"] = "Criterio tipo activo"
    ws["A17"].font = Font(italic=True, size=9)
    ws["B17"] = '=IF($B$10="Todos","*",$B$10)'

    # Título dinámico
    ws.merge_cells("A19:C19")
    ws["A19"] = ('="Entidad: "&B4&IF(B5<>""," vs "&B5,"")&IF(B6<>""," vs "&B6,"")'
                 '&"  ·  "&B8&" por "&B9&"  ·  "&B11')
    ws["A19"].font = Font(bold=True, size=12, color=AZUL[2:])
    ws["A19"].alignment = Alignment(horizontal="left")

    # Tabla de resultados: Categoría | Ent1 | Ent2 | Ent3 | Benchmark
    ws["D2"] = "Categoría"
    ws["E2"] = "=B4"
    ws["F2"] = '=IF(B5="(ninguna)","",B5)'
    ws["G2"] = '=IF(B6="(ninguna)","",B6)'
    ws["H2"] = "Benchmark"
    for celda in ("D2", "E2", "F2", "G2", "H2"):
        ws[celda].font = BOLD_WHITE
        ws[celda].fill = PatternFill("solid", fgColor=AZUL)

    def sumifs(ent, serie):
        return (f'SUMIFS(Datos!$G$2:$G${n},Datos!$A$2:$A${n},{ent},'
                f'Datos!$B$2:$B${n},$B$17,Datos!$C$2:$C${n},$B$8,'
                f'Datos!$D$2:$D${n},$B$9,Datos!$E$2:$E${n},$D{{r}},'
                f'Datos!$F$2:$F${n},"{serie}")')

    for i in range(MAX_CATS):
        r = 3 + i
        ws.cell(row=r, column=4, value=(
            f'=IF((ROW()-2)>$B$16,"",INDEX(INDIRECT("Cat_"&$B$9),'
            f'COUNTA(INDIRECT("Cat_"&$B$9))-$B$16+(ROW()-2)))'
        ))
        ws.cell(row=r, column=5, value=(
            f'=IF(OR($B$4="",$D{r}=""),"",{sumifs("$B$4", "Cartera").format(r=r)})'))
        ws.cell(row=r, column=6, value=(
            f'=IF(OR($B$5="",$B$5="(ninguna)",$D{r}=""),"",{sumifs("$B$5", "Cartera").format(r=r)})'))
        ws.cell(row=r, column=7, value=(
            f'=IF(OR($B$6="",$B$6="(ninguna)",$D{r}=""),"",{sumifs("$B$6", "Cartera").format(r=r)})'))
        ws.cell(row=r, column=8, value=(
            f'=IF(OR($B$12="Sin benchmark",$B$4="",$D{r}=""),"",'
            f'{sumifs("$B$4", "Benchmark").format(r=r)})'))
        for cc in (5, 6, 7, 8):
            ws.cell(row=r, column=cc).number_format = "#,##0.00"

    # Gráfico inicial (Ent1 + Benchmark; la macro lo hace dinámico y multi-serie)
    vis0 = TABLA_N["3A"][DIM_TIEMPO.index("Trimestral")]
    last = 2 + vis0
    chart = BarChart()
    chart.type = "col"
    chart.title = "Comparación"
    chart.height = 8.5
    chart.width = 18
    chart.y_axis.title = "Duración"
    chart.x_axis.title = "Trimestral"
    chart.y_axis.delete = False
    chart.x_axis.delete = False
    chart.add_data(Reference(ws, min_col=5, min_row=2, max_row=last), titles_from_data=True)  # Ent1
    chart.add_data(Reference(ws, min_col=8, min_row=2, max_row=last), titles_from_data=True)  # Benchmark
    chart.set_categories(Reference(ws, min_col=4, min_row=3, max_row=last))
    ws.add_chart(chart, "H2")

    for c, w in {"A": 20, "B": 22, "D": 16, "E": 14, "F": 14, "G": 14, "H": 12}.items():
        ws.column_dimensions[c].width = w

    # Vista previa de la SQL (Fase 2). La macro la escribe en A41. Ocupa SOLO las
    # columnas A:C (no D:H) para no chocar con la tabla del gráfico al volcar datos.
    ws["A40"] = "SQL generada (Fase 2 · BigQuery):"
    ws["A40"].font = Font(bold=True, size=10, color=AZUL[2:])
    ws.merge_cells("A41:C60")
    sqlc = ws["A41"]
    sqlc.value = ("« Pega la macro (macro_Panel.vba), guarda como .xlsm y añade un "
                  "botón con la macro 'Actualizar'. Al pulsarlo, aquí verás la SQL, "
                  "en la columna W el resultado y el gráfico ya dibujado. »")
    sqlc.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    sqlc.font = Font(name="Consolas", size=9)
    sqlc.fill = PatternFill("solid", fgColor="F7F9FC")

    build_tablas(wb, n)
    build_sectorial(wb)
    build_comparativa(wb)
    return wb


def build_comparativa(wb):
    """Hoja 'Comparativa': eje X = LISTA DE CARTERAS, con varias versiones por
    cartera en columnas AGRUPADAS (Gestionada / IMP / Benchmark) + una tabla con
    columnas 'diff' (cartera - benchmark) en verde/rojo. Reproduce el caso
    'Comportamiento de las Carteras' (DPM Global Managers) de los decks. Es un
    corte transversal: las entidades son las categorías, no las series. Su tabla
    se calcula con SUMIFS (hoja 'DatosComp') -> sin macro."""
    VERSIONES = ["Gestionada", "IMP", "Benchmark"]

    # Datos de ejemplo (Entidad | Metrica | Version | Valor): un valor escalar por
    # cartera y versión (snapshot del periodo, como el "MTD return" del deck).
    ws_d = wb.create_sheet("DatosComp")
    ws_d.append(["Entidad", "Metrica", "Version", "Valor"])
    for c in ws_d[1]:
        c.font = BOLD
    for ei, ent in enumerate(_ENT_LIST):
        for met in ALL_METRICS:
            base = _valor(met, 0, ei)
            factor = PARAMS[met][2]
            vals = {
                "Gestionada": base,
                "IMP": round(base * (1.0 + 0.06 * math.sin(ei + 1)), 2),
                "Benchmark": round(base * factor, 2),
            }
            for ver in VERSIONES:
                ws_d.append([ent, met, ver, vals[ver]])
    n3 = ws_d.max_row

    ws = wb.create_sheet("Comparativa")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Comparativa de carteras (eje X = carteras) + columnas 'diff'"
    ws["A1"].font = Font(bold=True, size=13)

    for celda, txt in {"A3": "Tipo de entidad", "A4": "Métrica"}.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    ws["B3"], ws["B4"] = "Cartera", "Rentabilidad"
    for celda in ("B3", "B4"):
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)
    dvs = [
        ("B3", '"Fondo,Cartera,Indice"'),
        ("B4", f'"{",".join(ALL_METRICS)}"'),
    ]
    for celda, formula in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[celda])

    # Nº máximo de entidades de un tipo (Cartera = 4).
    NROW = max(len(e) for e in ENTIDADES.values())
    HDR = 6
    headers = ["Cartera"] + VERSIONES + ["Dif. vs Bench", "Dif. IMP vs Bench"]
    for j, h in enumerate(headers):
        c = ws.cell(HDR, 1 + j, h)
        c.font = BOLD_WHITE
        c.fill = PatternFill("solid", fgColor=AZUL)
    for i in range(NROW):
        r = HDR + 1 + i
        ws.cell(r, 1, f'=IFERROR(INDEX(INDIRECT("Ent_"&$B$3),{i + 1}),"")').font = BOLD
        for j, ver in enumerate(VERSIONES):
            ws.cell(r, 2 + j, (
                f'=IF($A{r}="","",SUMIFS(DatosComp!$D$2:$D${n3},'
                f'DatosComp!$A$2:$A${n3},$A{r},DatosComp!$B$2:$B${n3},$B$4,'
                f'DatosComp!$C$2:$C${n3},"{ver}"))'
            ))
            ws.cell(r, 2 + j).number_format = "#,##0.00"
        # Columnas diff: Gestionada - Benchmark, IMP - Benchmark.
        ws.cell(r, 5, f'=IF($A{r}="","",B{r}-D{r})')
        ws.cell(r, 6, f'=IF($A{r}="","",C{r}-D{r})')
        for cc in (5, 6):
            ws.cell(r, cc).number_format = "+#,##0.00;-#,##0.00;0.00"
    last = HDR + NROW

    # Columnas AGRUPADAS: una serie por versión, categorías = carteras.
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    bar.title = "Comportamiento de las carteras"
    bar.height, bar.width = 9, 19
    bar.add_data(Reference(ws, min_col=2, max_col=1 + len(VERSIONES),
                           min_row=HDR, max_row=last), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=HDR + 1, max_row=last))
    bar.y_axis.title = "Valor"
    bar.x_axis.title = "Cartera"
    bar.y_axis.delete = False
    bar.x_axis.delete = False
    ws.add_chart(bar, "I3")

    # Diff en verde (positivo) / rojo (negativo), como el "diff Return" del deck.
    diff_rng = f"E{HDR + 1}:F{last}"
    ws.conditional_formatting.add(diff_rng, CellIsRule(
        operator="greaterThan", formula=["0"], font=Font(color="FF107C41")))
    ws.conditional_formatting.add(diff_rng, CellIsRule(
        operator="lessThan", formula=["0"], font=Font(color="FFC00000")))

    ws.column_dimensions["A"].width = 22
    for cl in ("B", "C", "D", "E", "F"):
        ws.column_dimensions[cl].width = 15
    return ws


def build_sectorial(wb):
    """Hoja 'Sectorial': columnas APILADAS (componentes por categoría) + serie de
    puntos en EJE SECUNDARIO (Contribución al Tracking Error). Ahora se puede
    AGRUPAR POR varias dimensiones (Sector, Industria, Geografia, Divisa, Rating,
    Activo). Combo con openpyxl enlazado a una tabla SUMIFS -> sin macro."""
    GROUP_DIMS = ["Sector", "Industria", "Geografia", "Divisa", "Rating", "Activo"]
    COMPONENTES = ["Large Cap", "Mid Cap", "Small Cap"]
    MAXCAT = max(len(CATEGORIAS[d]) for d in GROUP_DIMS)  # filas de la tabla/chart

    # Datos de ejemplo (Entidad | Dim | Categoria | Serie | Valor).
    ws_d = wb.create_sheet("DatosSector")
    ws_d.append(["Entidad", "Dim", "Categoria", "Serie", "Valor"])
    for c in ws_d[1]:
        c.font = BOLD
    for ei, ent in enumerate(_ENT_LIST):
        for dim in GROUP_DIMS:
            for si, cat in enumerate(CATEGORIAS[dim]):
                for ci, comp in enumerate(COMPONENTES):
                    v = round(0.15 + 0.10 * ((si + ci) % 3)
                              + 0.05 * abs(math.sin(ei + si + ci)), 3)
                    ws_d.append([ent, dim, cat, comp, v])
                te = round(0.004 + 0.003 * abs(math.sin(si + ei * 0.5)), 4)
                ws_d.append([ent, dim, cat, "Tracking Error", te])
    n2 = ws_d.max_row

    # Lista de todas las entidades para el desplegable.
    ws_l = wb["Listas"]
    ws_l.cell(row=1, column=34, value="EntidadesAll").font = BOLD  # col AH
    for i, e in enumerate(_ENT_LIST, start=2):
        ws_l.cell(row=i, column=34, value=e)
    wb.defined_names.add(DefinedName("EntidadesAll",
                                     attr_text=f"Listas!$AH$2:$AH${1 + len(_ENT_LIST)}"))

    # Fase 2: mapa Nombre -> PK_PORTFOLIO_ID (la macro filtra por ID). Se usan como
    # EJEMPLO los PK_PORTFOLIO_ID reales del entorno; reconcilia cada nombre visible
    # con su ID real (los nombres del mock son ficticios).
    IDS_REALES = ["CBNKITER", "DIVERDIN", "FALBUSFI", "GESTIO30", "GESTIO60",
                  "GESTOTAL", "IMP030RV", "IMP060RV", "IMP100RV", "MIXTRF10",
                  "MIXTRF15", "PRO030RV", "PRO060RV", "PRO100RV"]
    ws_l.cell(row=1, column=36, value="Entidad").font = BOLD          # col AJ
    ws_l.cell(row=1, column=37, value="PK_PORTFOLIO_ID").font = BOLD  # col AK
    for i, e in enumerate(_ENT_LIST, start=2):
        ws_l.cell(row=i, column=36, value=e)
        ws_l.cell(row=i, column=37, value=IDS_REALES[(i - 2) % len(IDS_REALES)])
    wb.defined_names.add(DefinedName("MapaEntidades",
                                     attr_text=f"Listas!$AJ$2:$AK${1 + len(_ENT_LIST)}"))

    ws = wb.create_sheet("Sectorial")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Distribución por componentes (apilado) + Tracking Error (eje secundario)"
    ws["A1"].font = Font(bold=True, size=13)
    for celda, txt in {"A3": "Entidad", "A4": "Agrupar por"}.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    ws["B3"] = _ENT_LIST[0]
    ws["B4"] = "Sector"
    for celda in ("B3", "B4"):
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)
    dv1 = DataValidation(type="list", formula1="=EntidadesAll", allow_blank=False)
    ws.add_data_validation(dv1)
    dv1.add(ws["B3"])
    dv2 = DataValidation(type="list", formula1=f'"{",".join(GROUP_DIMS)}"', allow_blank=False)
    ws.add_data_validation(dv2)
    dv2.add(ws["B4"])

    HDR = 6
    hc = ws.cell(HDR, 1, "Categoría")
    hc.font = BOLD_WHITE
    hc.fill = PatternFill("solid", fgColor=AZUL)
    headers = COMPONENTES + ["Tracking Error"]
    for j, h in enumerate(headers):
        c = ws.cell(HDR, 2 + j, h)
        c.font = BOLD_WHITE
        c.fill = PatternFill("solid", fgColor=AZUL)
    for i in range(MAXCAT):
        r = HDR + 1 + i
        # Categoría i-ésima de la dimensión elegida (vacía si la dim tiene menos).
        ws.cell(r, 1, f'=IFERROR(INDEX(INDIRECT("Cat_"&$B$4),{i + 1}),"")')
        for j, h in enumerate(headers):
            ws.cell(r, 2 + j, (
                f'=IF($A{r}="","",SUMIFS(DatosSector!$E$2:$E${n2},'
                f'DatosSector!$A$2:$A${n2},$B$3,DatosSector!$B$2:$B${n2},$B$4,'
                f'DatosSector!$C$2:$C${n2},$A{r},DatosSector!$D$2:$D${n2},"{h}"))'
            ))
            ws.cell(r, 2 + j).number_format = "0.000"
    last = HDR + MAXCAT

    # Columnas apiladas (los componentes)
    bar = BarChart()
    bar.type = "col"
    bar.grouping = "stacked"
    bar.overlap = 100
    bar.title = "Distribución por componentes + Tracking Error"
    bar.height, bar.width = 9, 19
    bar.add_data(Reference(ws, min_col=2, max_col=1 + len(COMPONENTES),
                           min_row=HDR, max_row=last), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=HDR + 1, max_row=last))
    bar.y_axis.title = "Peso activo (%)"

    # Serie de puntos (Tracking Error) en eje secundario
    line = LineChart()
    line.add_data(Reference(ws, min_col=2 + len(COMPONENTES), max_col=2 + len(COMPONENTES),
                            min_row=HDR, max_row=last), titles_from_data=True)
    s = line.series[0]
    s.marker = Marker(symbol="circle", size=7)
    s.graphicalProperties = GraphicalProperties()
    s.graphicalProperties.line = LineProperties(noFill=True)  # solo marcadores, sin línea
    line.y_axis.axId = 200
    line.y_axis.title = "Contrib. Tracking Error"
    bar.y_axis.crosses = "autoZero"
    line.y_axis.crosses = "max"
    bar += line
    ws.add_chart(bar, "H3")

    ws.column_dimensions["A"].width = 16
    return ws


# Layout de la hoja Tablas (compartido con la macro FormatearTablas de VBA).
TAB_HDR = 17            # fila de cabecera de categorías
TAB_NCOL = 12           # nº máx. de columnas de categoría (B..M)
TAB_ROW0 = TAB_HDR + 1  # primera fila de datos (18)


def build_tablas(wb, n):
    """Hoja 'Tablas': generador CONFIGURABLE (con macro). Selectores de tipo de
    entidad, métrica, dimensión, serie, filtro tipo activo, periodo, estilo y
    decimales. Los valores salen con SUMIFS; la macro `FormatearTablas` ajusta
    columnas visibles, decimales y el estilo de formato condicional."""
    NROW = max(len(e) for e in ENTIDADES.values())
    HDR, NCOL, ROW0 = TAB_HDR, TAB_NCOL, TAB_ROW0
    last_row = HDR + NROW
    last_col = 1 + NCOL
    last_col_l = get_column_letter(last_col)

    ws = wb.create_sheet("Tablas")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Generador de tablas"
    ws["A1"].font = Font(bold=True, size=14)

    etiquetas = {
        "A3": "Tipo de entidad", "A4": "Métrica", "A5": "Dimensión (columnas)",
        "A6": "Serie", "A7": "Filtro: tipo de activo", "A8": "Periodo",
        "A9": "Estilo", "A10": "Decimales",
    }
    for celda, txt in etiquetas.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    defaults = {
        "B3": "Cartera", "B4": "Rentabilidad", "B5": "Anual", "B6": "Cartera",
        "B7": "Todos", "B8": "3A", "B9": "Mapa de calor", "B10": 2,
    }
    for celda, val in defaults.items():
        ws[celda] = val
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)

    dvs = [
        ("B3", '"Fondo,Cartera,Indice"'),
        ("B4", f'"{",".join(ALL_METRICS)}"'),
        ("B5", f'"{",".join(DIMS)}"'),
        ("B6", '"Cartera,Benchmark"'),
        ("B7", '"Todos,RF,RV"'),
        ("B8", f'"{",".join(PERIODOS)}"'),
        ("B9", '"Mapa de calor,Barras de datos,Signos +/-,Sin formato"'),
        ("B10", '"0,1,2"'),
    ]
    for celda, formula in dvs:
        dv = DataValidation(type="list", formula1=formula, allow_blank=False)
        ws.add_data_validation(dv)
        dv.add(ws[celda])

    # Helpers (como en el Panel): buckets visibles y criterio de tipo de activo.
    ws["A12"] = "Buckets visibles"
    ws["A12"].font = Font(italic=True, size=9)
    ws["B12"] = ('=IFERROR(INDEX(TablaN,MATCH($B$8,Periodos,0),MATCH($B$5,DimTiempo,0)),'
                 'COUNTA(INDIRECT("Cat_"&$B$5)))')
    ws["A13"] = "Criterio tipo activo"
    ws["A13"].font = Font(italic=True, size=9)
    ws["B13"] = '=IF($B$7="Todos","*",$B$7)'

    # Subtítulo dinámico.
    ws.merge_cells("A15:C15")
    ws["A15"] = ('=$B$3&": "&$B$4&"  ·  por "&$B$5&IF($B$7<>"Todos"," ("&$B$7&")","")'
                 '&"  ·  "&$B$8&"  ·  "&$B$6')
    ws["A15"].font = Font(bold=True, size=12, color=AZUL[2:])

    # Cabecera de categorías (muestra los últimos "buckets visibles" de la dimensión).
    ws.cell(HDR, 1, "Entidad \\ Categoría").font = BOLD_WHITE
    ws.cell(HDR, 1).fill = PatternFill("solid", fgColor=AZUL)
    for j in range(NCOL):
        k = j + 1
        c = ws.cell(HDR, 1 + k, (
            f'=IF({k}>$B$12,"",INDEX(INDIRECT("Cat_"&$B$5),'
            f'COUNTA(INDIRECT("Cat_"&$B$5))-$B$12+{k}))'
        ))
        c.font = BOLD_WHITE
        c.fill = PatternFill("solid", fgColor=AZUL)
        c.alignment = Alignment(horizontal="center")

    thin = Side(style="thin", color="D6DEE8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for i in range(NROW):
        r = ROW0 + i
        ec = ws.cell(r, 1, f'=IFERROR(INDEX(INDIRECT("Ent_"&$B$3),{i + 1}),"")')
        ec.font = BOLD
        ec.border = border
        for j in range(NCOL):
            cc = 2 + j
            cl = get_column_letter(cc)
            cell = ws.cell(r, cc, (
                f'=IF(OR($A{r}="",{cl}${HDR}=""),"",'
                f'SUMIFS(Datos!$G$2:$G${n},Datos!$A$2:$A${n},$A{r},'
                f'Datos!$B$2:$B${n},$B$13,Datos!$C$2:$C${n},$B$4,'
                f'Datos!$D$2:$D${n},$B$5,Datos!$E$2:$E${n},{cl}${HDR},'
                f'Datos!$F$2:$F${n},$B$6))'
            ))
            cell.number_format = "0.00"
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

    # Formato condicional por defecto (Mapa de calor). La macro lo cambia según B9.
    body = f"B{ROW0}:{last_col_l}{last_row}"
    ws.conditional_formatting.add(body, ColorScaleRule(
        start_type="min", start_color="F8696B",
        mid_type="percentile", mid_value=50, mid_color="FFEB84",
        end_type="max", end_color="63BE7B"))

    ws["A" + str(last_row + 2)] = ("Consejo: al abrir la pestaña o cambiar un desplegable, "
                                   "la macro ajusta columnas, decimales y estilo.")
    ws["A" + str(last_row + 2)].font = Font(italic=True, size=9, color="808080")

    ws.column_dimensions["A"].width = 24
    for j in range(NCOL):
        ws.column_dimensions[get_column_letter(2 + j)].width = 11
    return ws


if __name__ == "__main__":
    out = Path(__file__).parent / "Panel_Generico.xlsx"
    build().save(out)
    print(f"Generado: {out}  (máx categorías: {MAX_CATS})")
