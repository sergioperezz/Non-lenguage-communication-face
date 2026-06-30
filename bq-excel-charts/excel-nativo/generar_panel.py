"""Genera un panel interactivo de Excel (Fase 1).

Sin backend, sin add-in, sin macros para lo esencial:
- Desplegables (validación de datos) para Tipo de activo, Métrica (en cascada) y
  Tipo de gráfico.
- Una tabla de resultados calculada con SUMIFS según lo elegido.
- Un gráfico enlazado a esa tabla => cuando el usuario cambia un desplegable,
  el gráfico se actualiza solo.

Datos de ejemplo (en la hoja "Datos"). En la Fase 2 esa hoja se sustituye por
una consulta de Power Query contra BigQuery (vía el driver ODBC), parametrizada
por las mismas celdas, y el resto del panel no cambia.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# --- Datos de ejemplo -------------------------------------------------------
SEGMENTOS = ["Europa", "EEUU", "Asia", "Emergentes", "Global"]

# {tipo_activo: {metrica: [valores por segmento]}}
DATOS = {
    "RF": {
        "Duration": [4.2, 6.1, 5.3, 3.8, 5.0],       # años
        "TIR": [3.1, 4.5, 4.0, 6.2, 4.1],            # %
        "Spread": [90, 120, 110, 250, 130],          # pb
    },
    "RV": {
        "Rentabilidad": [8.5, 11.2, 9.0, 12.5, 10.1],   # %
        "PER": [14.2, 19.5, 15.0, 12.0, 16.3],
        "DividendYield": [3.4, 1.8, 2.6, 3.1, 2.5],     # %
    },
}

AZUL = "FF0072CE"      # azul corporativo aproximado
GRIS = "FFF2F2F2"
BOLD = Font(bold=True)
BOLD_WHITE = Font(bold=True, color="FFFFFFFF")


def build() -> Workbook:
    wb = Workbook()

    # === Hoja Datos (formato largo: TipoActivo | Metrica | Segmento | Valor) ===
    ws_datos = wb.active
    ws_datos.title = "Datos"
    ws_datos.append(["TipoActivo", "Metrica", "Segmento", "Valor"])
    for c in ws_datos[1]:
        c.font = BOLD
    for tipo, metricas in DATOS.items():
        for metrica, valores in metricas.items():
            for seg, val in zip(SEGMENTOS, valores):
                ws_datos.append([tipo, metrica, seg, val])

    # === Hoja Listas (para el desplegable de métrica en cascada) ===
    ws_listas = wb.create_sheet("Listas")
    ws_listas["A1"] = "RF"
    ws_listas["B1"] = "RV"
    for i, m in enumerate(DATOS["RF"], start=2):
        ws_listas.cell(row=i, column=1, value=m)
    for i, m in enumerate(DATOS["RV"], start=2):
        ws_listas.cell(row=i, column=2, value=m)
    # Rangos con nombre RF y RV -> los usa INDIRECT para la cascada.
    n_rf = len(DATOS["RF"]) + 1
    n_rv = len(DATOS["RV"]) + 1
    wb.defined_names.add(DefinedName("RF", attr_text=f"Listas!$A$2:$A${n_rf}"))
    wb.defined_names.add(DefinedName("RV", attr_text=f"Listas!$B$2:$B${n_rv}"))

    # === Hoja Panel (lo que usa el analista) ===
    ws = wb.create_sheet("Panel", 0)  # primera pestaña
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Panel interactivo de gráficos"
    ws["A1"].font = Font(bold=True, size=14)

    # Controles
    etiquetas = {"A3": "Tipo de activo", "A4": "Métrica", "A5": "Tipo de gráfico"}
    for celda, txt in etiquetas.items():
        ws[celda] = txt
        ws[celda].font = BOLD
    ws["B3"] = "RF"
    ws["B4"] = "Duration"
    ws["B5"] = "Columnas"
    for celda in ("B3", "B4", "B5"):
        ws[celda].fill = PatternFill("solid", fgColor=GRIS)

    # Validaciones (desplegables)
    dv_tipo = DataValidation(type="list", formula1='"RF,RV"', allow_blank=False)
    dv_metrica = DataValidation(type="list", formula1="=INDIRECT($B$3)", allow_blank=False)
    dv_grafico = DataValidation(
        type="list",
        formula1='"Columnas,Barras,Líneas,Área,Circular,Anillo,Radar"',
        allow_blank=False,
    )
    ws.add_data_validation(dv_tipo)
    ws.add_data_validation(dv_metrica)
    ws.add_data_validation(dv_grafico)
    dv_tipo.add(ws["B3"])
    dv_metrica.add(ws["B4"])
    dv_grafico.add(ws["B5"])

    # Título dinámico (acotado a A:C para no pisar la tabla de resultados en D:E)
    ws.merge_cells("A7:C7")
    ws["A7"] = '="Métrica: "&B4&"   ·   Activo: "&B3'
    ws["A7"].font = Font(bold=True, size=12, color=AZUL[2:])
    ws["A7"].alignment = Alignment(horizontal="left")

    # Tabla de resultados (SUMIFS según los desplegables)
    ws["D3"] = "Segmento"
    ws["E3"] = "Valor"
    ws["D3"].font = BOLD_WHITE
    ws["E3"].font = BOLD_WHITE
    ws["D3"].fill = PatternFill("solid", fgColor=AZUL)
    ws["E3"].fill = PatternFill("solid", fgColor=AZUL)
    for i, seg in enumerate(SEGMENTOS):
        r = 4 + i
        ws.cell(row=r, column=4, value=seg)
        ws.cell(
            row=r,
            column=5,
            value=(
                f"=SUMIFS(Datos!$D:$D,Datos!$A:$A,$B$3,"
                f"Datos!$B:$B,$B$4,Datos!$C:$C,D{r})"
            ),
        )

    # Gráfico enlazado a la tabla -> se actualiza solo al cambiar los desplegables.
    chart = BarChart()
    chart.type = "col"  # vertical (ver README para alternar a horizontal)
    chart.title = "Valor por segmento"
    chart.legend = None
    chart.height = 8
    chart.width = 15
    last = 3 + len(SEGMENTOS)
    data = Reference(ws, min_col=5, min_row=3, max_row=last)      # E3:E8 (con cabecera)
    cats = Reference(ws, min_col=4, min_row=4, max_row=last)      # D4:D8
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "G3")

    # Anchos
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 12

    return wb


if __name__ == "__main__":
    out = Path(__file__).parent / "Panel_Interactivo.xlsx"
    build().save(out)
    print(f"Generado: {out}")
