// Integración con Excel vía Office.js: escribe los datos en una hoja y crea un
// gráfico NATIVO de Excel (no una imagen). Así el usuario puede ajustarlo y, en
// la Fase 2, se puede pegar en PowerPoint como objeto editable.

import type { ChartType, QueryResult } from "./api";

const CHART_MAP: Record<ChartType, Excel.ChartType | null> = {
  column: Excel.ChartType.columnClustered,
  bar: Excel.ChartType.barClustered,
  line: Excel.ChartType.line,
  area: Excel.ChartType.area,
  pie: Excel.ChartType.pie,
  table: null, // "table" = solo datos, sin gráfico
};

export async function insertChart(
  result: QueryResult,
  chartType: ChartType,
  title: string,
): Promise<void> {
  await Excel.run(async (context) => {
    // Hoja nueva por inserción para no pisar datos existentes.
    const sheet = context.workbook.worksheets.add();
    sheet.activate();

    const values = [result.columns, ...result.rows];
    const rows = values.length;
    const cols = result.columns.length;

    const dataRange = sheet.getRangeByIndexes(0, 0, rows, cols);
    dataRange.values = values as unknown[][];
    sheet.getUsedRange().format.autofitColumns();

    const excelType = CHART_MAP[chartType];
    if (excelType) {
      const chart = sheet.charts.add(excelType, dataRange, Excel.ChartSeriesBy.columns);
      chart.title.text = title;
      chart.setPosition("A1", "H20");
    }

    await context.sync();
  });
}
