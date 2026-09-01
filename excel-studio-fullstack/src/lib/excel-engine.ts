import ExcelJS from 'exceljs';
import { Product, Sheet, Column, KPI, ThemeName } from '../types/schema';
import { PALETTES } from './palettes';
import { injectCharts, quoteSheetRef, ChartPlacement } from './chart-injector';

export class ExcelProductEngine {
  private product: Product;
  private pendingCharts: ChartPlacement[] = [];

  constructor(product: Product) {
    this.product = product;
  }

  /**
   * Converts a 0-based column index to its Excel letter (0 -> A, 25 -> Z,
   * 26 -> AA, ...). String.fromCharCode(65 + idx) only works up to index 25
   * and silently produces invalid characters beyond it — this handles any
   * sheet width correctly.
   */
  private static columnIndexToLetter(index: number): string {
    let n = index + 1;
    let letters = '';
    while (n > 0) {
      const rem = (n - 1) % 26;
      letters = String.fromCharCode(65 + rem) + letters;
      n = Math.floor((n - 1) / 26);
    }
    return letters;
  }

  private static resolveColumnRef(columns: Column[], keyOrLabel: string): { index: number; letter: string } | null {
    const index = columns.findIndex((c) => c.key === keyOrLabel || c.label === keyOrLabel);
    if (index === -1) return null;
    return { index, letter: ExcelProductEngine.columnIndexToLetter(index) };
  }

  /**
   * Sheet.rows is typed string[][], but rows built from the AI-import path
   * (api/ai/route.ts) come from an LLM's JSON response that isn't validated
   * against that type — a cell can arrive as a number, boolean, or null at
   * runtime. Every parser below calls .replace()/.trim() on the raw value,
   * which throws if it isn't actually a string, crashing the whole export
   * over what should just be a coercible value.
   */
  private static toCellString(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number' || typeof value === 'boolean') return String(value);
    try {
      return JSON.stringify(value);
    } catch {
      return String(value);
    }
  }

  public async generateWorkbook(): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();
    workbook.creator = this.product.author || 'Excel Product Studio';
    workbook.lastModifiedBy = this.product.author || 'Excel Product Studio';
    workbook.created = new Date();
    workbook.modified = new Date();
    workbook.title = this.product.name;

    const palette = PALETTES[this.product.theme] || PALETTES.premium;
    const usedSheetNames = new Set<string>();

    for (const sheetData of this.product.sheets) {
      const safeSheetName = ExcelProductEngine.sanitizeSheetName(sheetData.name, usedSheetNames);
      this.buildSheet(workbook, sheetData, palette, safeSheetName);
    }

    return workbook;
  }

  /**
   * Excel worksheet names must be non-empty, <= 31 chars, contain none of
   * \ / ? * [ ] :, and not start/end with an apostrophe — and must be
   * unique within the workbook. Sheet names reaching here can come from
   * free-form AI/CSV import (see api/ai/route.ts), so none of that is
   * guaranteed. Without this, workbook.addWorksheet() throws and the whole
   * export fails with no useful message to the user.
   */
  private static sanitizeSheetName(rawName: string, usedNames: Set<string>): string {
    let name = (rawName || 'Sheet').replace(/[\\/?*[\]:]/g, '').trim();
    name = name.replace(/^'+|'+$/g, '').trim();
    if (!name) name = 'Sheet';
    if (name.length > 31) name = name.slice(0, 31).trim();

    let candidate = name;
    let suffix = 2;
    while (usedNames.has(candidate.toLowerCase())) {
      const suffixText = ` (${suffix})`;
      candidate = name.slice(0, Math.max(1, 31 - suffixText.length)) + suffixText;
      suffix++;
    }
    usedNames.add(candidate.toLowerCase());
    return candidate;
  }

  private buildSheet(workbook: ExcelJS.Workbook, sheetData: Sheet, palette: typeof PALETTES.premium, safeSheetName: string) {
    const worksheet = workbook.addWorksheet(safeSheetName, {
      views: [{ showGridLines: true }],
    });

    let currentRow = 1;
    const colCount = Math.max(sheetData.columns.length, 6);

    // 1. BANNER HEADER
    worksheet.mergeCells(currentRow, 1, currentRow + 1, colCount);
    const bannerCell = worksheet.getCell(currentRow, 1);
    bannerCell.value = `📊  ${this.product.name.toUpperCase()} — ${sheetData.name.toUpperCase()}`;
    bannerCell.font = {
      name: 'Segoe UI',
      size: 14,
      bold: true,
      color: { argb: 'FF' + palette.bannerText },
    };
    bannerCell.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 };
    bannerCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF' + palette.bannerBg },
    };

    currentRow += 3; // Gap after banner

    // 2. KPI CARDS SECTION (if any KPIs exist)
    if (sheetData.kpis && sheetData.kpis.length > 0) {
      const kpis = sheetData.kpis.slice(0, 4);
      const kpiSpan = Math.max(1, Math.floor(colCount / kpis.length));

      // Row for KPI Labels
      kpis.forEach((kpi, index) => {
        const startCol = index * kpiSpan + 1;
        const endCol = Math.min(startCol + kpiSpan - 1, colCount);

        if (startCol < endCol) {
          worksheet.mergeCells(currentRow, startCol, currentRow, endCol);
        }
        const labelCell = worksheet.getCell(currentRow, startCol);
        labelCell.value = kpi.label.toUpperCase();
        labelCell.font = { name: 'Segoe UI', size: 9, bold: true, color: { argb: 'FF64748B' } };
        labelCell.alignment = { horizontal: 'center', vertical: 'middle' };
        labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + palette.kpiBg } };

        // Value row
        const valRow = currentRow + 1;
        if (startCol < endCol) {
          worksheet.mergeCells(valRow, startCol, valRow, endCol);
        }
        const valueCell = worksheet.getCell(valRow, startCol);
        
        // Find column index
        const colRef = ExcelProductEngine.resolveColumnRef(sheetData.columns, kpi.column);
        if (!colRef) {
          console.error(`[excel-engine] KPI "${kpi.label}": column "${kpi.column}" not found on sheet "${sheetData.name}" — value left at 0.`);
        }
        const colLetter = colRef?.letter ?? 'A';
        const colIdx = colRef?.index ?? -1;
        const dataStartRow = currentRow + 4; // Start of data table
        const dataEndRow = dataStartRow + Math.max(sheetData.rows.length - 1, 0);

        if (!colRef) {
          valueCell.value = 0;
        } else if (kpi.aggregation === 'sum') {
          valueCell.value = { formula: `SUM(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})` };
        } else if (kpi.aggregation === 'avg') {
          valueCell.value = { formula: `AVERAGE(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})` };
        } else if (kpi.aggregation === 'count') {
          valueCell.value = { formula: `COUNTA(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})` };
        } else if (kpi.aggregation === 'max') {
          valueCell.value = { formula: `MAX(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})` };
        } else if (kpi.aggregation === 'min') {
          valueCell.value = { formula: `MIN(${colLetter}${dataStartRow}:${colLetter}${dataEndRow})` };
        } else {
          valueCell.value = 0;
        }

        valueCell.font = { name: 'Segoe UI', size: 14, bold: true, color: { argb: 'FF' + palette.bannerBg } };
        valueCell.alignment = { horizontal: 'center', vertical: 'middle' };
        valueCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + palette.kpiBg } };

        // Format
        if (kpi.format === 'currency' || (!kpi.format && sheetData.columns[colIdx]?.type === 'currency')) {
          valueCell.numFmt = '$#,##0.00';
        } else if (kpi.format === 'percent' || (!kpi.format && sheetData.columns[colIdx]?.type === 'percent')) {
          valueCell.numFmt = '0.0%';
        } else {
          valueCell.numFmt = '#,##0';
        }

        // Card borders
        for (let r = currentRow; r <= valRow; r++) {
          for (let c = startCol; c <= endCol; c++) {
            worksheet.getCell(r, c).border = {
              top: { style: 'thin', color: { argb: 'FF' + palette.kpiBorder } },
              left: { style: 'thin', color: { argb: 'FF' + palette.kpiBorder } },
              bottom: { style: 'thin', color: { argb: 'FF' + palette.kpiBorder } },
              right: { style: 'thin', color: { argb: 'FF' + palette.kpiBorder } },
            };
          }
        }
      });

      currentRow += 3; // Gap after KPI section
    }

    // 3. TABLE SECTION
    const tableHeaderRow = currentRow;
    sheetData.columns.forEach((col, idx) => {
      const cell = worksheet.getCell(tableHeaderRow, idx + 1);
      cell.value = col.label;
      cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF' + palette.headerText } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + palette.headerBg } };
      cell.alignment = {
        vertical: 'middle',
        horizontal: col.type === 'number' || col.type === 'currency' || col.type === 'percent' ? 'right' : 'left',
      };
      cell.border = {
        top: { style: 'medium', color: { argb: 'FF' + palette.bannerBg } },
        bottom: { style: 'medium', color: { argb: 'FF' + palette.bannerBg } },
      };
    });

    currentRow++;

    // 4. DATA ROWS
    sheetData.rows.forEach((row, rowIdx) => {
      const isZebra = rowIdx % 2 === 1;
      const rowFillColor = isZebra ? palette.zebraBg : 'FFFFFF';

      sheetData.columns.forEach((col, colIdx) => {
        const cell = worksheet.getCell(currentRow, colIdx + 1);
        const rawVal = ExcelProductEngine.toCellString(row[colIdx]);

        if (col.type === 'number') {
          const num = parseFloat(rawVal);
          cell.value = isNaN(num) ? rawVal : num;
          cell.numFmt = '#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (col.type === 'currency') {
          const num = parseFloat(rawVal.replace(/[^0-9.-]+/g, ''));
          cell.value = isNaN(num) ? rawVal : num;
          cell.numFmt = '$#,##0.00';
          cell.alignment = { horizontal: 'right' };
        } else if (col.type === 'percent') {
          let num = parseFloat(rawVal.replace('%', ''));
          if (!isNaN(num) && num > 1) num = num / 100;
          cell.value = isNaN(num) ? rawVal : num;
          cell.numFmt = '0.0%';
          cell.alignment = { horizontal: 'right' };
        } else {
          cell.value = rawVal;
          cell.alignment = { horizontal: 'left' };
        }

        cell.font = { name: 'Segoe UI', size: 10, color: { argb: 'FF1E293B' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + rowFillColor } };
        cell.border = {
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };
      });

      currentRow++;
    });

    const lastDataRow = currentRow - 1;

    // 5. TOTALS ROW (if numerical data exists)
    if (sheetData.rows.length > 0) {
      const hasNumeric = sheetData.columns.some((c) => c.type === 'currency' || c.type === 'number');
      if (hasNumeric) {
        sheetData.columns.forEach((col, colIdx) => {
          const cell = worksheet.getCell(currentRow, colIdx + 1);
          const colLetter = ExcelProductEngine.columnIndexToLetter(colIdx);
          const startR = tableHeaderRow + 1;
          const endR = currentRow - 1;

          if (colIdx === 0) {
            cell.value = 'TOTAL';
            cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
            cell.alignment = { horizontal: 'left' };
          } else if (col.type === 'currency') {
            cell.value = { formula: `SUM(${colLetter}${startR}:${colLetter}${endR})` };
            cell.numFmt = '$#,##0.00';
            cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
            cell.alignment = { horizontal: 'right' };
          } else if (col.type === 'number') {
            cell.value = { formula: `SUM(${colLetter}${startR}:${colLetter}${endR})` };
            cell.numFmt = '#,##0.00';
            cell.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FF0F172A' } };
            cell.alignment = { horizontal: 'right' };
          }

          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + palette.totalBg } };
          cell.border = {
            top: { style: 'thin', color: { argb: 'FF' + palette.bannerBg } },
            bottom: { style: 'double', color: { argb: 'FF' + palette.bannerBg } },
          };
        });
      }
    }

    // 6. CHARTS — native Excel chart objects. ExcelJS can't create these
    // itself, so we only record where the data lives here; the actual
    // chart/drawing XML is added as a post-processing step in
    // exportAsBuffer() via chart-injector.ts, once the base workbook buffer
    // exists.
    if (sheetData.charts && sheetData.charts.length > 0 && sheetData.rows.length > 0) {
      const dataStartRow = tableHeaderRow + 1;
      const sheetRefName = quoteSheetRef(safeSheetName);
      let chartSlot = 0;

      sheetData.charts.forEach((chart) => {
        const catRef = ExcelProductEngine.resolveColumnRef(sheetData.columns, chart.categoryColumn);
        if (!catRef) {
          console.error(
            `[excel-engine] Chart "${chart.title}": category column "${chart.categoryColumn}" not found on sheet "${sheetData.name}" — skipping this chart.`
          );
          return;
        }

        const series = chart.valueColumns
          .map((vc) => {
            const ref = ExcelProductEngine.resolveColumnRef(sheetData.columns, vc);
            if (!ref) {
              console.error(
                `[excel-engine] Chart "${chart.title}": value column "${vc}" not found on sheet "${sheetData.name}" — skipping this series.`
              );
              return null;
            }
            return {
              name: sheetData.columns[ref.index].label,
              nameRef: `${sheetRefName}!$${ref.letter}$${tableHeaderRow}`,
              valuesRef: `${sheetRefName}!$${ref.letter}$${dataStartRow}:$${ref.letter}$${lastDataRow}`,
            };
          })
          .filter((s): s is { name: string; nameRef: string; valuesRef: string } => s !== null);

        if (series.length === 0) {
          console.error(`[excel-engine] Chart "${chart.title}": no valid value columns — skipping.`);
          return;
        }

        const fromCol = colCount + 1; // one-column gap after the table (0-based)
        const fromRow = tableHeaderRow - 1 + chartSlot * 17; // 0-based; stack charts 17 rows apart
        chartSlot++;

        this.pendingCharts.push({
          sheetName: safeSheetName,
          title: chart.title,
          type: chart.type,
          categoryRef: `${sheetRefName}!$${catRef.letter}$${dataStartRow}:$${catRef.letter}$${lastDataRow}`,
          series,
          anchor: { fromCol, fromRow, toCol: fromCol + 8, toRow: fromRow + 15 },
        });
      });
    }

    // Auto calculate column widths
    worksheet.columns.forEach((column) => {
      let maxLen = 12;
      column.eachCell?.({ includeEmpty: true }, (cell) => {
        const len = cell.value ? String(cell.value).length : 0;
        if (len > maxLen) maxLen = Math.min(len + 4, 40);
      });
      column.width = maxLen;
    });
  }

  public async exportAsBuffer(): Promise<ArrayBuffer> {
    this.pendingCharts = [];
    const workbook = await this.generateWorkbook();
    const baseBuffer = await workbook.xlsx.writeBuffer();
    if (this.pendingCharts.length === 0) {
      return baseBuffer;
    }
    const chartBuffer = await injectCharts(baseBuffer, this.pendingCharts);
    // injectCharts returns a Node Buffer; slice out just its bytes as a
    // plain ArrayBuffer so every existing caller of exportAsBuffer (API
    // routes, download handlers) keeps working unchanged.
    return chartBuffer.buffer.slice(chartBuffer.byteOffset, chartBuffer.byteOffset + chartBuffer.byteLength) as ArrayBuffer;
  }

  public async exportAsBase64(): Promise<string> {
    const buffer = await this.exportAsBuffer();
    return Buffer.from(buffer).toString('base64');
  }
}
