import ExcelJS from 'exceljs';
import { findStyleObject } from './styleBridge';
import type { CellDT, SheetDT, WorkbookDT } from './dto';
import { MAX_COLS, MAX_ROWS, argbToCss, colCharsToPx, dateToExcelSerial, parseA1Range, randomId, rowPtToPx } from './util';

/* eslint-disable @typescript-eslint/no-explicit-any */
type Any = any;

const DEFAULT_COL_W_PX = 64;
const DEFAULT_ROW_H_PX = 20;

function decodeCellValue(value: Any): { v: CellDT['v']; t: number | null } {
  if (value == null) return { v: null, t: null };
  if (typeof value === 'number') return { v: value, t: 2 };
  if (typeof value === 'boolean') return { v: value ? 1 : 0, t: 3 };
  if (value instanceof Date) return { v: dateToExcelSerial(value), t: 2 };
  if (typeof value === 'object') {
    if (value.error) return { v: String(value.error), t: 1 };
    if (value.richText && Array.isArray(value.richText)) {
      return { v: value.richText.map((r: Any) => r?.text ?? '').join(''), t: 1 };
    }
    if (value.text != null && value.hyperlink != null) return { v: String(value.text), t: 1 };
    if (value.result !== undefined) return decodeCellValue(value.result);
    return { v: null, t: null };
  }
  return { v: String(value), t: 1 };
}

function readCell(cell: Any): CellDT | null {
  const out: CellDT = {};
  let hasContent = false;

  let formula: string = cell.formula || '';
  if (!formula && cell.formulas?.length) formula = `={${cell.formulas[0]}}`;
  if (formula) {
    out.f = formula.startsWith('=') ? formula : `=${formula}`;
    hasContent = true;
    const res = decodeCellValue(cell.value?.result);
    if (res.v !== null && res.v !== undefined) {
      out.v = res.v;
      out.t = res.t;
    }
  } else {
    const { v, t } = decodeCellValue(cell.value);
    if (typeof v === 'string' && v === '') {
      out.v = '';
      out.t = 1;
      hasContent = true;
    } else if (v !== null && v !== undefined) {
      out.v = v;
      out.t = t;
      hasContent = true;
    }
  }

  const style = findStyleObject(cell);
  let pattern: string | undefined;
  if (cell.value instanceof Date) {
    pattern = cell.numFmt && cell.numFmt !== 'General' ? cell.numFmt : (out.v as number) < 1 ? 'h:mm:ss AM/PM' : 'yyyy-mm-dd';
  } else if (cell.numFmt && cell.numFmt !== 'General' && typeof cell.numFmt === 'string') {
    pattern = cell.numFmt;
  }
  if (pattern) {
    out.s = { ...(style || {}), n: { pattern } } as CellDT['s'];
    hasContent = true;
  } else if (style) {
    out.s = style;
    hasContent = true;
  }

  return hasContent ? out : null;
}

function parseTopLeft(cellRef?: string): { row: number; col: number } | null {
  if (!cellRef) return null;
  const m = /^([A-Z]+)(\d+)$/.exec(cellRef);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { row: parseInt(m[2], 10) - 1, col: col - 1 };
}

export function makeSheet(worksheet: Any): SheetDT {
  const sheetId = randomId('s');
  const merged = ((worksheet.model?.merges ?? []) as string[])
    .map((m) => parseA1Range(m as never))
    .filter((m): m is NonNullable<ReturnType<typeof parseA1Range>> => !!m);

  let usedRow = Number.isFinite(worksheet.rowCount) ? worksheet.rowCount : 0;
  let usedCol = Number.isFinite(worksheet.columnCount) ? worksheet.columnCount : 0;
  for (let r = 1; r <= Math.min(usedRow || 0, MAX_ROWS); r++) {
    const row = worksheet.getRow(r);
    const cl = row?.cellCount ?? 0;
    if (cl > usedCol) usedCol = cl;
  }
  for (const m of merged) {
    usedRow = Math.max(usedRow, m.bottom + 1);
    usedCol = Math.max(usedCol, m.right + 1);
  }

  const rowCount = Math.min(Math.max(usedRow + 12, 200), MAX_ROWS);
  const columnCount = Math.min(Math.max(usedCol + 6, 26), MAX_COLS);

  const cellData: Record<string, Record<string, CellDT>> = {};
  const rowData: Record<string, { h?: number; hd?: number }> = {};
  const columnData: Record<string, { w?: number; hd?: number }> = {};

  for (let r = 1; r <= Math.min(usedRow, MAX_ROWS); r++) {
    const row = worksheet.getRow(r);
    if (!row) continue;
    const rowOut: { h?: number; hd?: number } = {};
    let rowHas = false;
    row.eachCell({ includeEmpty: true }, (cell: Any, colNumber: number) => {
      if (colNumber > Math.min(columnCount, MAX_COLS)) return;
      const parsed = readCell(cell);
      if (parsed) {
        (cellData[r - 1] = cellData[r - 1] || {})[colNumber - 1] = parsed;
      }
    });
    if (row.height != null) {
      rowOut.h = rowPtToPx(row.height);
      rowHas = true;
    }
    if (row.hidden) {
      rowOut.hd = 1;
      rowHas = true;
    }
    if (rowHas) rowData[r - 1] = rowOut;
  }

  for (let c = 1; c <= Math.min(columnCount, MAX_COLS); c++) {
    const col = worksheet.getColumn(c);
    if (!col) continue;
    const colOut: { w?: number; hd?: number } = {};
    let colHas = false;
    if (col.width != null) {
      colOut.w = colCharsToPx(col.width);
      colHas = true;
    }
    if (col.hidden) {
      colOut.hd = 1;
      colHas = true;
    }
    if (colHas) columnData[c - 1] = colOut;
  }

  const sheet: SheetDT = {
    id: sheetId,
    name: worksheet.name || `Sheet${worksheet.id}`,
    rowCount,
    columnCount,
    defaultColumnWidth: DEFAULT_COL_W_PX,
    defaultRowHeight: DEFAULT_ROW_H_PX,
    cellData,
    rowData,
    columnData,
    mergeData: merged.map((m) => ({
      startRow: m.top,
      startColumn: m.left,
      endRow: m.bottom,
      endColumn: m.right,
    })),
    freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    rightToLeft: 0,
    showGridlines: 1,
  };

  const tabColor = worksheet.properties?.tabColor;
  if (tabColor?.argb) {
    const rgb = argbToCss(tabColor.argb);
    if (rgb) sheet.tabColor = rgb;
  }

  if (worksheet.state === 'hidden') sheet.hidden = 1;
  else if (worksheet.state === 'veryHidden') sheet.hidden = 2;

  const view: Any = (worksheet.views || []).find((v: Any) => v?.state === 'frozen');
  if (view) {
    const xSplit = view.xSplit || 0;
    const ySplit = view.ySplit || 0;
    const tl = parseTopLeft(view.topLeftCell);
    sheet.freeze = {
      xSplit,
      ySplit,
      startRow: tl ? tl.row : ySplit,
      startColumn: tl ? tl.col : xSplit,
    };
  }

  return sheet;
}

export async function xlsxBufferToSnapshot(data: Uint8Array | ArrayBuffer, fallbackName = 'Workbook'): Promise<WorkbookDT> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(ensureExcelInput(data));

  const sheets: Record<string, SheetDT> = {};
  const sheetOrder: string[] = [];
  for (const worksheet of wb.worksheets) {
    const sheet = makeSheet(worksheet as Any);
    const sid = sheet.id as string;
    sheets[sid] = sheet;
    sheetOrder.push(sid);
  }
  if (!sheetOrder.length) {
    const id = randomId('s');
    sheets[id] = emptySheet(id, 'Sheet1');
    sheetOrder.push(id);
  }

  return {
    id: randomId('wb'),
    name: fallbackName,
    appVersion: '1.2.0',
    locale: 'enUS',
    styles: {},
    sheetOrder,
    sheets,
    views: [],
    resources: [],
  };
}

function ensureExcelInput(data: Uint8Array | ArrayBuffer): Any {
  if (data instanceof Uint8Array) {
    const ab = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
    return ab as never;
  }
  return data as never;
}

export function emptySheet(id: string, name: string): SheetDT {
  return {
    id,
    name,
    rowCount: 200,
    columnCount: 26,
    defaultColumnWidth: DEFAULT_COL_W_PX,
    defaultRowHeight: DEFAULT_ROW_H_PX,
    cellData: {},
    rowData: {},
    columnData: {},
    mergeData: [],
    freeze: { xSplit: 0, ySplit: 0, startRow: -1, startColumn: -1 },
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    rightToLeft: 0,
    showGridlines: 1,
  };
}
