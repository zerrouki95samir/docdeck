import ExcelJS from 'exceljs';
import type { CellDT, SheetDT, WorkbookDT } from './dto';
import { colPxToChars, cssToArgb, excelSerialToDateUTC, numDaysToIso, rowPxToPt } from './util';
import { univerStyleToExcel } from './styleBridge';

/* eslint-disable @typescript-eslint/no-explicit-any */
type Any = any;

function resolveNamedStyle(snapshot: WorkbookDT, s: CellDT['s']): Any {
  if (s == null) return undefined;
  if (typeof s === 'string') return snapshot.styles?.[s];
  return s;
}

function cellValueForExcel(cell: CellDT): Any {
  if (cell.f) {
    const formula = cell.f.trim();
    const isMatrix = formula.startsWith('={') && formula.endsWith('}');
    const body = isMatrix ? formula.slice(2, -1) : formula.replace(/^=/, '');
    return { formula: body, result: cell.v ?? null };
  }
  if (cell.t === 3) return cell.v === 1 || cell.v === true;
  if (typeof cell.v === 'string' && (cell.v as string) === '#N/A') return cell.v;
  if (cell.v == null) return '';
  return cell.v;
}

function a1(col0: number, row0: number): string {
  let col = col0 + 1;
  let out = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    col = Math.floor((col - mod) / 26);
  }
  return `${out}${row0 + 1}`;
}

function isDatePattern(p?: string): boolean {
  if (!p) return false;
  const q = p.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '').toLowerCase();
  return q.includes('yy') && (q.includes('m') || q.includes('d'));
}

export async function snapshotToXlsxBuffer(snapshot: WorkbookDT): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'DocDeck';

  const sheetOrder = snapshot.sheetOrder?.length
    ? snapshot.sheetOrder
    : Object.keys(snapshot.sheets || {});

  for (const sheetId of sheetOrder) {
    const sheet: SheetDT = snapshot.sheets?.[sheetId] as SheetDT;
    if (!sheet) continue;
    const state = sheet.hidden === 1 ? 'hidden' : sheet.hidden === 2 ? 'veryHidden' : 'visible';
    const tabColor = sheet.tabColor ? { argb: cssToArgb(sheet.tabColor) } : undefined;
    const ws = wb.addWorksheet(sheet.name || 'Sheet', {
      state,
      properties: { tabColor },
    } as Any);

    for (const m of sheet.mergeData || []) {
      try {
        ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1);
      } catch {
        /* ignore invalid merges */
      }
    }

    const cellData = sheet.cellData || {};
    for (const rowKey of Object.keys(cellData)) {
      const r = Number(rowKey);
      if (!Number.isInteger(r) || r < 0) continue;
      for (const colKey of Object.keys(cellData[r])) {
        const c = Number(colKey);
        if (!Number.isInteger(c) || c < 0) continue;
        const data = cellData[r][c];
        if (!data) continue;
        const excelCell: Any = ws.getCell(r + 1, c + 1);
        const value = cellValueForExcel(data);
        if (value !== '' && value !== null) excelCell.value = value;

        const style = resolveNamedStyle(snapshot, data.s) || {};
        const pattern = data.n?.pattern || (style as Any).n?.pattern;
        univerStyleToExcel(excelCell, style, pattern);
      }
    }

    for (const rk of Object.keys(sheet.rowData || {})) {
      const r = Number(rk);
      const rd = sheet.rowData![r];
      if (!rd) continue;
      const row = ws.getRow(r + 1);
      if (rd.h != null) row.height = rowPxToPt(rd.h);
      if (rd.hd) {
        row.hidden = true;
        // exceljs only persists hidden rows that also carry a height
        if (row.height == null) row.height = rowPxToPt(rd.h ?? sheet.defaultRowHeight ?? 20);
      }
    }

    for (const ck of Object.keys(sheet.columnData || {})) {
      const c = Number(ck);
      const cd = sheet.columnData![c];
      if (!cd) continue;
      const col = ws.getColumn(c + 1);
      if (cd.hd) {
        col.hidden = true;
        if (cd.w == null) col.width = 10;
      }
      if (cd.w != null) col.width = Math.round(colPxToChars(cd.w) * 100) / 100;
      if (cd.hd) col.hidden = true;
    }

    const freeze = sheet.freeze;
    if ((freeze?.xSplit ?? 0) > 0 || (freeze?.ySplit ?? 0) > 0) {
      const startRow = Math.max(freeze!.startRow ?? freeze!.ySplit ?? 0, freeze!.ySplit ?? 0);
      const startCol = Math.max(freeze!.startColumn ?? freeze!.xSplit ?? 0, freeze!.xSplit ?? 0);
      const view: Any = {
        state: 'frozen',
        xSplit: freeze!.xSplit || 0,
        ySplit: freeze!.ySplit || 0,
        topLeftCell: a1(startCol, startRow),
      };
      ws.views = [view];
    }

    if (sheet.showGridlines === 0) {
      const views: Any[] = ws.views || [];
      if (views.length) views[0].showGridLines = false;
      else ws.views = [{ showGridLines: false }];
    }

    if (sheet.defaultColumnWidth && sheet.defaultColumnWidth !== 64) {
      ws.properties.defaultColWidth = Math.round(colPxToChars(sheet.defaultColumnWidth) * 100) / 100;
    }
    if (sheet.defaultRowHeight && sheet.defaultRowHeight !== 20) {
      ws.properties.defaultRowHeight = rowPxToPt(sheet.defaultRowHeight);
    }
  }

  const out: Any = await wb.xlsx.writeBuffer();
  // exceljs may hand back an ArrayBuffer-like; normalize to a real Uint8Array copy (no Node Buffer in renderer)
  if (out instanceof Uint8Array) return out;
  const ab = out instanceof ArrayBuffer ? out : (out?.buffer as ArrayBuffer);
  return new Uint8Array(ab.slice ? ab.slice(0) : ab);
}

/** CSV/TSV export of one sheet */
export function sheetToCsvText(sheetData: SheetDT | undefined, delimiter = ','): string {
  const cellData = (sheetData?.cellData || {}) as Record<string, Record<string, CellDT>>;
  let maxRow = -1;
  let maxCol = -1;
  for (const rk of Object.keys(cellData)) {
    const r = Number(rk);
    maxRow = Math.max(maxRow, r);
    for (const ck of Object.keys(cellData[rk])) maxCol = Math.max(maxCol, Number(ck));
  }
  const lines: string[] = [];
  for (let r = 0; r <= maxRow; r++) {
    const parts: string[] = [];
    for (let c = 0; c <= maxCol; c++) {
      parts.push(csvEscape(cellForCsv(cellData[r]?.[c]), delimiter));
    }
    while (parts.length && parts[parts.length - 1] === '') parts.pop();
    lines.push(parts.join(delimiter));
  }
  return lines.join('\r\n') + '\r\n';
}

function cellForCsv(cell?: CellDT): string {
  if (!cell || cell.v == null) return '';
  const v = cell.v;
  if (typeof v === 'number') {
    const styleFromCell = (cell.s && typeof cell.s === 'object' ? cell.s : null) as Any;
    const pattern = cell.n?.pattern || styleFromCell?.n?.pattern;
    if (pattern && isDatePattern(pattern)) return numDaysToIso(excelSerialToDateUTC(v));
    if (cell.t === 3) return v === 1 ? 'TRUE' : 'FALSE';
    return String(v);
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}

function csvEscape(s: string, delimiter = ','): string {
  if (/[",\n\r]/.test(s) || s.includes(delimiter)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}
