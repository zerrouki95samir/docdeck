import Papa from 'papaparse';
import { emptySheet } from './excelImport';
import type { CellDT, WorkbookDT } from './dto';
import { dateToExcelSerial, randomId } from './util';

const NUM_RE = /^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/;
const PERCENT_RE = /^-?(\d+\.?\d*|\.\d+)%$/;
const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
const US_DATE_RE = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
const TIME_RE = /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm]?)?$/;
const BOOL_RE = /^(TRUE|FALSE)$/i;

function timeToSerial(h: number, m: number, s: number): number {
  return (h * 3600 + m * 60 + s) / 86400;
}

function smartValue(raw: string): CellDT | null {
  const trimmed = raw.trim();
  if (trimmed === '') return { v: '', t: 1 };

  if (BOOL_RE.test(trimmed)) {
    return { v: trimmed.toUpperCase() === 'TRUE' ? 1 : 0, t: 3 };
  }
  if (NUM_RE.test(trimmed) && trimmed.length < 16 && !/^0\d/.test(trimmed)) {
    const n = Number(trimmed);
    if (Number.isFinite(n)) return { v: n, t: 2 };
  }
  const pct = PERCENT_RE.exec(trimmed);
  if (pct) {
    const n = Number(pct[1]);
    if (Number.isFinite(n)) {
      return { v: n / 100, t: 2, s: { n: { pattern: '0%' } } };
    }
  }
  const iso = ISO_DATE_RE.exec(trimmed);
  if (iso) {
    const y = Number(iso[1]);
    const mo = Number(iso[2]);
    const da = Number(iso[3]);
    let serial = dateToExcelSerial(new Date(Date.UTC(y, mo - 1, da)));
    if (iso[4] != null) {
      serial += timeToSerial(Number(iso[4]), Number(iso[5] || 0), Number(iso[6] || 0));
      return { v: serial, t: 2, s: { n: { pattern: iso[6] ? 'yyyy-mm-dd hh:mm:ss' : 'yyyy-mm-dd h:mm' } } };
    }
    return { v: serial, t: 2, s: { n: { pattern: 'yyyy-mm-dd' } } };
  }
  const us = US_DATE_RE.exec(trimmed);
  if (us) {
    serialUs();
    return {
      v: dateToExcelSerial(new Date(Date.UTC(Number(us[3]), Number(us[1]) - 1, Number(us[2])))),
      t: 2,
      s: { n: { pattern: 'm/d/yyyy' } },
    };
  }
  const time = TIME_RE.exec(trimmed);
  if (time) {
    const h = Number(time[1]);
    const m = Number(time[2]);
    const s = Number(time[3] || 0);
    const suffix = time[4];
    let hours = h;
    let pattern = 'h:mm';
    if (s > 0) pattern = 'h:mm:ss';
    if (suffix) {
      pattern = s > 0 ? 'h:mm:ss AM/PM' : 'h:mm AM/PM';
      const isPM = suffix.toUpperCase().startsWith('P');
      if (isPM && h % 12 !== 0) hours = h % 12 + 12;
      if (!isPM && h === 12) hours = 0;
    }
    if (hours <= 24 && m < 60) {
      return { v: Math.round(timeToSerial(hours, m, s) * 1e6) / 1e6, t: 2, s: { n: { pattern } } };
    }
  }
  return { v: raw, t: 1 };
}

function serialUs() {
  /* placeholder for future 1904 handling */
}

export function csvTextToSnapshot(
  text: string,
  filename = 'Untitled.csv',
  delimiter?: string
): WorkbookDT {
  const parsed = Papa.parse(text, {
    delimiter: delimiter ?? undefined,
    skipEmptyLines: false,
  });
  const rows: string[][] = parsed.data as string[][];
  const sheetId = randomId('s');
  const sheet = emptySheet(sheetId, filename.replace(/\.(csv|tsv)$/i, '')) as never;
  const cellData: Record<string, Record<string, CellDT>> = {};

  let maxCol = 0;
  rows.forEach((row, r) => {
    if (row.length === 1 && row[0] === '' && r === rows.length - 1 && rows.length > 1) return;
    maxCol = Math.max(maxCol, row.length);
    row.forEach((v, c) => {
      const cell = smartValue(v);
      if (cell && cell.v !== '' && cell.v !== undefined) {
        (cellData[r] = cellData[r] || {})[c] = cell;
      }
    });
  });

  const columnData: Record<string, { w?: number; hd?: number }> = {};
  for (let c = 0; c < maxCol; c++) {
    let maxSize = 0;
    for (let r = 0; r < Math.min(rows.length, 500); r++) {
      const val = rows[r]?.[c] ?? '';
      maxSize = Math.max(maxSize, Math.min(String(val).length, 60));
    }
    const width = Math.min(maxSize * 7.5 + 12, 340);
    if (width > 70) columnData[c] = { w: width };
  }

  const rowCount = Math.min(Math.max(rows.length + 20, 200), 100000);
  const columnCount = Math.min(Math.max(maxCol + 8, 26), 500);
  const snapshot: WorkbookDT = {
    id: randomId('wb'),
    name: filename,
    appVersion: '1.2.0',
    locale: 'enUS',
    styles: {},
    sheetOrder: [sheetId],
    sheets: {
      [sheetId]: Object.assign(sheet, {
        rowCount,
        columnCount,
        cellData,
        columnData,
      }),
    },
    views: [],
    resources: [],
  };
  return snapshot;
}

