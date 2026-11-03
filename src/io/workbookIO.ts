import type { WorkbookDT } from './dto';
import { sheetToCsvText, snapshotToXlsxBuffer } from './excelExport';
import { xlsxBufferToSnapshot } from './excelImport';
import { csvTextToSnapshot } from './csv';
import { xlsBufferToSnapshot } from './xls';

export type FileExt = 'xlsx' | 'xlsm' | 'xls' | 'csv' | 'tsv' | 'uni';

export function extOfPath(path: string): FileExt {
  const m = /\.([a-z0-9]+)$/i.exec(path);
  const e = (m?.[1] || 'xlsx').toLowerCase();
  if (e === 'xlsm' || e === 'xls' || e === 'csv' || e === 'tsv' || e === 'uni') return e;
  return 'xlsx';
}

export function basename(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}

export function stripExt(name: string): string {
  return name.replace(/\.[a-z0-9]+$/i, '');
}

/** File bytes (any supported format) -> Univer workbook snapshot */
export async function parseSpreadsheet(path: string, bytes: Uint8Array | ArrayBuffer): Promise<WorkbookDT> {
  const ext = extOfPath(path);
  const name = basename(path);
  switch (ext) {
    case 'xls':
      return xlsBufferToSnapshot(new Uint8Array(bytes), stripExt(name));
    case 'csv':
      return csvTextToSnapshot(new TextDecoder().decode(bytes), name, ',');
    case 'tsv':
      return csvTextToSnapshot(new TextDecoder().decode(bytes), name, '\t');
    case 'uni':
      return JSON.parse(new TextDecoder().decode(bytes)) as WorkbookDT;
    case 'xlsx':
    case 'xlsm':
    default: {
      try {
        return await xlsxBufferToSnapshot(new Uint8Array(bytes), stripExt(name));
      } catch (err) {
        console.error('xlsx parse failed, retrying via SheetJS', err);
        const XLSX = await import('xlsx');
        const wb = XLSX.read(new Uint8Array(bytes), { type: 'array' });
        const repaired: ArrayBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
        return await xlsxBufferToSnapshot(new Uint8Array(repaired), stripExt(name));
      }
    }
  }
}

/** Univer snapshot -> bytes for the given target path */
type BytesInput = Uint8Array | ArrayBuffer;
export function asU8(b: BytesInput): Uint8Array {
  return b instanceof Uint8Array ? b : new Uint8Array(b);
}

export async function serializeSpreadsheet(
  snapshot: WorkbookDT,
  activeSheetId: string,
  path: string
): Promise<Uint8Array> {
  const ext = extOfPath(path);
  switch (ext) {
    case 'csv':
    case 'tsv': {
      const text = sheetToCsvText(snapshot.sheets?.[activeSheetId], ext === 'tsv' ? '\t' : ',');
      return new TextEncoder().encode(text);
    }
    case 'uni':
      return new TextEncoder().encode(JSON.stringify(snapshot));
    case 'xls': {
      const xlsxBuffer = await snapshotToXlsxBuffer(snapshot);
      const XLSX = await import('xlsx');
      const wb = XLSX.read(xlsxBuffer, { type: 'array' });
      const out: ArrayBuffer = XLSX.write(wb, { type: 'array', bookType: 'xls' });
      return new Uint8Array(out);
    }
    case 'xlsx':
    case 'xlsm':
    default: {
      return await snapshotToXlsxBuffer(snapshot);
    }
  }
}
