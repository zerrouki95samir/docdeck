import * as XLSX from 'xlsx';
import { xlsxBufferToSnapshot } from './excelImport';
import type { WorkbookDT } from './dto';

/** Legacy .xls (BIFF) -> xlsx via SheetJS, then import with full fidelity where possible. */
export async function xlsBufferToSnapshot(data: Uint8Array | ArrayBuffer, fallbackName: string): Promise<WorkbookDT> {
  const wb = XLSX.read(data, { type: 'array', cellStyles: true });
  const xlsxBytes: ArrayBuffer = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return xlsxBufferToSnapshot(new Uint8Array(xlsxBytes), fallbackName);
}
