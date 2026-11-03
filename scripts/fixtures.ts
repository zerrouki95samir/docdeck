import type { CellDT, SheetDT } from '../src/io/dto';

/** Access cell at 0-based row/col in a loose snapshot sheet. */
export function cellOf(sheet: SheetDT | undefined, row: number, col: number): CellDT | undefined {
  return sheet?.cellData != null ? sheet.cellData[String(row)]?.[String(col)] : undefined;
}
