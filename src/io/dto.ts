/** Loose DTO mirroring Univer's data model, so the io layer never fights deep generic chains. */

export interface CellStyleDT {
  ff?: string;
  fs?: number;
  it?: number;
  bl?: number;
  ul?: { s?: number; t?: number; c?: number; cl?: { rgb?: string } };
  st?: { s?: number; t?: number; c?: number; cl?: { rgb?: string } };
  ol?: { s?: number; t?: number; c?: number; cl?: { rgb?: string } };
  bg?: { rgb?: string };
  bd?: Record<string, { s?: number; cl?: { rgb?: string } } | undefined>;
  cl?: { rgb?: string };
  va?: number;
  tr?: { a?: number; v?: number };
  ht?: number;
  vt?: number;
  tb?: number;
  pd?: Record<string, number>;
  n?: { pattern?: string };
  [k: string]: unknown;
}

export interface CellDT {
  v?: string | number | boolean | null;
  t?: number | null;
  f?: string;
  si?: string;
  s?: string | CellStyleDT | null;
  n?: { pattern?: string };
  p?: unknown;
  [k: string]: unknown;
}

export interface FreezeDT {
  xSplit?: number;
  ySplit?: number;
  startRow?: number;
  startColumn?: number;
}

export interface MergeDT {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
  mergeType?: number;
}

export interface SheetDT {
  id?: string;
  name?: string;
  tabColor?: string;
  hidden?: number;
  rowCount?: number;
  columnCount?: number;
  defaultColumnWidth?: number;
  defaultRowHeight?: number;
  cellData?: Record<string, Record<string, CellDT>>;
  rowData?: Record<string, { h?: number; hd?: number }>;
  columnData?: Record<string, { w?: number; hd?: number }>;
  mergeData?: MergeDT[];
  freeze?: FreezeDT;
  rowHeader?: { width: number; hidden?: number };
  columnHeader?: { height: number; hidden?: number };
  rightToLeft?: number;
  showGridlines?: number;
  [k: string]: unknown;
}

export interface WorkbookDT {
  id?: string;
  name?: string;
  appVersion?: string;
  locale?: string;
  styles?: Record<string, CellStyleDT | undefined>;
  sheetOrder?: string[];
  sheets?: Record<string, SheetDT>;
  views?: unknown[];
  resources?: unknown[];
  [k: string]: unknown;
}
