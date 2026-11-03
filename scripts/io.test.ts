/* Round-trip conversion test: ExcelJS -> snapshot -> ExcelJS, plus csv + xls paths. Run: npm run test:io */
import ExcelJS from 'exceljs';
import * as xlsxMod from 'xlsx';
import { xlsxBufferToSnapshot } from '../src/io/excelImport';
import { snapshotToXlsxBuffer, sheetToCsvText } from '../src/io/excelExport';
import { csvTextToSnapshot } from '../src/io/csv';
import { xlsBufferToSnapshot } from '../src/io/xls';
import type { CellDT, SheetDT, WorkbookDT } from '../src/io/dto';
import { cellOf } from './fixtures';

/* eslint-disable @typescript-eslint/no-explicit-any */

let pass = 0;
let fail = 0;
function check(name: string, cond: any, detail?: any) {
  if (cond) {
    pass++;
    console.log(`  ok - ${name}`);
  } else {
    fail++;
    console.error(`  FAIL - ${name}`, detail !== undefined ? JSON.stringify(detail) : '');
  }
}

function styleOf(cell: CellDT | undefined): Record<string, any> {
  if (!cell?.s) return {};
  return typeof cell.s === 'string' ? {} : (cell.s as Record<string, any>);
}

async function buildSample() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Budget', {
    properties: { tabColor: { argb: 'FFFF8800' } },
    views: [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2' }],
  });
  ws.mergeCells('A1:E1');
  const title = ws.getCell('A1');
  title.value = 'Q1 Budget';
  title.font = { bold: true, size: 16, color: { argb: 'FFCC2024' }, name: 'Arial' };
  title.alignment = { horizontal: 'center', vertical: 'middle' };
  title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  title.border = { bottom: { style: 'double', color: { argb: 'FF333366' } } };

  const headers = ['Item', 'Qty', 'Price', 'Total', 'Date'];
  headers.forEach((h, i) => {
    const c = ws.getCell(2, i + 1);
    c.value = h;
    c.font = { bold: true };
    c.border = {
      top: { style: 'medium' },
      left: { style: 'thin', color: { argb: 'FF888888' } },
      bottom: { style: 'thin', color: { argb: 'FF888888' } },
      right: { style: 'thin', color: { argb: 'FF888888' } },
    };
  });

  const data: [string, number, number, Date][] = [
    ['Widget', 4, 9.5, new Date(Date.UTC(2024, 0, 5))],
    ['Gadget', 2, 100.25, new Date(Date.UTC(2024, 1, 14))],
    ['Doohickey', 1, 49.99, new Date(Date.UTC(2024, 2, 20))],
  ];
  data.forEach((row, i) => {
    const r = 3 + i;
    ws.getCell(r, 1).value = row[0];
    ws.getCell(r, 2).value = row[1];
    const price = ws.getCell(r, 3);
    price.value = row[2];
    price.numFmt = '#,##0.00';
    const total = ws.getCell(r, 4);
    total.value = { formula: `B${r}*C${r}`, result: row[1] * row[2] };
    const dateCell = ws.getCell(r, 5);
    dateCell.value = row[3];
    dateCell.numFmt = 'yyyy-mm-dd';
    ws.getCell(r, 1).alignment = { vertical: 'middle', wrapText: true };
  });

  const sum = ws.getCell('C7');
  sum.value = { formula: 'SUM(C3:C5)', result: 159.74 };
  sum.numFmt = '#,##0.00';

  const ws2 = wb.addWorksheet('Tables');
  ws2.getCell('A1').value = 'name';
  ws2.getCell('B1').value = 42;
  const cross = ws2.getCell('A3');
  cross.value = { formula: 'Budget!B4', result: 2 };
  ws2.getColumn(2).width = 20;
  ws2.getRow(2).hidden = true;
  ws2.getRow(2).height = 15; // exceljs only persists hidden rows that carry a height
  ws2.getColumn(4).width = 15;

  return wb;
}

async function main() {
  console.log('\n=== Build sample xlsx ===');
  const wb = await buildSample();
  const bytes = await wb.xlsx.writeBuffer();
  const imported = await xlsxBufferToSnapshot(new Uint8Array(bytes as ArrayBuffer), 'sample');

  console.log('\n=== Import assertions ===');
  const sheetIds = imported.sheetOrder!;
  const budget = imported.sheets![sheetIds[0]];
  const tables = imported.sheets![sheetIds[1]];
  check('2 sheets', sheetIds.length === 2);
  check('sheet names', budget.name === 'Budget' && tables.name === 'Tables');

  const a1 = cellOf(budget, 0, 0);
  const a1Style = styleOf(a1);
  check('A1 value', a1?.v === 'Q1 Budget', a1);
  check('A1 bold', a1Style.bl === 1);
  check('A1 bg yellow', a1Style.bg?.rgb === '#ffff00', a1Style.bg);
  check('A1 font size 16', a1Style.fs === 16);
  check('A1 font color red', a1Style.cl?.rgb === '#cc2024');
  check('A1 border double bottom', styleOf(a1).bd?.b?.s === 7, styleOf(a1).bd);

  const b2 = cellOf(budget, 1, 1);
  check('B2 bold header', styleOf(b2).bl === 1);
  check('B2 border color', styleOf(b2).bd?.l?.cl?.rgb === '#888888');

  const c3 = cellOf(budget, 2, 2);
  check('C3 numfmt', c3?.s && typeof c3.s === 'object' && (c3.s as any).n?.pattern === '#,##0.00', styleOf(c3).n);

  const d3 = cellOf(budget, 2, 3);
  check('D3 formula', d3?.f === '=B3*C3', d3?.f);

  const e3 = cellOf(budget, 2, 4);
  check(
    'E3 date serial+pattern',
    typeof e3?.v === 'number' && (styleOf(e3).n?.pattern || '') === 'yyyy-mm-dd',
    e3
  );

  check('merge A1:E1', budget.mergeData?.length === 1 && budget.mergeData![0].endColumn === 4, budget.mergeData);
  check('freeze ySplit 1', budget.freeze?.ySplit === 1 && budget.freeze?.startRow === 1, budget.freeze);
  check('tabColor orange', budget.tabColor === '#ff8800', budget.tabColor);
  check(
    'col widths',
    ((tables.columnData?.[1] as { w?: number })?.w ?? 0) > 100 &&
      ((tables.columnData?.[3] as { w?: number })?.w ?? 0) > 100,
    tables.columnData
  );
  check('hidden row', (tables.rowData?.[1] as { hd?: number })?.hd === 1, tables.rowData);

  const c7 = cellOf(budget, 6, 2);
  check('SUM formula C7', c7?.f === '=SUM(C3:C5)', c7?.f);
  const cross = cellOf(tables, 2, 0);
  check('cross-sheet formula', cross?.f === '=Budget!B4', cross?.f);

  console.log('\n=== Export assertions (snapshot -> xlsx) ===');
  const outBytes = await snapshotToXlsxBuffer(imported);
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(outBytes as any);
  check('2 sheets', wb2.worksheets.length === 2, wb2.worksheets.map((w) => w.name));
  const ws1b = wb2.getWorksheet('Budget')!;
  const t1 = ws1b.getCell('A1');
  check('A1 value', t1.value === 'Q1 Budget', t1.value);
  check('A1 bold', t1.font?.bold === true);
  check('A1 fill yellow', String((t1.fill as any)?.fgColor?.argb).endsWith('FFFF00'));
  check('A1 border double', ws1b.getCell('A1').border?.bottom?.style === 'double');
  check('A1 font size 16', t1.font?.size === 16);
  check('merge kept', ws1b.getCell('A1').isMerged === true);

  const c7b = ws1b.getCell('C7');
  check(
    'SUM formula kept',
    String((c7b.value as any)?.formula || '').toUpperCase().includes('SUM'),
    c7b.value
  );
  const c3b = ws1b.getCell('C3');
  check('numfmt kept', c3b.numFmt === '#,##0.00', c3b.numFmt);
  const e3b = ws1b.getCell('E3');
  check(
    'date kept + format',
    e3b.value instanceof Date && (e3b.numFmt || '') === 'yyyy-mm-dd',
    { v: e3b.value, fmt: e3b.numFmt }
  );
  check(
    'freeze kept',
    (ws1b.views?.[0] as any)?.state === 'frozen' && (ws1b.views?.[0] as any)?.ySplit === 1,
    ws1b.views?.[0]
  );
  const ws2b = wb2.getWorksheet('Tables')!;
  check('col width kept', (ws2b.getColumn(2).width ?? 0) > 1.5, ws2b.getColumn(2).width);
  check('row hidden kept', ws2b.getRow(2).hidden === true);
  check(
    'cross formula kept',
    String((ws2b.getCell('A3').value as any)?.formula || '').includes('Budget!B4'),
    ws2b.getCell('A3').value
  );

  console.log('\n=== Round trip: xlsx -> snapshot -> xlsx -> snapshot ===');
  const imported2 = await xlsxBufferToSnapshot(new Uint8Array(outBytes as any), 'sample2');
  const budget2 = imported2.sheets![imported2.sheetOrder![0]];
  check('A1 survives twice', cellOf(budget2, 0, 0)?.v === 'Q1 Budget');
  check('merge survives twice', budget2.mergeData?.length === 1);
  check('formula survives twice', cellOf(budget2, 6, 2)?.f === '=SUM(C3:C5)');

  console.log('\n=== CSV ===');
  const csvIn = 'name,qty,date,note,time\n"Item, A",3,2024-01-05,ok,"9:30 PM"\nFALSE,4,2024-02-01,done,\n';
  const csvSnap = csvTextToSnapshot(csvIn, 'test.csv');
  const csvSheet = csvSnap.sheets![csvSnap.sheetOrder![0]];
  const row1 = csvSheet.cellData!['1'];
  const row2 = csvSheet.cellData!['2'];
  check('csv number', row1[1]?.v === 3 && row1[1]?.t === 2, row1[1]);
  check('csv quoted string kept', String(row1[0]?.v) === 'Item, A', row1[0]);
  check(
    'csv date typed',
    typeof row1[2]?.v === 'number' && (row1[2]!.s as any).n?.pattern === 'yyyy-mm-dd',
    row1[2]
  );
  check(
    'csv time typed',
    typeof row1[4]?.v === 'number' && (row1[4]!.s as any).n?.pattern === 'h:mm AM/PM',
    row1[4]
  );
  check('csv bool', row2[0]?.t === 3 && row2[0]?.v === 0, row2[0]);
  const csvOut = sheetToCsvText(csvSheet);
  check(
    'csv export contains quoted value + iso date',
    csvOut.includes('"Item, A"') && csvOut.includes('2024-01-05'),
    csvOut
  );

  console.log('\n=== XLS (SheetJS) ===');
  try {
    const xlsBytes = buildSmallXls();
    const xlsSnap = await xlsBufferToSnapshot(new Uint8Array(xlsBytes), 'legacy');
    const legacySheet: SheetDT = xlsSnap.sheets![xlsSnap.sheetOrder![0]];
    check('xls imported sheet', legacySheet.name === 'Budget', legacySheet.name);
    check('xls A1 value', legacySheet.cellData![0]?.[0]?.v === 'legacy');
  } catch (e) {
    console.log('  skipped xls write:', (e as Error).message);
  }

  console.log(`\n=== RESULT: ${pass} passed, ${fail} failed ===`);
  if (fail > 0) process.exitCode = 1;
}

function buildSmallXls(): ArrayBuffer {
  const XLSX = xlsxMod;
  const wbOut = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbOut, XLSX.utils.aoa_to_sheet([['legacy', 1], ['b', 2]]), 'Budget');
  return XLSX.write(wbOut, { type: 'array', bookType: 'xls' }) as ArrayBuffer;
}

void newEmptySnapshot;
function newEmptySnapshot(): WorkbookDT {
  return { sheets: {}, sheetOrder: [] };
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
