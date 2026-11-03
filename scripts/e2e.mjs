/* End-to-end drive of the real Electron renderer via CDP:
   node sample xlsx -> openBuffer (real app parse + Univer load) -> saveBytes (fwb.save + serialize) -> verify in node */
import ExcelJS from 'exceljs';
import { appendFileSync, writeFileSync } from 'node:fs';

const LOG = '/tmp/docdeck-e2e.log';
writeFileSync(LOG, '=== e2e start ===\n');
const log = (...a) => {
  const line = a.join(' ');
  console.log(line);
  appendFileSync(LOG, line + '\n');
};

const port = process.env.CDP_PORT || '9223';

const wait = (ms) => new Promise((r) => setTimeout(r, ms));

const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
const page = targets.find((t) => t.type === 'page');
if (!page) throw new Error('NO PAGE TARGET — is the app running with DOCDECK_DEBUG_PORT?');

const ws = new WebSocket(page.webSocketDebuggerUrl);
let seq = 0;
const pending = new Map();

function send(method, params, sessionId) {
  return new Promise((resolve, reject) => {
    const id = ++seq;
    pending.set(id, { resolve, reject });
    ws.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
  });
}

ws.onmessage = (ev) => {
  const msg = JSON.parse(String(ev.data));
  if (msg.id && pending.has(msg.id)) {
    const { resolve, reject } = pending.get(msg.id);
    pending.delete(msg.id);
    if (msg.error) reject(new Error(JSON.stringify(msg.error)));
    else resolve(msg.result);
  }
};

await new Promise((r) => (ws.onopen = r));

function evaluate(expression, label = 'eval') {
  const timeout = new Promise((_, reject) =>
    setTimeout(() => reject(new Error(`EVAL-TIMEOUT after 30s: ${label}`)), 30000)
  );
  return Promise.race([
    send('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
    }).then((r) => {
      if (r.exceptionDetails) throw new Error(JSON.stringify(r.exceptionDetails, null, 2).slice(0, 1200));
      return r.result?.value;
    }),
    timeout,
  ]);
}

// ---- build sample workbook in node ----
const wb = new ExcelJS.Workbook();
const wsx = wb.addWorksheet('Budget', {
  properties: { tabColor: { argb: 'FFFF8800' } },
  views: [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2' }],
});
wsx.mergeCells('A1:E1');
const t1 = wsx.getCell('A1');
t1.value = 'Q1 Budget';
t1.font = { bold: true, size: 16, color: { argb: 'FFCC2024' } };
wsx.getCell('B2').value = 'Qty';
wsx.getCell('B2').font = { bold: true };
wsx.getCell('C3').value = 19.5;
wsx.getCell('C3').numFmt = '#,##0.00';
wsx.getCell('C7').value = { formula: 'SUM(C3:C5)', result: 159.74 };
const bytes = Buffer.from(await wb.xlsx.writeBuffer());
const b64 = bytes.toString('base64');

// ---- drive the renderer ----
await send('Page.enable');
await send('Page.reload');
await wait(2500); // home screen boots first; no lens hooks until a doc opens

log('step: checking app hook...');
const probe = await evaluate(`typeof (window.__cellDeskDevApp?.openB64)`, 'probe');
log('app hook present:', probe);

// home-first: open the workbook through the app-level hook; this mounts the sheet lens
await evaluate(`window.__cellDeskDevApp.openB64(${JSON.stringify('sample.xlsx')}, ${JSON.stringify(b64)}, 'sheet')`);
await wait(2500);

log('step: checking lens dev hook...');
const probe2 = await evaluate(`typeof (window.__cellDeskDev?.openBuffer)`, 'probe2');
log('lens hook present:', probe2);

const opened = await evaluate(
  `(async () => {
    const bin = atob(${JSON.stringify(b64)});
    const arr = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
    const r = await window.__cellDeskDev.openBuffer('sample.xlsx', arr.buffer);
    return { ...r, title: document.title };
  })()`,
  'openBuffer'
);
log('opened in real app:', JSON.stringify(opened));

if (opened.sheets !== 1 || opened.a1 !== 'Q1 Budget') {
  console.error('FAIL: openBuffer did not produce expected snapshot');
  process.exit(1);
}

// need base64 of the Uint8Array; evaluate returns arrays as plain arrays via returnByValue
const u8 = await evaluate(
  `(async () => {
    const u8 = await window.__cellDeskDev.saveBytes('saved-from-app.xlsx');
    let s = '';
    for (let i = 0; i < u8.length; i++) s += String.fromCharCode(u8[i]);
    return { len: u8.length, b64: btoa(s) };
  })()`,
  'saveBytes'
);
log('saved bytes:', u8.len);

const out = Buffer.from(u8.b64, 'base64');
const wb2 = new ExcelJS.Workbook();
await wb2.xlsx.load(out);
const checkWs = wb2.getWorksheet('Budget');
const cellA1 = checkWs.getCell('A1');
log('verify: A1 =', JSON.stringify(cellA1.value), 'bold =', cellA1.font?.bold, 'sheets =', wb2.worksheets.map((w) => w.name).join(','));
const c7Value = wb2.getWorksheet('Budget')?.getCell('C7').value || {};
const cellA1Formula = c7Value.formula || '';
log('verify: C7 formula =', c7Value.formula);
log('verify: C3 numFmt =', wb2.getWorksheet('Budget')?.getCell('C3').numFmt);

const ok =
  cellA1.value === 'Q1 Budget' &&
  cellA1.font?.bold === true &&
  String(cellA1Formula || '').includes('SUM') &&
  wb2.getWorksheet('Budget')?.getCell('C3').numFmt === '#,##0.00';

log(ok ? '\nE2E-PASS: real app opened and saved an xlsx correctly' : '\nE2E-FAIL');
if (!ok) process.exit(1);
process.exit(0);
