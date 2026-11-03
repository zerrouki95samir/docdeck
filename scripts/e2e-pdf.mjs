/* Real-app PDF E2E: open a sample PDF in the live renderer, search, rotate,
   fill a form, stamp text, save, verify bytes with pdf-lib in node. */
import { PDFDocument, StandardFonts } from 'pdf-lib';

const port = process.env.CDP_PORT || '9223';

async function wait(ms) {
  await new Promise((r) => setTimeout(r, ms));
}

let passed = 0;
let failed = 0;
const okm = (cond, label) => {
  if (cond) {
    passed++;
    console.log('  ok -', label);
  } else {
    failed++;
    console.log('  FAIL -', label);
  }
};

const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
const page = targets.find((t) => t.type === 'page');
if (!page) {
  console.log('NO PAGE TARGET — start the app with DOCDECK_DEBUG_PORT=9223');
  process.exit(1);
}

const ws = new WebSocket(page.webSocketDebuggerUrl);
let seq = 0;
const pending = new Map();

function send(method, params) {
  return new Promise((resolve, reject) => {
    const id = ++seq;
    pending.set(id, { resolve, reject });
    ws.send(JSON.stringify({ id, method, params }));
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
await send('Runtime.enable');

function evaluate(expression) {
  return send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }).then((r) => {
    if (r.exceptionDetails) throw new Error(JSON.stringify(r.exceptionDetails).slice(0, 800));
    return r.result?.value;
  });
}

/* ---------- build sample PDF ---------- */
async function buildSample() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < 3; i++) {
    const p = doc.addPage([612, 792]);
    p.drawText(`Page ${i + 1} Arbor inspectable`, { x: 70, y: 700, size: 20, font });
    if (i === 0) {
      const form = doc.getForm();
      const name = form.createTextField('employee');
      name.setText('');
      name.addToPage(p, { x: 70, y: 620, width: 160, height: 18 });
      const cb = form.createCheckBox('signed');
      cb.uncheck();
      cb.addToPage(p, { x: 70, y: 590, width: 12, height: 12 });
    }
  }
  const bytes = Buffer.from(await doc.save());
  return bytes.toString('base64');
}

const b64Doc = await buildSample();

/* ---------- drive the app ---------- */
await evaluate(`window.__cellDeskDevApp.openB64(${JSON.stringify('sample.pdf')}, ${JSON.stringify(b64Doc)}, 'pdf')`);
await wait(2500);
const opened = await evaluate(`window.__cellDeskDevPdf?.view?.() ?? null`);
console.log('view:', JSON.stringify(opened));
console.log('open result:', JSON.stringify(opened));
okm(opened?.pages === 3, 'PDF opened in real app (3 pages)');
await wait(500);

const title = await evaluate('document.title');
okm(String(title).includes('sample.pdf'), 'title switched to sample.pdf');

const hits = await evaluate(`window.__cellDeskDevPdf.search('Arbor')`);
okm(Array.isArray(hits) && hits.length >= 3, `search found ${hits?.length ?? 0} pages with the term`);
await wait(200);

const spun = await evaluate(`window.__cellDeskDevPdf.rotate(0)`);
okm(spun?.pages === 3, 'rotate page 0');
await wait(300);

const fieldSet = await evaluate(`window.__cellDeskDevPdf.setField('employee', 'Ada Lovelace')`);
okm(fieldSet === true, 'form field set');
await wait(500);

const stampRes = await evaluate(
  `window.__cellDeskDevPdf.stamp({ type: 'text', pts: [{ x: 100, y: 300 }], text: 'Reviewed by CellDesk', color: '#1e4db7' })`
);
okm(stampRes?.pending === 1, 'stamp entered pending queue');

const saveB64 = await evaluate(`window.__cellDeskDevPdf.save()`);
okm(typeof saveB64 === 'string' && saveB64.length > 500, 'save() returned bytes');

/* ---------- verify the saved bytes in node ---------- */
const bin = Buffer.from(saveB64, 'base64');
const doc = await PDFDocument.load(bin);
okm(doc.getPageCount() === 3, 'saved PDF has 3 pages');
okm(doc.getPages()[0].getRotation().angle === 90, 'saved rotation persisted');
const forms = doc.getForm();
const employeeField = forms.getFieldMaybe('employee');
const employeeValue =
  employeeField && employeeField.constructor.name === 'PDFTextField' ? employeeField.getText() : '';
okm(employeeValue === 'Ada Lovelace', 'saved text field value persisted');
okm(forms.getFieldMaybe('signed')?.isChecked?.() === true || forms.getFieldMaybe('signed') !== undefined, 'checkbox field exists');

console.log(`\nE2E-PDF: ${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
