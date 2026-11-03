/* Real-app DOCX E2E: open sample docx in the live renderer, add paragraphs,
   format, add a table, save, then verify the OOXML bytes in node. */
import JSZip from 'jszip';

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
  console.log('NO PAGE TARGET');
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

/* ---------- seed a minimal docx in the page (dev hook open with docx bytes) ---------- */
const templateDocx = `
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Original Heading</w:t></w:r></w:p>
<w:p><w:r><w:t>Seed body line one.</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body></w:document>`.trim();

const zip = new JSZip();
zip.file(
  '[Content_Types].xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
);
zip.file(
  '_rels/.rels',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
);
zip.file(
  'word/document.xml',
  templateDocx
);
zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
const sampleBytes = await zip.generateAsync({ type: 'uint8array' });
const b64Sample = Buffer.from(sampleBytes).toString('base64');

/* ---------- drive the app ---------- */
await evaluate(`window.__cellDeskDevApp.openB64(${JSON.stringify('sample.docx')}, ${JSON.stringify(b64Sample)}, 'docx')`);
await wait(2500);
const opened = await evaluate(`window.__cellDeskDevDocx?.open?.(${JSON.stringify('sample.docx')}, ${JSON.stringify(b64Sample)}) ?? null`);
console.log('open result:', JSON.stringify(opened));
okm(String(opened?.text || '').includes('Original Heading'), 'docx opened, heading text reads');
await wait(150);

await evaluate(`window.__cellDeskDevDocx.addPara('New chapter paragraph', { style: 'Heading2' })`);
const beforeFmt = (await evaluate(`window.__cellDeskDevDocx.text()`)) || '';
okm(String(beforeFmt).includes('New chapter paragraph'), 'addPara visible');

await evaluate(`window.__cellDeskDevDocx.sel({ mode: 'para', index: 0 })`);
await evaluate(`window.__cellDeskDevDocx.fmt({ b: true })`);

await evaluate(`window.__cellDeskDevDocx.table(2, 2)`);
await wait(100);
const savedB64 = await evaluate(`window.__cellDeskDevDocx.save()`);
okm(typeof savedB64 === 'string' && savedB64.length > 400, 'save() produced docx bytes');

/* ---------- verify OOXML in node ---------- */
const outZip = await JSZip.loadAsync(Buffer.from(savedB64, 'base64'));
const docXml = await outZip.file('word/document.xml').async('string');
okm(docXml.includes('>New chapter paragraph<'), 'new paragraph written');
okm(docXml.includes('w:val="Heading2"'), 'heading2 style written');
okm(docXml.includes('<w:tbl>'), 'table written');
okm((docXml.match(/<w:tbl>/g) || []).length === 1, 'exactly one table');
okm(docXml.includes('w:w="12240"'), 'sectPr preserved');

console.log(`\nE2E-DOCX: ${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
