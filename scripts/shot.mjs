/* capture screenshots of key views for visual QA */
import { writeFileSync } from 'node:fs';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const port = process.env.CDP_PORT || '9223';
const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
const page = targets.find((t) => t.type === 'page');
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
    const p = pending.get(msg.id);
    pending.delete(msg.id);
    msg.error ? reject(new Error(JSON.stringify(msg.error))) : p.resolve(msg.result);
  }
};
await new Promise((r) => (ws.onopen = r));
await send('Runtime.enable');
const evaluate = async (expression) => {
  const r = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
  if (r.exceptionDetails) throw new Error(JSON.stringify(r.exceptionDetails).slice(0, 500));
  return r.result?.value;
};
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const shot = async (name) => {
  const { data } = await send('Page.captureScreenshot', { format: 'png' });
  writeFileSync(`/tmp/cd-shot-${name}.png`, Buffer.from(data, 'base64'));
  console.log('shot:', name);
};

const openB64 = (name, b64, kind) =>
  evaluate(`window.__cellDeskDevApp.openB64(${JSON.stringify(name)}, ${JSON.stringify(b64)}, ${JSON.stringify(kind)})`);

/* --- samples built in node --- */
const pdfBytes = Buffer.from(
  await (async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const bold = await doc.embedFont(StandardFonts.HelveticaBold);
    for (let i = 0; i < 3; i++) {
      const p = doc.addPage([612, 792]);
      p.drawText('Annual Report — Page ' + (i + 1), { x: 70, y: 700, size: 22, font: bold });
      for (let l = 0; l < 8; l++) {
        p.drawText(
          `Section ${l + 1}: revenue grew ${(12.4 + i * 3.1 + l * 0.4).toFixed(1)}% year over year across all divisions of the group holding`,
          { x: 70, y: 660 - l * 26, size: 12, font }
        );
      }
    }
    return doc.save();
  })()
);

const docxZip = new JSZip();
docxZip.file(
  '[Content_Types].xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`
);
docxZip.file(
  '_rels/.rels',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`
);
docxZip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`);
docxZip.file(
  'word/document.xml',
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Quarterly Business Review</w:t></w:r></w:p>
<w:p><w:r><w:t>Growth was steady across the quarter, led by the infrastructure division and subscription renewals.</w:t></w:r></w:p>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Region</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Δ YoY</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>Americas</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>+14.2%</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>EMEA</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>+9.8%</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`
);
const docxBytes = await docxZip.generateAsync({ type: 'nodebuffer' });

const wb = new ExcelJS.Workbook();
const wsx = wb.addWorksheet('Budget');
wsx.mergeCells('A1:D1');
const t1 = wsx.getCell('A1');
t1.value = 'Marketing spend — Q1 review';
t1.font = { bold: true, size: 14, color: { argb: 'FF2451B3' } };
wsx.getCell('A3').value = 'Line item';
wsx.getCell('B3').value = 'Jan';
wsx.getCell('C3').value = 'Feb';
wsx.getCell('D3').value = 'Mar';
for (const c of ['A3', 'B3', 'C3', 'D3']) wsx.getCell(c).font = { bold: true };
const rows = [
  ['Paid search', 4120, 4360, 4590],
  ['Events', 2000, 2000, 3000],
  ['Content', 1500, 1500, 1800],
];
rows.forEach((r, i) => {
  r.forEach((v, j) => {
    wsx.getCell(3 + i + 1, j + 1).value = v;
  });
});
const xlsxBytes = Buffer.from(await wb.xlsx.writeBuffer());

/* --- run --- */
await send('Page.enable');
await send('Page.reload');
await wait(2500);
await shot('home');

await openB64('annual-report.pdf', pdfBytes.toString('base64'), 'pdf');
await wait(2600);
await shot('pdf');

await openB64('qbr.docx', docxBytes.toString('base64'), 'docx');
await wait(2600);
await shot('docx');

await openB64('q1-budget.xlsx', xlsxBytes.toString('base64'), 'sheet');
await wait(3200);
await shot('sheet');

await evaluate(`window.__cellDeskDevApp; document.querySelector('.chrome-home-btn')?.click()`);
await wait(700);
await shot('home-with-resume');

ws.close();
process.exit(0);
