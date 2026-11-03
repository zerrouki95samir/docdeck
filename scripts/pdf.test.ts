import {
  PDFDocument,
  StandardFonts,
} from 'pdf-lib';
import {
  newBlankPdf,
  pageCountOf,
  applyPageRotation,
  deletePages,
  insertBlankPage,
  reorderPages,
  extractPages,
  mergePdfs,
  flushEdits,
  listFormFields,
  type Stamp,
} from '../src/pdfview/pdfOps';

let passed = 0;
let failed = 0;

function ok(cond: unknown, label: string) {
  if (cond) {
    passed++;
    console.log(`  ok - ${label}`);
  } else {
    failed++;
    console.log(`  FAIL - ${label}`);
  }
}

async function makeSamplePdf(pages = 3, withForm = true): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  doc.addPage([612, 792]);
  if (pages > 1) doc.addPage([500, 700]);
  for (let i = 2; i < pages; i++) doc.insertPage(i, [600, 800]);
  doc.getPages().forEach((p, i) => {
    p.drawText(`Page ${i + 1} Hello`, { x: 70, y: 700 - i * 30, size: 20, font });
  });
  if (withForm) {
    const form = doc.getForm();
    const page0 = doc.getPage(0);
    const name = form.createTextField('name');
    name.setText('');
    name.addToPage(page0, { x: 70, y: 600, width: 160, height: 20 });
    const tick = form.createCheckBox('agree');
    tick.uncheck();
    tick.addToPage(page0, { x: 70, y: 560, width: 14, height: 14 });
  }
  return new Uint8Array(await doc.save());
}

async function main() {
  console.log('pdf-ops tests');
  const p1 = await newBlankPdf();
  ok(p1.length > 400, 'blank pdf produced');
  ok((await pageCountOf(p1)) === 1, 'blank pdf has 1 page');

  const base = await makeSamplePdf(4);
  ok((await pageCountOf(base)) === 4, 'sample pdf has 4 pages');

  // form fill
  const filled = await flushEdits(base, [], { name: 'Martha Tunisian', agree: 'true' });
  const fields = await listFormFields(filled);
  ok(fields.some((f) => f.name === 'name' && f.value === 'Martha Tunisian'), 'text field persisted');
  ok(fields.some((f) => f.name === 'agree' && f.value === 'true'), 'checkbox persisted');

  // stamps
  const stamps: Stamp[] = [
    { page: 0, type: 'highlight', pts: [{ x: 60, y: 760 }, { x: 220, y: 735 }] },
    { page: 0, type: 'rect', pts: [{ x: 60, y: 700 }, { x: 200, y: 640 }], color: '#1e4db7', width: 2 },
    { page: 1, type: 'ellipse', pts: [{ x: 90, y: 600 }, { x: 220, y: 520 }] },
    { page: 1, type: 'line', pts: [{ x: 80, y: 500 }, { x: 300, y: 560 }] },
    { page: 2, type: 'freehand', pts: [{ x: 80, y: 400 }, { x: 120, y: 430 }, { x: 180, y: 410 }, { x: 260, y: 440 }] },
    { page: 3, type: 'text', pts: [{ x: 90, y: 640 }], text: 'Approved by DocDeck', color: '#1e4db7' },
  ];
  const stamped = await flushEdits(base, stamps, { name: 'Sample O. Person' });
  ok(stamped.length > 1000, 'stamp flush produced bytes');
  const nameBack = await listFormFields(stamped);
  ok(nameBack.find((f) => f.name === 'name')?.value === 'Sample O. Person', 'form fill alongside stamps');

  // page rotation
  const rotRes = await applyPageRotation(base, 1, 90);
  const rotDoc = await PDFDocument.load(new Uint8Array(rotRes));
  ok(rotDoc.getPages()[1].getRotation().angle === 90, 'rotation baked');

  // page delete
  const delRes = await deletePages(base, [2]);
  ok((await pageCountOf(delRes)) === 3, 'delete page reduces count');

  // blank insert
  const insRes = await insertBlankPage(base, 1, 500, 600);
  ok((await pageCountOf(insRes)) === 5, 'blank page inserted');

  // reorder
  const reorderRes = await reorderPages(base, [2, 1, 0, 3]);
  const reorderDoc = await PDFDocument.load(new Uint8Array(reorderRes));
  ok(reorderDoc.getPageCount() === 4, 'reorder keeps count');

  // extract
  const extractRes = await extractPages(base, [1, 3]);
  ok((await pageCountOf(extractRes)) === 2, 'extract pages');

  // merge
  const other = await makeSamplePdf(2, false);
  const merged = await mergePdfs(base, [other]);
  ok((await pageCountOf(merged)) === 6, 'merge appends pages');

  // full edit chain over a fresh doc
  const chain = await flushEdits(
    await applyPageRotation(await deletePages(await insertBlankPage(base, 0, 612, 792), 0), 2, 180),
    [{ page: 0, type: 'text', pts: [{ x: 50, y: 300 }], text: 'Working chain' }],
    { agree: 'true' }
  );
  ok((await pageCountOf(chain)) > 0, 'chain edit still valid');
  const chainFields = await listFormFields(chain);
  ok(chainFields.some((f) => f.name === 'agree' && f.value === 'true'), 'chain keep form values');

  console.log(`\n=== RESULT: ${passed} passed, ${failed} failed ===\n`);
  process.exit(failed ? 1 : 0);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
