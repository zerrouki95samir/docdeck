import { JSDOM } from 'jsdom';

const dom = new JSDOM('<!doctype html><html><body></body></html>');
(globalThis as unknown as { DOMParser: typeof DOMParser }).DOMParser = dom.window.DOMParser as unknown as typeof DOMParser;

import JSZip from 'jszip';
import { newDocxBytes, loadDocx, saveDocx, addImage } from '../src/docx/docxIO';
import { parseBody, bodyXml, newPara, type Para, type Table } from '../src/docx/docxModel';

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

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

/* 1x1 red PNG (real bytes so createImageBitmap-equivalent flows stay valid) */
const PNG_1PX = new Uint8Array(
  atob(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='
  )
    .split('')
    .map((c) => c.charCodeAt(0))
);

async function main() {
  console.log('docx model tests');

  // --- template zip ---
  const blank = await newDocxBytes();
  ok(blank.length > 200, 'blank docx produced');
  const zs = await JSZip.loadAsync(blank);
  const files = Object.keys(zs.files);
  ok(files.includes('word/document.xml'), 'template has document.xml');
  ok(files.includes('word/styles.xml'), 'template has styles.xml');
  ok(files.includes('[Content_Types].xml'), 'template has content types');

  // --- load + save round trip of the blank template ---
  const src = await loadDocx(blank);
  ok(src.doc.blocks.length >= 1, 'blank doc has a paragraph');
  const saved = await saveDocx(
    {
      ...src,
      doc: {
        ...src.doc,
        blocks: [newPara([{ text: 'Hello Cellular', props: { b: true } }], { style: 'Heading1' })],
      },
    },
    []
  );
  const src2 = await loadDocx(saved);
  const p0 = src2.doc.blocks[0];
  ok(p0.kind === 'para' && p0.runs[0]?.text === 'Hello Cellular', 'text round trip');
  ok(p0.kind === 'para' && p0.runs[0]?.props.b === true, 'bold round trip');
  ok(p0.kind === 'para' && p0.props.style === 'Heading1', 'style round trip');
  ok(Boolean(src2.doc.sectPr && src2.doc.sectPr!.includes('pgSz')), 'sectPr preserved');

  // --- synthetic document.xml with rich content ---
  const rich = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <w:body>
  <w:p>
    <w:pPr><w:pStyle w:val="Title"/><w:jc w:val="center"/></w:pPr>
    <w:r><w:rPr><w:b/><w:color w:val="C00000"/><w:sz w:val="36"/><w:rFonts w:ascii="Georgia"/></w:rPr><w:t xml:space="preserve">Rich Document</w:t></w:r>
  </w:p>
  <w:p>
    <w:r><w:t>has</w:t><w:tab/><w:t>tabs</w:t><w:br/><w:t>and breaks</w:t></w:r>
    <w:hyperlink r:id="rIdLink1"><w:r><w:rPr><w:u/></w:rPr><w:t>a link</w:t></w:r></w:hyperlink>
  </w:p>
  <w:tbl>
    <w:tr><w:tc><w:p><w:r><w:t>cell A1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>cell B1</w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:p><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:inline><wp:extent cx="475000" cy="285000"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rIdImage7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
  <w:sectPr><w:pgSz w:w="11907" w:h="16839"/></w:sectPr>
 </w:body>
</w:document>`;

  const docBody = parseXml(rich).getElementsByTagName('w:body')[0];
  const parsed = parseBody(docBody);
  ok(parsed.blocks.length === 4, 'body parsed into 4 blocks');

  const title = parsed.blocks[0] as Para;
  ok(title.kind === 'para' && title.props.style === 'Title', 'title style parsed');
  ok(title.kind === 'para' && title.props.align === 'center', 'align center parsed');
  const titleRun = title.kind === 'para' ? title.runs[0] : undefined;
  ok(titleRun?.text === 'Rich Document', 'title text');
  ok(
    Boolean(titleRun?.props.b && titleRun?.props.color === 'C00000' && titleRun?.props.size === 36),
    'title run props (bold, color, size)'
  );
  ok(titleRun?.props.font === 'Georgia', 'title font parsed');

  const mixed = parsed.blocks[1] as Para;
  const mruns = mixed.runs;
  ok(mruns.some((r) => r.text === 'tabs'), 'tab-separated piece retained');
  ok(mruns.some((r) => r.text === 'sand breaks' === false && r.breakBefore === true && r.text === 'and breaks'), 'line break flag');
  const lastRun = mruns[mruns.length - 1];
  ok(lastRun?.linkRelId === 'rIdLink1', 'hyperlink rel captured');
  ok(lastRun?.props.u === true, 'hyperlink underline captured');

  const tbl = parsed.blocks[2] as Table;
  ok(tbl.kind === 'table' && tbl.rows[0]?.length === 2, 'table row cells parsed');
  ok(
    tbl.kind === 'table' &&
      ((tbl.rows[0][0].blocks[0] as Para).runs[0]?.text === 'cell A1'),
    'table cell text parsed'
  );

  const imgPara = parsed.blocks[3] as Para;
  ok(imgPara.kind === 'para' && imgPara.runs[0]?.image?.relId === 'rIdImage7', 'image rel captured');
  ok(
    imgPara.kind === 'para' && imgPara.runs[0]?.image?.width === 50,
    'image extent converted (475000 EMU → 50px)'
  );

  // --- serialize back ---
  const out = bodyXml(parsed);
  ok(out.includes('<w:b/>'), 'serialized bold');
  ok(out.includes('w:val="C00000"'), 'serialized color');
  ok(out.includes('w:val="36"'), 'serialized size');
  ok(out.includes('rFonts w:ascii="Georgia"'), 'serialized font');
  ok(out.includes('<w:hyperlink r:id="rIdLink1">'), 'serialized hyperlink');
  ok(out.includes('r:embed="rIdImage7"'), 'serialized image');
  ok(out.includes('<w:tab/>'), 'serialized tab');
  ok(out.includes('<w:br/>'), 'serialized break');
  ok(out.includes('w:val="Title"'), 'serialized style');
  ok(out.includes('>cell A1<'), 'serialized table cell text');
  ok(out.includes('w:w="11907"'), 'sectPr page size preserved in body chunk');

  // serialization re-parses consistently
  const roundDoc = parseXml(
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${out}</w:body></w:document>`
  );
  const round = parseBody(roundDoc.getElementsByTagName('w:body')[0]);
  const rtitle = round.blocks[0] as Para;
  ok(rtitle.props.style === 'Title' && rtitle.runs[0]?.props.color === 'C00000', 're-parse keeps style/props');
  const rmixed = round.blocks[1] as Para;
  ok(rmixed.runs[rmixed.runs.length - 1]?.linkRelId === 'rIdLink1', 're-parse keeps link');
  ok(round.blocks[2]?.kind === 'table', 're-parse keeps table');

  // --- image insertion through docxIO (media + rels + content types) ---
  const s = await loadDocx(blank);
  const relId = await addImage(s, PNG_1PX, 'png', 100, 80);
  ok(relId.startsWith('rId'), 'addImage returns a rel id');
  const outBytes = await saveDocx({ ...s, doc: s.doc }, []);
  const z2 = await JSZip.loadAsync(outBytes);
  const mediaNames = Object.keys(z2.files).filter((n) => n.startsWith('word/media/') && !n.endsWith('/'));
  ok(mediaNames.length === 1, 'media entry added');
  const relsEntry2 = z2.file('word/_rels/document.xml.rels')!;
  const rels2 = await relsEntry2.async('string');
  ok(rels2.includes(relId) && rels2.includes('media/image1.png'), 'relationship added');
  const typesEntry2 = z2.file('[Content_Types].xml')!;
  const types2 = await typesEntry2.async('string');
  ok(types2.includes('Extension="png"'), 'content type default added');

  console.log(`\n=== RESULT: ${passed} passed, ${failed} failed ===\n`);
  process.exit(failed ? 1 : 0);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
