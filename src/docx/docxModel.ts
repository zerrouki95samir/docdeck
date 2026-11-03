/* OOXML document model for DocDeck.
   Parse word/document.xml body into Block[]; serialize back into fresh <w:body> chunk.
   Unknown constructs are carried verbatim (Opaque) so unedited parts of any real
   document survive round-trips. One nesting level for tables keeps cell editing simple. */

export interface RunProps {
  b?: boolean;
  i?: boolean;
  u?: boolean;
  strike?: boolean;
  color?: string; // 6-hex RRGGBB
  size?: number; // half-points
  font?: string;
  super?: boolean;
  sub?: boolean;
  highlight?: string;
}

export interface Run {
  text: string;
  props: RunProps;
  linkRelId?: string; // w:hyperlink r:id
  image?: ImgRef;
  breakBefore?: boolean; // <w:br/> right before this run's text
}
export interface ImgRef {
  relId: string;
  width?: number; // px at 96dpi (extent cx / 9525)
  height?: number;
}

export interface ParaProps {
  style?: string;
  align?: 'left' | 'center' | 'right' | 'both';
  numId?: number; // 0 = explicitly boxed-off list
  ilvl?: number;
  pageBreakBefore?: boolean;
}

export interface Para {
  kind: 'para';
  props: ParaProps;
  runs: Run[];
}

export interface Cell {
  blocks: Block[];
  coverFwd?: boolean; // this cell absorbs horizontal span
}
export interface Table {
  kind: 'table';
  rows: Cell[][];
}
export interface Opaque {
  kind: 'opaque';
  tag: string;
  xml: string;
}
export type Block = Para | Table | Opaque;

export interface DocxDoc {
  blocks: Block[];
  sectPr?: string;
}

/* =============================== parse =============================== */

const W_NS = 'w';

function isOn(elm: Element | null | undefined): boolean {
  if (!elm) return false;
  const v = elm.getAttribute('w:val');
  return v !== '0' && v !== 'false' && v !== 'none';
}

export function parseRunProps(rpr: Element | null): RunProps {
  const p: RunProps = {};
  if (!rpr) return p;
  const child = (n: string): Element | null => {
    for (const c of Array.from(rpr.children)) if (c.localName === n) return c;
    return null;
  };
  if (isOn(child('b'))) p.b = true;
  if (isOn(child('i'))) p.i = true;
  if (isOn(child('u'))) p.u = true;
  if (isOn(child('strike'))) p.strike = true;
  const col = child('color')?.getAttribute('w:val');
  if (col && col.toLowerCase() !== 'auto') p.color = col;
  const sz = child('sz')?.getAttribute('w:val');
  if (sz && /^\d+$/.test(sz)) p.size = parseInt(sz, 10);
  const f = child('rFonts');
  const font = f?.getAttribute('w:ascii') || f?.getAttribute('w:hAnsi');
  if (font) p.font = font;
  const hl = child('highlight')?.getAttribute('w:val');
  if (hl && hl !== 'none') p.highlight = hl;
  const va = child('vertAlign')?.getAttribute('w:val');
  if (va === 'superscript') p.super = true;
  if (va === 'subscript') p.sub = true;
  return p;
}

function childNamed(node: Element, n: string): Element | null {
  for (const c of Array.from(node.children)) if (c.localName === n) return c;
  return null;
}

export function parseParaProps(p: Element): ParaProps {
  const pr = childNamed(p, 'pPr');
  if (!pr) return {};
  const style = childNamed(pr, 'pStyle')?.getAttribute('w:val');
  const jc = childNamed(pr, 'jc')?.getAttribute('w:val');
  const align: ParaProps['align'] | undefined =
    jc === undefined
      ? undefined
      : jc === 'center'
        ? 'center'
        : jc === 'right' || jc === 'end'
          ? 'right'
          : jc === 'both'
            ? 'both'
            : 'left';
  const numPr = childNamed(pr, 'numPr');
  let numId: number | undefined;
  let ilvl: number | undefined;
  if (numPr) {
    const nid = childNamed(numPr, 'numId')?.getAttribute('w:val');
    const lvl = childNamed(numPr, 'ilvl')?.getAttribute('w:val');
    const n1 = nid ? parseInt(nid, 10) : NaN;
    const n2 = lvl ? parseInt(lvl, 10) : NaN;
    numId = Number.isFinite(n1) ? n1 : undefined;
    ilvl = Number.isFinite(n2) ? n2 : undefined;
  }
  const pageBreakBefore = isOn(childNamed(pr, 'pageBreakBefore'));
  return {
    style: style || undefined,
    align,
    numId,
    ilvl,
    pageBreakBefore,
  };
}

function imgRefOf(r: Element): ImgRef | undefined {
  const blip = r.querySelector('a\\:blip, blip');
  const relId = blip?.getAttribute('r:embed');
  if (!relId) return undefined;
  const extent = r.querySelector('wp\\:extent, extent');
  const cx = extent?.getAttribute('cx');
  const cy = extent?.getAttribute('cy');
  return {
    relId,
    width: cx ? Math.round(parseInt(cx, 10) / 9525) : undefined,
    height: cy ? Math.round(parseInt(cy, 10) / 9525) : undefined,
  };
}

export function parseRuns(p: Element): Run[] {
  const runs: Run[] = [];
  const walk = (node: Element, inheritedBreak: boolean): boolean => {
    let brk = inheritedBreak;
    for (const ch of Array.from(node.children)) {
      if (ch.localName === 'r') {
        const props = parseRunProps(childNamed(ch, 'rPr'));
        for (const piece of Array.from(ch.children)) {
          if (piece.localName === 'br') {
            brk = true;
            continue;
          }
          if (piece.localName === 't') {
            const text = piece.textContent ?? '';
            if (!text) continue;
            runs.push({ text, props, breakBefore: brk });
            brk = false;
          } else if (piece.localName === 'tab') {
            runs.push({ text: '\t', props, breakBefore: brk });
            brk = false;
          } else if (piece.localName === 'drawing' || piece.localName === 'pict') {
            const img = imgRefOf(ch);
            if (img) {
              runs.push({ text: '', props, image: img, breakBefore: brk });
              brk = false;
            }
          }
        }
      } else if (ch.localName === 'hyperlink') {
        const relId = ch.getAttribute('r:id') || undefined;
        const from = runs.length;
        brk = walk(ch, brk);
        if (relId) for (let k = from; k < runs.length; k++) runs[k].linkRelId = relId;
      } else if (ch.localName === 'smartTag' || ch.localName === 'ins' || ch.localName === 'sdt') {
        brk = walk(ch, brk);
      }
    }
    return brk;
  };
  walk(p, false);
  return runs;
}

export function parsePara(p: Element): Para {
  return { kind: 'para', props: parseParaProps(p), runs: parseRuns(p) };
}

export function parseTable(t: Element): Table {
  const rows: Cell[][] = [];
  for (const tr of Array.from(t.children).filter((c) => c.localName === 'tr')) {
    const cells: Cell[] = [];
    for (const tc of Array.from(tr.children).filter((c) => c.localName === 'tc')) {
      cells.push({ blocks: parseChildren(tc) });
    }
    if (cells.length) rows.push(cells);
  }
  return { kind: 'table', rows };
}

export function parseChildren(node: Element): Block[] {
  const blocks: Block[] = [];
  for (const node2 of Array.from(node.children)) {
    if (node2.localName === 'p') blocks.push(parsePara(node2));
    else if (node2.localName === 'tbl') blocks.push(parseTable(node2));
  }
  return blocks;
}

export function parseBody(body: Element): DocxDoc {
  const blocks: Block[] = [];
  let sectPr: string | undefined;
  for (const node of Array.from(body.children)) {
    switch (node.localName) {
      case 'p':
        blocks.push(parsePara(node));
        break;
      case 'tbl':
        blocks.push(parseTable(node));
        break;
      case 'sectPr':
        sectPr = node.outerHTML;
        break;
      default:
        blocks.push({ kind: 'opaque', tag: node.localName || '', xml: node.outerHTML });
    }
  }
  return { blocks, sectPr };
}

/* ============================== serialize ============================== */

let imageSeqCounter = 9000;

function runPrXml(pr: RunProps): string {
  const parts: string[] = [];
  if (pr.b) parts.push('<w:b/>');
  if (pr.i) parts.push('<w:i/>');
  if (pr.u) parts.push('<w:u w:val="single"/>');
  if (pr.strike) parts.push('<w:strike/>');
  if (pr.color) parts.push(`<w:color w:val="${pr.color}"/>`);
  if (pr.size) parts.push(`<w:sz w:val="${pr.size}"/>`);
  if (pr.font) parts.push(`<w:rFonts w:ascii="${esc(pr.font)}" w:hAnsi="${esc(pr.font)}"/>`);
  if (pr.highlight) parts.push(`<w:highlight w:val="${pr.highlight}"/>`);
  if (pr.super) parts.push('<w:vertAlign w:val="superscript"/>');
  if (pr.sub) parts.push('<w:vertAlign w:val="subscript"/>');
  return parts.length ? `<w:rPr>${parts.join('')}</w:rPr>` : '';
}

export function runXml(r: Run): string {
  const pieces: string[] = [runPrXml(r.props)];
  if (r.breakBefore) pieces.push('<w:br/>');
  if (r.image) {
    pieces.push(imageXml(r.image));
  } else {
    const segs = r.text.split('\t');
    segs.forEach((s, i) => {
      if (i) pieces.push('<w:tab/>');
      if (s) pieces.push(`<w:t xml:space="preserve">${esc(s)}</w:t>`);
    });
  }
  const inner = `<w:r>${pieces.join('')}</w:r>`;
  return r.linkRelId ? `<w:hyperlink r:id="${esc(r.linkRelId)}">${inner}</w:hyperlink>` : inner;
}

export function imageXml(img: ImgRef): string {
  const id = ++imageSeqCounter;
  const cx = Math.max(1, Math.round((img.width || 200) * 9525));
  const cy = Math.max(1, Math.round((img.height || 200) * 9525));
  return (
    `<w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
    `<wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${cx}" cy="${cy}"/>` +
    `<wp:docPr id="${id}" name="Image${id}"/>` +
    `<wp:cNvGraphicFramePr/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="${id}" name="image${id}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${esc(img.relId)}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`
  );
}

export function paraXml(p: Para): string {
  const pr: string[] = [];
  if (p.props.style) pr.push(`<w:pStyle w:val="${esc(p.props.style)}"/>`);
  if (p.props.align)
    pr.push(
      `<w:jc w:val="${p.props.align === 'left' ? 'start' : p.props.align === 'right' ? 'end' : p.props.align === 'both' ? 'both' : 'center'}"/>`
    );
  if (p.props.numId != null)
    pr.push(`<w:numPr><w:ilvl w:val="${p.props.ilvl ?? 0}"/><w:numId w:val="${p.props.numId}"/></w:numPr>`);
  if (p.props.pageBreakBefore) pr.push('<w:pageBreakBefore/>');
  const runs = (p.runs || []).map(runXml).join('');
  return `<w:p>${pr.length ? `<w:pPr>${pr.join('')}</w:pPr>` : ''}${runs}</w:p>`;
}

export function tableXml(t: Table): string {
  const rows = t.rows
    .map((r) => `<w:tr>${r.map((c) => cellXml(c)).join('')}</w:tr>`)
    .join('');
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>` +
    `<w:top w:val="single" w:sz="4" w:color="C9C9C9"/>` +
    `<w:left w:val="single" w:sz="4" w:color="C9C9C9"/>` +
    `<w:bottom w:val="single" w:sz="4" w:color="C9C9C9"/>` +
    `<w:right w:val="single" w:sz="4" w:color="C9C9C9"/>` +
    `<w:insideH w:val="single" w:sz="4" w:color="DDDDDD"/>` +
    `<w:insideV w:val="single" w:sz="4" w:color="DDDDDD"/>` +
    `</w:tblBorders></w:tblPr>${rows}</w:tbl>`
  );
}

export function cellXml(c: Cell): string {
  const inner = c.blocks.length ? c.blocks.map(blockXml).join('') : '<w:p/>';
  return `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr>${inner}</w:tc>`;
}

export function blockXml(b: Block): string {
  switch (b.kind) {
    case 'para':
      return paraXml(b);
    case 'table':
      return tableXml(b);
    default:
      return b.xml;
  }
}

/** Build the new body chunk: blocks + verbatim trailing sectPr. */
export function bodyXml(doc: DocxDoc): string {
  return doc.blocks.map(blockXml).join('') + (doc.sectPr ?? '');
}

export function esc(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/* Convenience helpers used by the editor UI */

export function newPara(runs: Run[] = [], props: ParaProps = {}): Para {
  return { kind: 'para', props, runs };
}

export function emptyRuns(text = ''): Run[] {
  return text ? [{ text, props: {} }] : [];
}

export function moveBlock<T>(arr: T[], from: number, to: number): T[] {
  const copy = [...arr];
  const [x] = copy.splice(from, 1);
  copy.splice(to, 0, x);
  return copy;
}

export const W = W_NS;
