/* Pure conversion between Block[] model and Tiptap ProseMirror JSON.
   No editor imports — unit-testable in node. Run-level formatting rides on
   textStyle marks (bold/italic/underline/strike/color/fontFamily/fontSize/
   highlight/subscript/superscript); links are Link marks; tabs are an atom node. */

import type { JSONContent } from '@tiptap/react';
import type { Block, Cell, DocxDoc, Para, Run, Table } from './docxModel';

export const CD_PARA = 'paragraph';
export const CD_TAB = 'cdTab';

type Mark = { type: string; attrs?: Record<string, unknown> };

/* ---------- Block[] -> PM JSON ---------- */

function runMarks(r: Run, hrefOfRel: (relId?: string) => string | undefined): Mark[] {
  const marks: Mark[] = [];
  const p = r.props || {};
  if (p.b) marks.push({ type: 'bold' });
  if (p.i) marks.push({ type: 'italic' });
  if (p.u) marks.push({ type: 'underline' });
  if (p.strike) marks.push({ type: 'strike' });
  if (p.super) marks.push({ type: 'superscript' });
  if (p.sub) marks.push({ type: 'subscript' });
  if (p.color) marks.push({ type: 'textStyle', attrs: { color: p.color } });
  if (p.size) marks.push({ type: 'textStyle', attrs: { fontSize: `${p.size / 2}pt` } });
  if (p.font) marks.push({ type: 'textStyle', attrs: { fontFamily: p.font } });
  if (p.highlight) marks.push({ type: 'highlight', attrs: { color: p.highlight } });
  const rel = r.linkRelId;
  const href = hrefOfRel(rel);
  if (rel || href) marks.push({ type: 'link', attrs: { href: href ?? '', relId: rel ?? null } });
  return marks;
}

function mergeMarks(list: Mark[][]): Mark[] {
  const out = new Map<string, Mark>();
  for (const m of list.flat()) {
    if (m.type === 'textStyle') {
      const prev = out.get('textStyle');
      out.set('textStyle', { type: 'textStyle', attrs: { ...(prev?.attrs || {}), ...(m.attrs || {}) } });
    } else {
      out.set(m.type, m);
    }
  }
  return [...out.values()];
}

function paraAttrsOf(p: Para): JSONContent['attrs'] {
  const props = p.props || {};
  return {
    styleName: props.style && props.style !== 'Normal' ? props.style : null,
    align: props.align === 'both' ? 'justify' : props.align ?? null,
    numId: props.numId ?? null,
    ilvl: props.ilvl ?? 0,
    pageBreak: props.pageBreakBefore === true,
    opaqueXml: null,
  };
}

function runsToInline(runs: Run[], hrefOfRel: (relId?: string) => string | undefined): JSONContent[] {
  const out: JSONContent[] = [];
  for (const r of runs) {
    if (r.breakBefore) out.push({ type: 'hardBreak' });
    if (r.image) {
      out.push({
        type: 'image',
        attrs: { relId: r.image.relId, wpx: r.image.width ?? null, hpx: r.image.height ?? null, src: null },
      });
      continue;
    }
    if (!r.text) continue;
    if (r.text === '\t') {
      out.push({ type: CD_TAB });
      continue;
    }
    out.push({ type: 'text', text: r.text, marks: runMarks(r, hrefOfRel) });
  }
  return out;
}

interface ListGroup {
  type: 'bulletList' | 'orderedList';
  numId: number;
  ilvl: number;
  items: JSONContent[];
}

function headingLevelOf(style: string | undefined | null): number | null {
  const m = style && /^Heading([1-9])$/i.exec(style);
  return m ? Math.min(3, Math.max(1, parseInt(m[1], 10))) : null;
}

/** Group *consecutive* paragraphs sharing numId+ilvl into tiptap lists. */
export function blocksToTiptap(
  blocks: Block[],
  hrefOfRel: (relId?: string) => string | undefined = () => undefined,
  fmtOfNum: (numId: number) => string = () => 'decimal'
): JSONContent[] {
  const paraNode = (p: Para): JSONContent => {
    const level = headingLevelOf(p.props?.style);
    const inline = runsToInline(p.runs || [], hrefOfRel);
    if (level) {
      return inline.length
        ? { type: 'heading', attrs: { level }, content: inline }
        : { type: 'heading', attrs: { level } };
    }
    return inline.length
      ? { type: CD_PARA, attrs: paraAttrsOf(p), content: inline }
      : { type: CD_PARA, attrs: paraAttrsOf(p) };
  };

  const out: JSONContent[] = [];
  let group: ListGroup | null = null;

  const flush = () => {
    if (!group) return;
    out.push({ type: group.type, attrs: { numId: group.numId, ilvl: group.ilvl }, content: group.items });
    group = null;
  };

  for (const b of blocks) {
    if (b.kind === 'para' && b.props?.numId != null && (b.runs?.length ?? 0) > 0) {
      const kind: ListGroup['type'] = fmtOfNum(b.props.numId) === 'bullet' ? 'bulletList' : 'orderedList';
      const ill = b.props.ilvl ?? 0;
      if (group && group.numId === b.props.numId && group.ilvl === ill && group.type === kind) {
        group.items.push({ type: 'listItem', content: [paraNode(b)] });
      } else {
        flush();
        group = { type: kind, numId: b.props.numId, ilvl: ill, items: [{ type: 'listItem', content: [paraNode(b)] }] };
      }
      continue;
    }
    flush();
    if (b.kind === 'para') {
      out.push(paraNode(b));
    } else if (b.kind === 'table') {
      out.push(tableToTiptap(b, hrefOfRel, fmtOfNum));
    } else {
      out.push({
        type: CD_PARA,
        attrs: {
          styleName: null,
          align: null,
          numId: null,
          ilvl: 0,
          pageBreak: false,
          opaqueXml: b.xml,
          opaqueTag: b.tag,
        },
        content: [{ type: 'text', text: b.tag ? `[unsupported element: ${b.tag}]` : '[unsupported element]' }],
      });
    }
  }
  flush();
  return out;
}

function cellToTiptap(
  c: Cell,
  hrefOfRel: (relId?: string) => string | undefined,
  fmtOfNum: (numId: number) => string
): JSONContent {
  const children = blocksToTiptap(c.blocks.length ? c.blocks : [], hrefOfRel, fmtOfNum);
  const acting = children.filter(
    (n) => n.type === CD_PARA || n.type === 'bulletList' || n.type === 'orderedList'
  );
  return { type: 'tableCell', content: acting.length ? acting : [{ type: CD_PARA }] };
}

function tableToTiptap(
  t: Table,
  hrefOfRel: (relId?: string) => string | undefined,
  fmtOfNum: (numId: number) => string
): JSONContent {
  const rows = t.rows.map((r) => ({
    type: 'tableRow',
    content: r.map((c) => cellToTiptap(c, hrefOfRel, fmtOfNum)),
  }));
  return { type: 'table', content: rows };
}

/* ---------- PM JSON -> Block[] ---------- */

interface LinkRef {
  relId: string | null;
  href: string;
}

function markLookupKinds(marks: Mark[] = []): {
  style: Record<string, unknown>;
  link?: Mark;
} {
  const style: Record<string, unknown> = {};
  let link: Mark | undefined;
  for (const m of marks) {
    switch (m.type) {
      case 'bold':
        style.b = true;
        break;
      case 'italic':
        style.i = true;
        break;
      case 'underline':
        style.u = true;
        break;
      case 'strike':
        style.strike = true;
        break;
      case 'superscript':
        style.super = true;
        break;
      case 'subscript':
        style.sub = true;
        break;
      case 'highlight':
        if (m.attrs?.color) style.highlight = String(m.attrs.color);
        break;
      case 'textStyle': {
        const a = m.attrs || {};
        if (a.color) style.color = a.color;
        if (a.fontSize) {
          const pt = parseFloat(String(a.fontSize));
          if (Number.isFinite(pt)) style.size = Math.round(pt * 2);
        }
        if (a.fontFamily) style.font = String(a.fontFamily).replace(/^['"]|['"]$/g, '');
        break;
      }
      case 'link':
        link = m;
        break;
    }
  }
  return { style, link };
}

function inlineToRuns(content: JSONContent[] | undefined, linkOfMark: (m: Mark | undefined) => LinkRef | null): {
  runs: Run[];
} {
  const runs: Run[] = [];
  let pendingBreak = false;
  for (const n of content || []) {
    if (n.type === 'hardBreak') {
      pendingBreak = true;
      continue;
    }
    if (n.type === CD_TAB) {
      runs.push({ text: '\t', props: {}, breakBefore: pendingBreak });
      pendingBreak = false;
      continue;
    }
    if (n.type === 'image') {
      runs.push({
        text: '',
        props: {},
        breakBefore: pendingBreak,
        image: {
          relId: String(n.attrs?.relId ?? ''),
          width: typeof n.attrs?.wpx === 'number' ? n.attrs.wpx : undefined,
          height: typeof n.attrs?.hpx === 'number' ? n.attrs.hpx : undefined,
        },
      });
      pendingBreak = false;
      continue;
    }
    if (n.type === 'text') {
      const { style, link } = markLookupKinds(n.marks as Mark[]);
      const ref = linkOfMark(link);
      runs.push({
        text: String(n.text ?? ''),
        props: style as Run['props'],
        breakBefore: pendingBreak,
        linkRelId: ref?.relId ?? undefined,
      });
      pendingBreak = false;
    }
  }
  return { runs };
}

function paraPropsOf(attrs: Record<string, unknown> | undefined): Para['props'] {
  const a = attrs || {};
  const alignRaw = a.align;
  const align =
    alignRaw === 'justify' ? 'both' : alignRaw === 'left' || alignRaw === 'center' || alignRaw === 'right' ? alignRaw : undefined;
  const numIdRaw = a.numId == null ? undefined : Number(a.numId);
  const ilvlRaw = a.ilvl == null ? undefined : Number(a.ilvl);
  const numId = typeof numIdRaw === 'number' && Number.isFinite(numIdRaw) ? numIdRaw : undefined;
  const ilvl = typeof ilvlRaw === 'number' && Number.isFinite(ilvlRaw) ? ilvlRaw : undefined;
  return {
    style: typeof a.styleName === 'string' ? a.styleName : undefined,
    align,
    numId,
    ilvl: ilvl !== undefined && ilvl > 0 ? ilvl : ilvl === 0 ? 0 : undefined,
    pageBreakBefore: a.pageBreak === true,
  };
}

export function tiptapToBlocks(
  json: JSONContent,
  linkOfMark: (m: Mark | undefined) => LinkRef | null = () => null
): Block[] {
  return nodesToBlocks(json.content || [], linkOfMark);
}

interface BlockUnfolder {
  newBlocks: Block[];
  linkOfMark: (m: Mark | undefined) => LinkRef | null;
}

function unfoldParagraph(n: JSONContent, env: BlockUnfolder, props: Para['props']): void {
  if (n.attrs?.opaqueXml) {
    env.newBlocks.push({
      kind: 'opaque',
      tag: String(n.attrs.opaqueTag ?? 'unknown'),
      xml: String(n.attrs.opaqueXml),
    });
    return;
  }
  const { runs } = inlineToRuns(n.content, env.linkOfMark);
  env.newBlocks.push({ kind: 'para', props, runs });
}

function unfoldList(n: JSONContent, env: BlockUnfolder): void {
  const numId = Number(n.attrs?.numId ?? 0);
  const ilvl = Number(n.attrs?.ilvl ?? 0);
  for (const li of n.content || []) {
    for (const inner of li.content || []) {
      if (inner.type === 'bulletList' || inner.type === 'orderedList') {
        unfoldList(inner, { ...env });
      } else if (inner.type === CD_PARA) {
        unfoldParagraph(inner, env, { ...paraPropsOf(inner.attrs), numId, ilvl });
      }
    }
  }
}

function unfoldTable(n: JSONContent, env: BlockUnfolder): void {
  const rows: Cell[][] = [];
  for (const tr of n.content || []) {
    const cells: Cell[] = [];
    for (const node of tr.content || []) {
      const blocks = nodesToBlocks(node.content || [], env.linkOfMark);
      cells.push({ blocks: blocks.length ? blocks : [{ kind: 'para', props: {}, runs: [] }] });
    }
    if (cells.length) rows.push(cells);
  }
  env.newBlocks.push({ kind: 'table', rows } as Table);
}

function nodesToBlocks(nodes: JSONContent[], linkOfMark: (m: Mark | undefined) => LinkRef | null): Block[] {
  const env: BlockUnfolder = { newBlocks: [], linkOfMark };
  for (const n of nodes) {
    switch (n.type) {
      case CD_PARA:
        unfoldParagraph(n, env, paraPropsOf(n.attrs));
        break;
      case 'heading': {
        const level = Math.min(3, Math.max(1, Number(n.attrs?.level ?? n.attrs?.styleName?.slice(-1) ?? 1)));
        unfoldParagraph(
          { ...n, type: CD_PARA },
          env,
          {
            ...paraPropsOf(n.attrs),
            style: typeof n.attrs?.styleName === 'string' ? n.attrs.styleName : `Heading${level}`,
          }
        );
        break;
      }
      case 'bulletList':
      case 'orderedList':
        unfoldList(n, env);
        break;
      case 'table':
        unfoldTable(n, env);
        break;
      case 'listItem':
        for (const inner of n.content || []) {
          if (inner.type === CD_PARA) unfoldParagraph(inner, env, paraPropsOf(inner.attrs));
          else if (inner.type === 'bulletList' || inner.type === 'orderedList') unfoldList(inner, env);
        }
        break;
      default:
        break;
    }
  }
  return env.newBlocks;
}

/** Scan a PM tree for paragraph attrs that preserve unedited OOXML. */
export function docFromJson(json: JSONContent, keepSectPr: string | undefined): DocxDoc {
  return { blocks: tiptapToBlocks(json), sectPr: keepSectPr };
}
