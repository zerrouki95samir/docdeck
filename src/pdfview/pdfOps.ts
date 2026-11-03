import {
  PDFDocument,
  PDFCheckBox,
  PDFDropdown,
  PDFRadioGroup,
  PDFTextField,
  StandardFonts,
  degrees,
  rgb,
  type PDFPage,
  type PDFFont,
  type RGB,
} from 'pdf-lib';

/* Coordinate convention: stamps hold PDF-page content-space points (bottom-left origin),
   taken from pdf.js viewport.convertToPdfPoint() at gesture time — that mapping already
   accounts for the page's current /Rotate, so stamps are exact for the visible state.
   Page ops (rotate/delete/insert/reorder/merge/extract) are baked eagerly into working
   bytes with pdf-lib; stamps and form fills overlay lazily and flush into the bytes on
   save/flip (single load+save, then they become document content). */

export type StampType = 'highlight' | 'rect' | 'ellipse' | 'line' | 'freehand' | 'text';

export interface Stamp {
  page: number;
  type: StampType;
  pts: Array<{ x: number; y: number }>; // content-space points
  color?: string; // '#RRGGBB'
  width?: number; // stroke width pt
  fontSize?: number;
  text?: string;
}

export interface FormFieldInfo {
  name: string;
  type: 'text' | 'checkbox' | 'dropdown' | 'radio' | 'other';
  value?: string;
  options?: string[];
}

export async function newBlankPdf(width = 612, height = 792, pageCount = 1): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pageCount; i++) doc.addPage([width, height]);
  return u8(await doc.save());
}

export async function pageCountOf(bytes: Uint8Array): Promise<number> {
  const doc = await PDFDocument.load(alive(bytes));
  return doc.getPageCount();
}

export async function applyPageRotation(bytes: Uint8Array, pageIdx: number, deltaDeg: number): Promise<Uint8Array> {
  const doc = await PDFDocument.load(alive(bytes));
  const page = doc.getPage(pageIdx);
  const cur = ((page.getRotation().angle || 0) + deltaDeg + 360) % 360;
  page.setRotation(degrees(cur));
  return u8(await doc.save());
}

export async function deletePages(bytes: Uint8Array, target: number | number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(alive(bytes));
  const indices = (Array.isArray(target) ? target : [target]).filter(
    (i) => i >= 0 && i < doc.getPageCount()
  );
  [...indices]
    .sort((a, b) => b - a)
    .forEach((i) => doc.removePage(i));
  return u8(await doc.save());
}

export async function insertBlankPage(
  bytes: Uint8Array,
  afterIndex: number, // -1 = append at end
  width: number,
  height: number
): Promise<Uint8Array> {
  const doc = await PDFDocument.load(alive(bytes));
  if (typeof (doc as unknown as { insertPage?: unknown }).insertPage === 'function') {
    const spot = afterIndex < 0 ? doc.getPageCount() : afterIndex + 1;
    (doc as unknown as { insertPage: (i: number, w: [number, number]) => PDFPage }).insertPage(spot, [width, height]);
  } else {
    const blankDoc = await PDFDocument.create();
    blankDoc.addPage([width, height]);
    const [copied] = await doc.copyPages(blankDoc, [0]);
    const spot = afterIndex < 0 ? doc.getPageCount() : afterIndex + 1;
    doc.insertPage(spot, copied);
  }
  return u8(await doc.save());
}

export async function reorderPages(bytes: Uint8Array, newOrder: number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(alive(bytes));
  const out = await PDFDocument.create();
  const copied = await out.copyPages(doc, newOrder);
  copied.forEach((p) => out.addPage(p));
  return u8(await out.save());
}

export async function extractPages(bytes: Uint8Array, indices: number[]): Promise<Uint8Array> {
  const doc = await PDFDocument.load(alive(bytes));
  const out = await PDFDocument.create();
  const safe = indices.filter((i) => i >= 0 && i < doc.getPageCount());
  const copied = await out.copyPages(doc, safe);
  copied.forEach((p) => out.addPage(p));
  return u8(await out.save());
}

export async function mergePdfs(dstBytes: Uint8Array, sources: Uint8Array[]): Promise<Uint8Array> {
  const out = await PDFDocument.load(alive(dstBytes));
  for (const s of sources) {
    const other = await PDFDocument.load(alive(s));
    const copied = await out.copyPages(other, other.getPageIndices());
    copied.forEach((p) => out.addPage(p));
  }
  return u8(await out.save());
}

/* flush stamps + form values into the working bytes */
export async function flushEdits(
  bytes: Uint8Array,
  stamps: Stamp[],
  formValues: Record<string, string> = {}
): Promise<Uint8Array> {
  if (!stamps.length && !Object.keys(formValues).length) return bytes;
  const doc = await PDFDocument.load(alive(bytes));
  const helv = await doc.embedFont(StandardFonts.Helvetica);
  const pages = doc.getPages();
  for (const st of stamps) {
    const page = pages[Math.min(st.page, pages.length - 1)];
    if (!page) continue;
    try {
      drawStamp(page, st, helv);
    } catch {
      /* non-latin text or geometry problems: skip this stamp rather than fail the file */
    }
  }
  if (Object.keys(formValues).length) {
    const form = doc.getForm();
    for (const [name, value] of Object.entries(formValues)) {
      try {
        const field = form.getFieldMaybe(name);
        if (field instanceof PDFTextField) field.setText(value);
        else if (field instanceof PDFCheckBox) {
          if (value === 'true') field.check();
          else field.uncheck();
        } else if (field instanceof PDFDropdown) field.select(value);
        else if (field instanceof PDFRadioGroup) field.select(value);
      } catch {
        /* skip malformed field */
      }
    }
    form.updateFieldAppearances();
  }
  return u8(await doc.save());
}

function drawStamp(page: PDFPage, st: Stamp, helv: PDFFont) {
  const color = rgbFromHex(st.color || '#1e4db7');
  const size = st.width || 2;
  const pts = st.pts;
  if (!pts.length) return;
  switch (st.type) {
    case 'highlight':
    case 'rect': {
      const a = pts[0];
      const b = pts[pts.length - 1];
      page.drawRectangle({
        x: Math.min(a.x, b.x),
        y: Math.min(a.y, b.y),
        width: Math.abs(b.x - a.x),
        height: Math.abs(b.y - a.y),
        color: st.type === 'highlight' ? rgb(1, 0.6, 0) : undefined,
        opacity: st.type === 'highlight' ? 0.35 : undefined,
        borderColor: st.type === 'rect' ? color : undefined,
        borderWidth: st.type === 'rect' ? Math.max(0.5, size) : undefined,
      });
      void color;
      break;
    }
    case 'ellipse': {
      const a = pts[0];
      const b = pts[pts.length - 1];
      page.drawEllipse({
        x: (a.x + b.x) / 2,
        y: (a.y + b.y) / 2,
        xScale: Math.abs(b.x - a.x) / 2 || 12,
        yScale: Math.abs(b.y - a.y) / 2 || 12,
        borderColor: color,
        borderWidth: size,
      });
      break;
    }
    case 'line':
    case 'freehand': {
      if (st.type === 'line') {
        const a = pts[0];
        const b = pts[pts.length - 1];
        page.drawLine({ start: a, end: b, thickness: size, color });
      } else {
        for (let i = 1; i < pts.length; i++) {
          page.drawLine({ start: pts[i - 1], end: pts[i], thickness: size, color });
        }
      }
      break;
    }
    case 'text': {
      page.drawText((st.text || '').slice(0, 2000), {
        x: pts[0].x,
        y: pts[0].y,
        size: st.fontSize || 14,
        font: helv,
        color,
      });
      break;
    }
  }
}

/* ---------- read helpers ---------- */

export async function listFormFields(bytes: Uint8Array): Promise<FormFieldInfo[]> {
  const doc = await PDFDocument.load(alive(bytes));
  const form = doc.getForm();
  const out: FormFieldInfo[] = [];
  for (const f of form.getFields()) {
    const name = f.getName();
    let info: FormFieldInfo = { name, type: 'other' };
    try {
      if (f instanceof PDFTextField) {
        info = { name, type: 'text', value: f.getText() || '' };
      } else if (f instanceof PDFCheckBox) {
        info = { name, type: 'checkbox', value: String(f.isChecked()) };
      } else if (f instanceof PDFDropdown) {
        info = { name, type: 'dropdown', value: f.getSelected()?.[0] || '', options: f.getOptions() || [] };
      } else if (f instanceof PDFRadioGroup) {
        info = { name, type: 'radio', value: f.getSelected() || '', options: f.getOptions() || [] };
      }
    } catch {
      /* unreadable field: still list it as other */
    }
    out.push(info);
  }
  return out;
}

/* ---------- utils ---------- */

/**
 * pdf.js transfers (detaches) the ArrayBuffer you hand it; every pdf.js entry must get a
 * private copy. pdf-lib loads retaining the buffer are fine, but keeping one `alive()` helper
 * makes the rule obvious at each call site.
 */
export function alive(bytes: Uint8Array): Uint8Array {
  return new Uint8Array(bytes).slice();
}

export function u8(x: Uint8Array | ArrayBuffer): Uint8Array {
  return x instanceof Uint8Array ? new Uint8Array(x) : new Uint8Array(x.slice(0) as ArrayBuffer);
}

function rgbFromHex(hex: string): RGB {
  const h = hex.replace('#', '');
  const v = h.length === 3 ? h.split('').map((c) => c + c).join('') : h;
  const num = parseInt(v, 16);
  return rgb(((num >> 16) & 255) / 255, ((num >> 8) & 255) / 255, (num & 255) / 255);
}

export type { PDFPage, PDFFont };
