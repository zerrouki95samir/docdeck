import type { CellStyleDT } from './dto';
import { argbToCss, cssToArgb } from './util';

/* eslint-disable @typescript-eslint/no-explicit-any */
type Any = any;

const EJS_TO_UV_BORDER: Record<string, number> = {
  thin: 1,
  hair: 2,
  dotted: 3,
  dashed: 4,
  dashDot: 5,
  dashDotDot: 6,
  double: 7,
  medium: 8,
  mediumDashed: 9,
  mediumDashDot: 10,
  mediumDashDotDot: 11,
  slantDashDot: 12,
  thick: 13,
};
const UV_TO_EJS_BORDER: Record<number, string> = Object.fromEntries(
  Object.entries(EJS_TO_UV_BORDER).map(([k, v]) => [v, k])
);

/** ExcelJS cell -> Univer style object (no numfmt; numfmt is read separately) */
export function findStyleObject(cell: { style?: Any }): CellStyleDT | undefined {
  const s: Any = cell.style;
  if (!s) return undefined;
  const out: CellStyleDT = {};
  const font = s.font;
  if (font) {
    if (font.bold) out.bl = 1;
    if (font.italic) out.it = 1;
    if (font.underline) out.ul = { s: 1, t: font.underline === 'double' ? 1 : 0 };
    if (font.strike) out.st = { s: 1, t: 0 };
    if (font.size && font.size !== 11) out.fs = font.size;
    if (font.name && font.name !== 'Calibri') out.ff = font.name;
    const rgb = argbToCss(font.color);
    if (rgb) out.cl = { rgb };
  }
  if (s.fill && s.fill.pattern === 'solid' && s.fill.fgColor?.argb) {
    const rgb = argbToCss(s.fill.fgColor.argb);
    if (rgb) out.bg = { rgb };
  }
  const bd = new Map<string, { s: number; cl?: { rgb: string } }>();
  const side = (ejsKey: 'top' | 'bottom' | 'left' | 'right', uvKey: string) => {
    const v: Any = s.border?.[ejsKey];
    if (!v || !v.style) return;
    const sv = EJS_TO_UV_BORDER[v.style];
    if (sv == null) return;
    const rgb = argbToCss(v.color);
    bd.set(uvKey, { s: sv, ...(rgb ? { cl: { rgb } } : {}) });
  };
  side('top', 't');
  side('bottom', 'b');
  side('left', 'l');
  side('right', 'r');
  if (bd.size) out.bd = Object.fromEntries(bd) as CellStyleDT['bd'];

  const a = s.alignment;
  if (a) {
    if (a.horizontal === 'left') out.ht = 1;
    else if (a.horizontal === 'center' || a.horizontal === 'centerContinuous') out.ht = 2;
    else if (a.horizontal === 'right') out.ht = 3;
    if (a.vertical === 'top') out.vt = 1;
    else if (a.vertical === 'middle') out.vt = 2;
    else if (a.vertical === 'bottom') out.vt = 3;
    if (a.wrapText) out.tb = 3;
    if (a.rotation != null && a.rotation !== 0) {
      if (a.rotation === 255) out.tr = { a: 0, v: 1 };
      else {
        const deg = a.rotation < 0 ? 360 + a.rotation : a.rotation;
        out.tr = { a: deg % 360, v: 0 };
      }
    }
  }

  const keys: Array<keyof CellStyleDT> = ['bl', 'it', 'ul', 'st', 'fs', 'ff', 'cl', 'bg', 'bd', 'ht', 'vt', 'tb', 'tr'];
  if (keys.every((k) => out[k] == null)) return undefined;
  return out;
}

/** ---------- Univer -> ExcelJS ---------- */

export function univerStyleToExcel(cell: { style: Any }, style: CellStyleDT | undefined, pattern?: string): void {
  if (!style) return;
  const weStyle = cell.style;

  if (style.bl === 1 || style.it === 1 || style.fs != null || style.ff != null || style.cl?.rgb || style.ul?.s || style.st?.s) {
    const font: Any = { ...(weStyle.font || {}) };
    if (style.bl === 1) font.bold = true;
    if (style.it === 1) font.italic = true;
    if (style.ul?.s === 1) font.underline = style.ul.t === 1 ? 'double' : true;
    if (style.st?.s === 1) font.strike = true;
    if (style.fs != null) font.size = style.fs;
    if (style.ff != null) font.name = style.ff;
    const argb = cssToArgb(style.cl?.rgb);
    if (argb) font.color = { argb };
    weStyle.font = font;
  }

  if (style.bg?.rgb) {
    const argb = cssToArgb(style.bg.rgb);
    if (argb) weStyle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb } };
  }

  if (style.bd) {
    const b = (style.bd as Record<string, { s?: number; cl?: { rgb?: string } }>) || {};
    const side = (uv: { s?: number; cl?: { rgb?: string } } | undefined, key: string) => {
      if (!uv || uv.s == null || uv.s === 0) return;
      const jsStyle = UV_TO_EJS_BORDER[uv.s] || 'thin';
      const border: Any = { style: jsStyle };
      const argb = cssToArgb(uv.cl?.rgb);
      if (argb) border.color = { argb };
      weStyle.border = weStyle.border || {};
      weStyle.border[key] = border;
    };
    side(b.t, 'top');
    side(b.b, 'bottom');
    side(b.l, 'left');
    side(b.r, 'right');
  }

  const align: Any = { ...(weStyle.alignment || {}) };
  if (style.ht === 1) align.horizontal = 'left';
  else if (style.ht === 2) align.horizontal = 'center';
  else if (style.ht === 3) align.horizontal = 'right';
  if (style.vt === 1) align.vertical = 'top';
  else if (style.vt === 2) align.vertical = 'middle';
  else if (style.vt === 3) align.vertical = 'bottom';
  if (style.tb === 3) align.wrapText = true;
  if (style.tr?.v === 1) align.rotation = 255;
  else if (style.tr?.a) align.rotation = style.tr.a;
  if (Object.values(align).some((v) => v != null)) weStyle.alignment = align;

  if (pattern && pattern !== 'General') weStyle.numFmt = pattern;
}
