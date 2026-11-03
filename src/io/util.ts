export function randomId(prefix = 'u'): string {
  return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 10)}`;
}

/** Excel ARGB color ('FFFF8800') -> Univer css color (#FF8800). Theme/indexed colors fall back. */
export function argbToCss(color: unknown): string | undefined {
  if (!color) return undefined;
  if (typeof color === 'string') {
    let s = color.trim();
    if (/^[0-9a-fA-F]{8}$/.test(s)) s = s.slice(2);
    if (/^[0-9a-fA-F]{6}$/.test(s)) return `#${s.toLowerCase()}`;
    return undefined;
  }
  if (typeof color === 'object') {
    const anyColor = color as { argb?: string; rgb?: string };
    if (anyColor.argb) return argbToCss(anyColor.argb);
    if (anyColor.rgb) return argbToCss(anyColor.rgb);
  }
  return undefined;
}

/** Univer css color (#rgb / #rrggbb) -> Excel ARGB (alpha FF). */
export function cssToArgb(css?: string | null): string | undefined {
  if (!css) return undefined;
  let hex = css.trim();
  if (hex.startsWith('#')) hex = hex.slice(1);
  let alpha = 'FF';
  if (hex.length === 8) {
    alpha = hex.slice(0, 2);
    hex = hex.slice(2);
  }
  if (hex.length === 3) hex = hex.split('').map((c) => c + c).join('');
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return undefined;
  return `${alpha}${hex.toUpperCase()}`;
}

const EPOCH_1900_DIFF = 25569; // days between 1899-12-30 and 1970-01-01
const EPOCH_1904_DIFF = 24207; // days between 1904-01-01 and 1970-01-01  (25569 - 1462)

/** JS Date -> Excel serial (based on workbook date system). Uses UTC parts to match ExcelJS parsing. */
export function dateToExcelSerial(d: Date, date1904 = false): number {
  const days = (d.getTime() - new Date(Date.UTC(1899, 11, 30)).getTime()) / 86400000;
  return date1904 ? days + 1462 : days;
}

/** Excel serial -> ms epoch used for CSV/Uni export display decisions. */
export function excelSerialToDateUTC(serial: number, date1904 = false): Date {
  const offset = date1904 ? EPOCH_1904_DIFF : EPOCH_1900_DIFF;
  return new Date((serial - offset) * 86400000);
}

export function numDaysToIso(d: Date): string {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/** chars (Excel width unit) <-> px (Univer). ~7px per char + 5px padding */
export function colCharsToPx(w: number): number {
  return Math.round(w * 7 + 5);
}
export function colPxToChars(px: number): number {
  return Math.max(0, (px - 5) / 7);
}
/** points (Excel row height) <-> px (Univer) */
export function rowPtToPx(pt: number): number {
  return Math.round(pt * (4 / 3) * 100) / 100;
}
export function rowPxToPt(px: number): number {
  return Math.round(px * 0.75 * 100) / 100;
}

function colLettersToIndex(letters: string): number {
  let col = 0;
  for (const ch of letters.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return col - 1;
}

/** 'A1:E1' string -> 0-based {top,left,bottom,right}; also validates object form. */
export function parseA1Range(
  ref: string | { top?: number; left?: number; bottom?: number; right?: number }
): { top: number; left: number; bottom: number; right: number } | null {
  if (typeof ref !== 'string') {
    if (ref && Number.isFinite(ref.top) && Number.isFinite(ref.left) && Number.isFinite(ref.bottom) && Number.isFinite(ref.right)) {
      return {
        top: Math.min(ref.top!, ref.bottom!),
        left: Math.min(ref.left!, ref.right!),
        bottom: Math.max(ref.top!, ref.bottom!),
        right: Math.max(ref.left!, ref.right!),
      };
    }
    return null;
  }
  const m = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/.exec(ref.trim());
  if (!m) return null;
  const c1 = colLettersToIndex(m[1]);
  const r1 = parseInt(m[2], 10) - 1;
  const c2 = m[3] ? colLettersToIndex(m[3]) : c1;
  const r2 = m[4] ? parseInt(m[4], 10) - 1 : r1;
  return { top: Math.min(r1, r2), left: Math.min(c1, c2), bottom: Math.max(r1, r2), right: Math.max(c1, c2) };
}

/** Shared workbook dimension guards */
export const MAX_ROWS = 100000;
export const MAX_COLS = 500;
