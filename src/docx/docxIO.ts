import JSZip from 'jszip';
import { bodyXml, parseBody, type DocxDoc } from './docxModel';

export interface DocxSource {
  zip: JSZip;
  headerOpen: string; // everything up to and including <w:body>
  footerClose: string; // </w:document> and any after-body remainder
  doc: DocxDoc;
  mediaList: string[]; // existing media paths
  relMap: Map<string, string>; // rId -> target (internal or external URL)
  numFmtMap: Map<number, string>; // numId -> numFmt w:val (level 0)
}

const DOC_XML = 'word/document.xml';
const RELS_XML = 'word/_rels/document.xml.rels';
const TYPES_XML = '[Content_Types].xml';
const NUMBERING_XML = 'word/numbering.xml';

/* ---------- relationships / numbering readers ---------- */

function readRelMap(zip: JSZip, rels: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
    const tag = m[0];
    const id = tag.match(/\bId="([^"]+)"/)?.[1];
    const target = tag.match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) map.set(id, target);
  }
  if (zip.file(RELS_XML)) {
    void zip.file(RELS_XML); // keep entry memoized for later writes
  }
  return map;
}

async function readNumbering(zip: JSZip): Promise<Map<number, string>> {
  const out = new Map<number, string>();
  const entry = zip.file(NUMBERING_XML);
  if (!entry) return out;
  try {
    const xml = new DOMParser().parseFromString(await entry.async('string'), 'application/xml');
    if (xml.querySelector('parsererror')) return out;
    const NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const abstractFmt = new Map<string, string>();
    for (const ab of xml.getElementsByTagNameNS(NS, 'abstractNum')) {
      const id = ab.getAttribute('w:abstractNumId');
      const lvl0 = ab.getElementsByTagNameNS(NS, 'lvl')[0];
      const fmt = lvl0?.getElementsByTagNameNS(NS, 'numFmt')[0]?.getAttribute('w:val');
      if (id && fmt) abstractFmt.set(id, fmt);
    }
    for (const num of xml.getElementsByTagNameNS(NS, 'num')) {
      const numId = num.getAttribute('w:numId');
      const abstractId = num.getElementsByTagNameNS(NS, 'abstractNumId')[0]?.getAttribute('w:val');
      if (numId && abstractId) {
        const fmt = abstractFmt.get(abstractId) || 'bullet';
        out.set(parseInt(numId, 10), fmt);
      }
    }
  } catch {
    /* keep whatever parsed */
  }
  return out;
}

export async function loadDocx(bytes: Uint8Array): Promise<DocxSource> {
  const zip = await JSZip.loadAsync(bytes.slice(0));
  const docEntry = zip.file(DOC_XML);
  if (!docEntry) throw new Error('not a Word document (missing word/document.xml)');
  const raw = await docEntry.async('string');
  const cut = raw.indexOf('<w:body>');
  const endCut = raw.lastIndexOf('</w:document>');
  if (cut < 0 || endCut < 0) throw new Error('malformed document.xml');
  const headerOpen = raw.slice(0, cut + '<w:body>'.length);
  const footerClose = raw.slice(endCut);

  const xml = new DOMParser().parseFromString(raw, 'application/xml');
  if (xml.querySelector('parsererror')) throw new Error('document.xml parse error');
  const body = xml.getElementsByTagName('w:body')[0];
  if (!body) throw new Error('document.xml has no body');

  const mediaList = Object.keys(zip.files).filter((n) => n.startsWith('word/media/'));

  const relsEntry = zip.file(RELS_XML);
  const rels = relsEntry ? await relsEntry.async('string') : '';
  const relMap = readRelMap(zip, rels);
  const numFmtMap = await readNumbering(zip);

  return { zip, headerOpen, footerClose, doc: parseBody(body), mediaList, relMap, numFmtMap };
}

export async function saveDocx(
  source: DocxSource,
  extraFiles: Array<[string, Uint8Array | string]> = []
): Promise<Uint8Array> {
  const { zip, headerOpen, footerClose, doc } = source;
  const updated = `${headerOpen}${bodyXml(doc)}</w:body>${footerClose}`;
  zip.file(DOC_XML, updated);
  extraFiles.forEach(([path, data]) => zip.file(path, data));
  return await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

/* ---------- images ---------- */

const REL_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

export async function addImage(
  source: DocxSource,
  bytes: Uint8Array,
  extNoDot: string,
  width: number,
  height: number
): Promise<string> {
  // 1. media entry
  let n = source.mediaList.length + 1;
  let name = `image${n}.${extNoDot}`;
  while (source.zip.file(`word/media/${name}`)) {
    n += 1;
    name = `image${n}.${extNoDot}`;
  }
  source.mediaList.push(`word/media/${name}`);
  source.zip.file(`word/media/${name}`, bytes);

  // 2. relationship
  const relsEntry = source.zip.file(RELS_XML);
  if (!relsEntry) throw new Error('missing document relationships');
  const rels = await relsEntry.async('string');
  const nums = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
  const maxId = nums.length ? Math.max(...nums) : 0;
  const relId = `rId${maxId + 1}`;
  const updatedRels = rels.replace('</Relationships>', `<Relationship Id="${relId}" Type="${REL_IMAGE}" Target="media/${name}"/></Relationships>`);
  source.zip.file(RELS_XML, updatedRels);

  // 3. content type (Default extension)
  const typesEntry = source.zip.file(TYPES_XML);
  if (!typesEntry) throw new Error('missing content types');
  const types = await typesEntry.async('string');
  const ext = extNoDot.toLowerCase();
  const mime = ext === 'jpg' ? 'jpeg' : ext;
  if (!new RegExp(`Extension="${ext}"`, 'i').test(types)) {
    source.zip.file(
      TYPES_XML,
      types.replace('</Types>', `<Default Extension="${ext}" ContentType="image/${mime}"/></Types>`)
    );
  }
  return relId;
}

/* ---------- numbering (bullet & numbered lists) ---------- */

export const BULLET_NUM_ID = 91;
export const DECIMAL_NUM_ID = 92;

/** Make sure a numbering part exists with our reserved ids for new lists.
    Real documents opened from disk already carry their own numbering. */
export async function ensureNumbering(source: DocxSource): Promise<void> {
  const entry = source.zip.file(NUMBERING_XML);
  if (entry) {
    const xml = await entry.async('string');
    const hasIds =
      xml.includes(`w:numId="${BULLET_NUM_ID}"`) && xml.includes(`w:numId="${DECIMAL_NUM_ID}"`);
    if (!hasIds) {
      source.zip.file(
        NUMBERING_XML,
        xml.replace('</w:numbering>', `${abstracts()}${reservedNums()}</w:numbering>`)
      );
    }
    await addNumberingType(source);
    await addNumberingRel(source);
    return;
  }
  source.zip.file(
    NUMBERING_XML,
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${abstracts()}${reservedNums()}</w:numbering>`
  );
  await addNumberingType(source);
  await addNumberingRel(source);
}

function abstracts(): string {
  return (
    `<w:abstractNum w:abstractNumId="8001"><w:multiLevelType w:val="hybridMultilevel"/>` +
    lvl(0, 'bullet', '•') +
    lvl(1, 'bullet', '◦') +
    lvl(2, 'bullet', '▪') +
    `</w:abstractNum>` +
    `<w:abstractNum w:abstractNumId="8002"><w:multiLevelType w:val="hybridMultilevel"/>` +
    lvl(0, 'decimal', '%1.') +
    lvl(1, 'decimal', '%2.') +
    lvl(2, 'decimal', '%3.') +
    `</w:abstractNum>`
  );
}

function reservedNums(): string {
  return (
    `<w:num w:numId="${BULLET_NUM_ID}"><w:abstractNumId w:val="8001"/></w:num>` +
    `<w:num w:numId="${DECIMAL_NUM_ID}"><w:abstractNumId w:val="8002"/></w:num>`
  );
}

function lvl(ilvl: number, fmt: string, text: string): string {
  return (
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
    `<w:lvlText w:val="${text === '•' ? '&#8226;' : text}"/>` +
    `<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 + ilvl * 720}" w:hanging="360"/></w:pPr>` +
    `</w:lvl>`
  );
}

async function addNumberingType(source: DocxSource): Promise<void> {
  const typesEntry = source.zip.file(TYPES_XML);
  if (!typesEntry) return;
  const types = await typesEntry.async('string');
  if (!types.includes('numbering+xml')) {
    source.zip.file(
      TYPES_XML,
      types.replace(
        '</Types>',
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>'
      )
    );
  }
}

async function addNumberingRel(source: DocxSource): Promise<void> {
  const relsEntry = source.zip.file(RELS_XML);
  if (!relsEntry) return;
  const rels = await relsEntry.async('string');
  if (rels.includes('numbering.xml')) return;
  const nums = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
  const relId = `rId${(nums.length ? Math.max(...nums) : 0) + 1}`;
  source.zip.file(
    RELS_XML,
    rels.replace(
      '</Relationships>',
      `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>`
    )
  );
}

/* ---------- external links ---------- */

const REL_HYPERLINK =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

/** Register an external URL relationship and return the new r:id */
export async function addExternalLink(source: DocxSource, href: string): Promise<string> {
  const relsEntry = source.zip.file(RELS_XML);
  if (!relsEntry) throw new Error('missing document relationships');
  const rels = await relsEntry.async('string');
  const nums = [...rels.matchAll(/Id="rId(\d+)"/g)].map((m) => parseInt(m[1], 10));
  const relId = `rId${(nums.length ? Math.max(...nums) : 0) + 1}`;
  source.zip.file(
    RELS_XML,
    rels.replace(
      '</Relationships>',
      `<Relationship Id="${relId}" Type="${REL_HYPERLINK}" Target="${href.replace(/&/g, '&amp;')}" TargetMode="External"/></Relationships>`
    )
  );
  return relId;
}

/* ---------- media previews ---------- */

export async function mediaDataUrl(source: DocxSource, relId: string): Promise<string | null> {
  const target = source.relMap.get(relId);
  if (!target) return null;
  const name = `word/${target.replace(/^\/?word\//, '')}`;
  const entry = source.zip.file(name) || source.zip.file(target);
  if (!entry) return null;
  const bytes = await entry.async('uint8array');
  const ext = (target.split('.').pop() || 'png').toLowerCase();
  const mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : ext === 'webp' ? 'image/webp' : 'image/png';
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return `data:${mime};base64,${btoa(bin)}`;
}

/* ---------- blank template ---------- */

const T = (s: string) => s;

export function blankDocxParts(): Array<[string, string]> {
  const documentXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><w:body>` +
    `${T('<w:p/>')}` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>` +
    `</w:body></w:document>`;
  const styles =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault/></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
    [...Array(3).keys()].map(
      (i) =>
        `<w:style w:type="paragraph" w:styleId="Heading${i + 1}"><w:name w:val="heading ${i}"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="${i}"/></w:pPr><w:rPr><w:b/><w:color w:val="1F3864"/><w:sz w:val="${[32, 28, 24][i]}"/></w:rPr></w:style>`
    ).join('') +
    `</w:styles>`;
  const types =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `</Types>`;
  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;
  const docRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`;

  return [
    ['[Content_Types].xml', types],
    ['_rels/.rels', rootRels],
    ['word/document.xml', documentXml],
    ['word/styles.xml', styles],
    ['word/_rels/document.xml.rels', docRels],
  ];
}

export async function newDocxBytes(): Promise<Uint8Array> {
  const zip = new JSZip();
  blankDocxParts().forEach(([path, text]) => zip.file(path, text));
  return await zip.generateAsync({ type: 'uint8array' });
}

export { saveDocx as saveDocxZip, RELS_XML };
