import { forwardRef, useEffect, useImperativeHandle, useMemo, useRef, useState, type ReactNode } from 'react';
import { EditorContent, useEditor, type Editor } from '@tiptap/react';
import type { JSONContent } from '@tiptap/core';
import { EditorState } from '@tiptap/pm/state';
import type { Node as ProseNode } from '@tiptap/pm/model';
import { renderAsync } from 'docx-preview';
import {
  loadDocx,
  saveDocx,
  newDocxBytes,
  addImage,
  addExternalLink,
  ensureNumbering,
  mediaDataUrl,
  type DocxSource,
} from './docxIO';
import { blocksToTiptap, tiptapToBlocks } from './blocksToTiptap';
import { buildDocxExtensions } from './cdDocxKit';
import type { LensHandle, LensProps } from '../lenses';

type Mark = { type: string; attrs?: Record<string, unknown> };

const FONT_LIST = ['Calibri', 'Times New Roman', 'Arial', 'Georgia', 'Courier New', 'Cambria'];
const SIZE_LIST = [16, 18, 20, 22, 24, 28, 32, 36, 48, 56, 72]; // half-points
const COLOR_LIST = [
  ['Black', '000000'],
  ['Dark Red', 'C00000'],
  ['Red', 'E23D2C'],
  ['Navy', '1F3864'],
  ['Blue', '2F6FED'],
  ['Green', '375623'],
  ['Purple', '7030A0'],
  ['Gray', '595959'],
];

const DEFAULT_MARKS: Mark[] = [];

/* Custom ribbon combo — native <select> popups are killed instantly by the
   editor's re-render churn in this WebView, and popovers give full styling. */
interface RbnOption {
  value: string;
  label: string;
  swatch?: string;
  fontFamily?: string;
  preview?: React.CSSProperties;
  on?: boolean;
  disabled?: boolean;
}

function RbnSelect(props: {
  title: string;
  value: string;
  placeholder?: string;
  width?: number;
  menuWidth?: number;
  options: RbnOption[];
  renderValue?: () => ReactNode;
  onChange: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const boxRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent) => {
      if (boxRef.current && !boxRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setOpen(false);
    };
    document.addEventListener('mousedown', onDoc, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('mousedown', onDoc, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [open]);

  const current = props.options.find((o) => o.value === props.value);
  return (
    <div className="cd-sel" ref={boxRef} style={props.width ? { width: props.width } : undefined}>
      <button
        type="button"
        className="cd-sel-btn"
        title={props.title}
        onMouseDown={(e) => e.preventDefault()}
        onClick={() => setOpen((o) => !o)}
      >
        <span className="cd-sel-label">{props.renderValue ? props.renderValue() : current?.label ?? props.placeholder ?? props.value}</span>
        <span className="cd-sel-caret">▾</span>
      </button>
      {open && (
        <div className="cd-sel-menu" style={props.menuWidth ? { minWidth: props.menuWidth } : undefined}>
          {props.options.map((o) => (
            <div
              key={o.value + '\u0000' + o.label}
              className={`cd-sel-item ${o.on || (!o.disabled && props.value === o.value) ? 'sel-on' : ''} ${o.disabled ? 'sel-off' : ''}`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                if (o.disabled) return;
                setOpen(false);
                props.onChange(o.value);
              }}
            >
              {o.swatch !== undefined && <span className="cd-sel-swatch" style={{ background: '#' + o.swatch }} />}
              <span className="cd-sel-item-label" style={{ ...(o.fontFamily ? { fontFamily: o.fontFamily } : {}), ...(o.preview || {}) }}>
                {o.label}
              </span>
              <span className="cd-sel-check">{o.on || (!o.disabled && props.value === o.value) ? '✓' : ''}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

export const DocxLens = forwardRef<LensHandle, LensProps>(function DocxLens(props, ref) {
  const propsRef = useRef(props);
  propsRef.current = props;

  const imgInputRef = useRef<HTMLInputElement>(null);
  const colorInputRef = useRef<HTMLInputElement>(null);

  const srcRef = useRef<DocxSource | null>(null);
  const nameRef = useRef('Untitled.docx');
  const pathRef = useRef<string | null>(null);
  const dirtyRef = useRef(false);
  const sectPrRef = useRef<string | undefined>(undefined);
  const applyTokenRef = useRef(0);
  const pendingJsonRef = useRef<JSONContent[] | null>(null);
  const linkCacheRef = useRef(new Map<string, string>()); // href -> relId (new external links)
  const hookRangeRef = useRef<{ from: number; to: number } | null>(null);
  const printHostRef = useRef<HTMLDivElement>(null);

  const [, setTick] = useState(0);
  const bump = () => setTick((t) => t + 1);

  /* ---------- editor ---------- */

  const docxExtensions = useMemo(() => buildDocxExtensions(), []);
  const editor = useEditor({
    extensions: docxExtensions,
    content: '<p></p>',
    editorProps: {
      attributes: {
        spellcheck: 'false',
      },
      handleKeyDown: (_view, event) => {
        if (event.key === 'Escape') {
          (document.activeElement as HTMLElement | null)?.blur();
        }
        return false;
      },
    },
    onUpdate: () => {
      if (!dirtyRef.current) {
        dirtyRef.current = true;
      }
      propsRef.current.onDirty();
      bump();
    },
  });
  const editorRef = useRef<Editor | null>(null);
  editorRef.current = editor;

  /* apply pending JSON whenever the editor is ready (mount race fallback) */
  useEffect(() => {
    if (!editor || !pendingJsonRef.current) return;
    const json = pendingJsonRef.current;
    pendingJsonRef.current = null;
    applyJson(json);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [editor, applyTokenRef.current]);

  const applyPagePadding = () => {
    const mar = sectPrRef.current || '';
    const px = (twips: string | undefined) => Math.round((parseInt(twips || '1440', 10) / 1440) * 96);
    const padTop = px(mar.match(/w:top="(\d+)"/)?.[1]);
    const padBot = px(mar.match(/w:bottom="(\d+)"/)?.[1]);
    const padL = px(mar.match(/w:left="(\d+)"/)?.[1]);
    const padR = px(mar.match(/w:right="(\d+)"/)?.[1]);
    const el = document.querySelector('.cd-docx-surface .ProseMirror') as HTMLElement | null;
    if (el) {
      el.style.paddingTop = `${padTop}px`;
      el.style.paddingBottom = `${padBot}px`;
      el.style.paddingLeft = `${padL}px`;
      el.style.paddingRight = `${padR}px`;
    }
  };

  /* PM forbids empty text nodes — strip any that sneak in from conversion, so
     nodeFromJSON never throws and the structural swap is always used. */
  const sanitizeTiptapJson = (nodes: JSONContent[]): JSONContent[] => {
    const out: JSONContent[] = [];
    for (const n of nodes) {
      const content = n.content ? sanitizeTiptapJson(n.content as JSONContent[]) : undefined;
      if (n.type === 'text' && !(typeof n.text === 'string' ? n.text : '').trim()) continue;
      out.push(content ? { ...n, content } : n);
    }
    return out;
  };

  /* apply content to the editor synchronously when it already exists.
     NOTE: tiptap 2.27's setContent() corrupts JSON block structure (folds every
     paragraph into one) when the editor was initialized with JSON content —
     so we swap the whole document through a raw EditorState instead. */
  const applyJson = (jsonIn: JSONContent[]): boolean => {
    const ed = editorRef.current;
    if (!ed) return false;
    const json = sanitizeTiptapJson(jsonIn);
    let newDoc: ProseNode;
    try {
      newDoc = ed.state.schema.nodeFromJSON({ type: 'doc', content: json });
    } catch {
      ed.commands.setContent({ type: 'doc', content: json } as JSONContent, false);
      applyPagePadding();
      bump();
      return true;
    }
    const state = EditorState.create({ doc: newDoc, plugins: ed.state.plugins });
    ed.view.updateState(state);
    ed.commands.setTextSelection(1);
    applyPagePadding();
    bump();
    return true;
  };

  /* hydrate image src attrs in a JSON tree */
  const hydrateImages = async (nodes: JSONContent[], src: DocxSource, cache: Map<string, string>): Promise<void> => {
    for (const n of nodes) {
      if (n.type === 'image' && n.attrs?.relId && !n.attrs.src) {
        const relId = String(n.attrs.relId);
        if (!cache.has(relId)) cache.set(relId, (await mediaDataUrl(src, relId)) || '');
        n.attrs = { ...n.attrs, src: cache.get(relId) || null };
      }
      if (n.content) await hydrateImages(n.content, src, cache);
    }
  };

  /* ---------- byte pipeline ---------- */

  const buildJsonBlocks = () => {
    const ed = editorRef.current;
    const s = srcRef.current;
    if (!ed || !s) throw new Error('no document loaded');
    const walk = async (nodes: JSONContent[]): Promise<void> => {
      for (const n of nodes) {
        for (const m of (n.marks as Mark[] | undefined) || DEFAULT_MARKS) {
          if (m.type === 'link' && !m.attrs?.relId && m.attrs?.href) {
            const href = String(m.attrs.href);
            if (!linkCacheRef.current.has(href)) {
              const relId = await addExternalLink(s, href);
              linkCacheRef.current.set(href, relId);
            }
          }
        }
        if (n.content) await walk(n.content);
      }
    };
    const json = ed.getJSON();
    const run = async (): Promise<void> => {
      await walk(json.content || []);
      await ensureNumbering(s);
    };
    return run().then(() =>
      tiptapToBlocks(json, (m: Mark | undefined) => {
        if (m?.type !== 'link') return null;
        const rel = (m.attrs?.relId as string) || (m.attrs?.href ? linkCacheRef.current.get(String(m.attrs.href)) : null);
        return rel ? { relId: rel, href: String(m.attrs?.href ?? '') } : null;
      })
    );
  };

  const buildBytes = async (): Promise<Uint8Array> => {
    const s = srcRef.current;
    if (!s) throw new Error('no document loaded');
    const blocks = await buildJsonBlocks();
    return await saveDocx({ ...s, doc: { blocks, sectPr: sectPrRef.current } });
  };

  /* ---------- load helpers ---------- */

  const adoptDoc = (src: DocxSource, name: string, path: string | null) => {
    srcRef.current = src;
    sectPrRef.current = src.doc.sectPr;
    nameRef.current = name;
    pathRef.current = path;
    dirtyRef.current = false;
    linkCacheRef.current = new Map();
    hookRangeRef.current = null;
    propsRef.current.onTitle(name);
    propsRef.current.onDirty();
  };

  /* ---------- imperative API ---------- */

  const selfRef = useRef<LensHandle | null>(null);

  useImperativeHandle<LensHandle, LensHandle>(ref, () => {
    const handle: LensHandle = {
      kind: 'docx',
      async open(name: string, bytes: Uint8Array, path: string | null) {
        try {
          const src = await loadDocx(bytes);
          adoptDoc(src, name, path);
          const json = blocksToTiptap(
            src.doc.blocks,
            (relId?: string) => (relId ? src.relMap.get(relId) : undefined),
            (numId: number) => src.numFmtMap.get(numId) || 'decimal'
          );
          await hydrateImages(json, src, linkCacheRef.current);
          pendingJsonRef.current = json;
          if (!applyJson(json)) {
            applyTokenRef.current += 1;
            bump();
          }
        } catch (err) {
          propsRef.current.onStatus(`Could not read DOCX: ${(err as Error).message}`);
          throw err;
        }
      },
      async save(asNew: boolean) {
        const s = srcRef.current;
        if (!s || !editorRef.current) return false;
        let target = pathRef.current;
        if (!target || asNew) {
          target = await window.cellDesk!.saveDialog({
            defaultPath: pathRef.current || nameRef.current,
            filters: [{ name: 'Word Document', extensions: ['docx'] }],
          });
          if (!target) return false;
        }
        propsRef.current.onBusy('Saving…');
        try {
          const bytes = await buildBytes();
          await window.cellDesk!.writeFile(target, bytes);
          pathRef.current = target;
          nameRef.current = target.split('/').pop() || target;
          dirtyRef.current = false;
          propsRef.current.onTitle(nameRef.current);
          propsRef.current.onDirty();
          propsRef.current.onStatus(`Saved ${nameRef.current}`);
          return true;
        } catch (err) {
          propsRef.current.onStatus(`Save failed: ${(err as Error).message}`);
          return false;
        } finally {
          propsRef.current.onBusy(null);
        }
      },
      async makeNew() {
        try {
          const bytes = await newDocxBytes();
          const src = await loadDocx(bytes);
          adoptDoc(src, 'Untitled.docx', null);
          const json = blocksToTiptap(src.doc.blocks, () => undefined, () => 'decimal');
          await hydrateImages(json, src, linkCacheRef.current);
          pendingJsonRef.current = json;
          if (!applyJson(json)) {
            applyTokenRef.current += 1;
            bump();
          }
        } catch (err) {
          propsRef.current.onStatus(`Could not create document: ${(err as Error).message}`);
          throw err;
        }
      },
      async exportAs() {
        propsRef.current.onStatus('Use Print… to produce a PDF of this document');
      },
      async print() {
        const s = srcRef.current;
        if (!s || !printHostRef.current) return;
        propsRef.current.onBusy('Preparing print…');
        try {
          const bytes = await buildBytes();
          await renderAsync(bytes.slice(0), printHostRef.current, undefined, {
            className: 'docx-preview',
            inWrapper: true,
          });
          setTimeout(() => window.print(), 250);
        } catch (err) {
          propsRef.current.onStatus(`Print failed: ${(err as Error).message}`);
        } finally {
          propsRef.current.onBusy(null);
        }
      },
      isDirty() {
        return dirtyRef.current;
      },
      fileName() {
        return nameRef.current;
      },
    };
    selfRef.current = handle;
    return handle;
  });

  /* ---------- print host (hidden; used only for printing) ---------- */

  useEffect(() => {
    const host = printHostRef.current;
    if (!host) return;
    const reset = () => {
      setTimeout(() => {
        if (host) host.innerHTML = '';
      }, 400);
    };
    window.addEventListener('afterprint', reset);
    return () => window.removeEventListener('afterprint', reset);
  }, []);

  /* ---------- word count ---------- */

  const docTextOf = (): string => {
    const ed = editorRef.current;
    if (!ed) return '';
    try {
      return ed.state.doc.textBetween(0, ed.state.doc.content.size, '\n', '\t');
    } catch {
      return '';
    }
  };

  /* ---------- dev hooks (kept for the E2E drivers) ---------- */

  useEffect(() => {
    if (props.active === false) return;
    const w = window as unknown as Record<string, unknown>;
    const hooks = {
      open: async (name: string, b64: string) => {
        const bin = atob(b64);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        if (!selfRef.current) return null;
        await selfRef.current.open(name, arr, null);
        return { blocks: textBlocks(), text: docTextOf() };
      },
      save: async () => {
        const bytes = await buildBytes();
        let s = '';
        for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
        return btoa(s);
      },
      text: () => docTextOf(),
      addPara: async (text: string, opts?: { style?: string; after?: number; b?: boolean }) => {
        const ed = editorRef.current;
        if (!ed) return '';
        ed.commands.focus('end');
        const inline: JSONContent[] = [
          { type: 'text', text, marks: opts?.b ? [{ type: 'bold' }] : undefined },
        ];
        if (opts?.style && /^Heading[1-3]$/.test(opts.style)) {
          ed.commands.insertContent([
            {
              type: 'heading',
              attrs: { level: parseInt(opts.style.slice(-1), 10) as 1 | 2 | 3 },
              content: inline,
            },
          ]);
        } else {
          ed.commands.insertContent([{ type: 'paragraph', content: inline }]);
        }
        ed.commands.focus('end');
        return docTextOf();
      },
      fmt: async (optsIn?: Record<string, unknown>) => {
        const ed = editorRef.current;
        if (!ed || !hookRangeRef.current) return false;
        const range = hookRangeRef.current;
        ed.commands.setTextSelection(range);
        const ch = ed.chain().focus();
        for (const [k, v] of Object.entries(optsIn || {})) {
          if (v !== true) continue;
          if (k === 'b') ch.setMark('bold');
          else if (k === 'i') ch.setMark('italic');
          else if (k === 'u') ch.setMark('underline');
          else if (k === 'strike') ch.setMark('strike');
          else if (k === 'super') ch.setMark('superscript');
          else if (k === 'sub') ch.setMark('subscript');
        }
        ch.run();
        return true;
      },
      sel: (m: { mode: 'para'; index: number } | null) => {
        const ed = editorRef.current;
        if (!ed || !m || m.mode !== 'para') return false;
        const doc = ed.state.doc;
        let paraIdx = -1;
        let found: { from: number; to: number } | null = null;
        doc.forEach((node, offset) => {
          if (node.type.name === 'paragraph' || node.type.name === 'heading') paraIdx += 1;
          if (found === null && paraIdx === m.index) {
            found = { from: offset + 1, to: offset + node.content.size + 1 };
          }
        });
        if (!found) return false;
        hookRangeRef.current = found;
        ed.commands.setTextSelection(found);
        return true;
      },
      table: async (rowsIn?: number, colsIn?: number) => {
        const ed = editorRef.current;
        if (!ed) return 0;
        const rows = rowsIn || 2;
        const cols = colsIn || 2;
        const cells = Array.from({ length: cols }, () => ({ type: 'tableCell', content: [{ type: 'paragraph' }] }));
        const content = Array.from({ length: rows }, () => ({ type: 'tableRow', content: cells.map((c) => ({ ...c })) }));
        ed.commands.focus('end');
        ed.commands.insertContent([{ type: 'table', content }]);
        let tables = 0;
        ed.state.doc.descendants((n) => {
          if (n.type.name === 'table') tables += 1;
        });
        return tables;
      },
      view: () => ({ editing: true, blocks: editorRef.current?.state.doc.childCount ?? 0 }),
    };
    w.__cellDeskDevDocx = hooks;
    return () => {
      delete w.__cellDeskDevDocx;
    };
  }, [props.active]);

  function textBlocks(): number {
    let count = 0;
    editorRef.current?.state.doc.forEach((node) => {
      if (node.type.name === 'paragraph' || node.type.name === 'heading') count += 1;
    });
    return count;
  }

  /* =================================================================
     render — Word-like ribbon + always-editable page surface
     ================================================================= */

  const ch = editor?.chain().focus();
  const parAttrs = editor ? editor.getAttributes('paragraph') : {};
  const cur = editor ? (editor.getAttributes('textStyle') as Record<string, unknown>) : {};
  const curSize = typeof cur.fontSize === 'string' ? String(Math.round(parseFloat(cur.fontSize) * 2)) : '';
  const curFont = typeof cur.fontFamily === 'string' ? (cur.fontFamily as string) : '';
  const curColor = typeof cur.color === 'string' ? (cur.color as string).toUpperCase() : '';
  const curHiColor = (() => {
    const h = editor ? (editor.getAttributes('highlight') as Record<string, unknown>) : {};
    return typeof h.color === 'string' ? (h.color as string).toUpperCase() : '';
  })();
  const activeHeading = editor?.isActive('heading') ? Number(editor?.getAttributes('heading').level) : 0;
  const curStyleSel = activeHeading ? `H${activeHeading}` : (parAttrs.styleName as string) || 'Normal';
  const inTable = Boolean(editor && editor.isActive('table'));
  const words = docTextOf().replace(/\t/g, ' ').trim().split(/\s+/).filter(Boolean).length;

  const ribbonBtn = (
    label: ReactNode,
    onClick: () => void,
    opts: { on?: boolean; title: string; style?: React.CSSProperties } | string
  ) => {
    const o = typeof opts === 'string' ? { title: opts, on: false } : opts;
    return (
      <button
        className={`cd-rbn-btn ${o.on ? 'tool-on' : ''}`}
        onMouseDown={(e) => e.preventDefault()}
        onClick={onClick}
        title={o.title}
        style={o.style}
      >
        {label}
      </button>
    );
  };

  return (
    <div className="lens-docx">
      <div className="cd-ribbon">
        <div className="cd-rbn-group">
          {ribbonBtn('↶', () => editor?.commands.undo(), 'Undo')}
          {ribbonBtn('↷', () => editor?.commands.redo(), 'Redo')}
        </div>
        <div className="cd-rbn-sep" />
        <RbnSelect
          title="Font"
          width={126}
          value={curFont}
          placeholder="Font"
          onChange={(v) => v && ch?.setFontFamily(v).run()}
          options={FONT_LIST.map((f) => ({ value: f, label: f, fontFamily: f }))}
        />
        <RbnSelect
          title="Font size"
          width={58}
          value={curSize}
          placeholder="Size"
          onChange={(v) => v && ch?.setFontSize(`${parseInt(v, 10) / 2}pt`).run()}
          options={SIZE_LIST.map((s) => ({ value: String(s), label: String(s / 2) }))}
        />
        <div className="cd-rbn-group">
          {ribbonBtn(<b>B</b>, () => editor?.commands.toggleBold(), { on: editor?.isActive('bold'), title: 'Bold' })}
          {ribbonBtn(<i>I</i>, () => editor?.commands.toggleItalic(), { on: editor?.isActive('italic'), title: 'Italic' })}
          {ribbonBtn(
            <u>U</u>,
            () => editor?.commands.toggleUnderline(),
            { on: editor?.isActive('underline'), title: 'Underline' }
          )}
          {ribbonBtn(
            <s>S</s>,
            () => editor?.commands.toggleStrike(),
            { on: editor?.isActive('strike'), title: 'Strikethrough' }
          )}
          {ribbonBtn('x²', () => editor?.commands.toggleSuperscript(), { on: editor?.isActive('superscript'), title: 'Superscript' })}
          {ribbonBtn('x₂', () => editor?.commands.toggleSubscript(), { on: editor?.isActive('subscript'), title: 'Subscript' })}
        </div>
        <div className="cd-rbn-sep" />
        <RbnSelect
          title="Text color"
          width={46}
          value={curColor}
          renderValue={() => (
            <span className="cd-color-chip" style={{ color: curColor ? '#' + curColor : '#111' }}>
              A<i style={{ background: curColor ? '#' + curColor : '#111' }} />
            </span>
          )}
          onChange={(v) => {
            if (v === '__more') {
              colorInputRef.current?.click();
              return;
            }
            if (!v) ch?.unsetColor().run();
            else ch?.setColor(v).run();
          }}
          options={[
            { value: '', label: 'Automatic' },
            ...COLOR_LIST.map(([n, hex]) => ({ value: hex, label: n, swatch: hex })),
            { value: '__more', label: 'More colors…' },
          ]}
        />
        <RbnSelect
          title="Highlight"
          width={46}
          value={curHiColor}
          placeholder="H"
          renderValue={() => (
            <span className="cd-color-chip" style={{ color: '#111' }}>
              H<i style={{ background: curHiColor ? '#' + curHiColor : '#FFE450' }} />
            </span>
          )}
          onChange={(v) => {
            if (!v) ch?.unsetHighlight().run();
            else ch?.setHighlight({ color: v }).run();
          }}
          options={[
            { value: '', label: 'None' },
            ...(
              [
                ['Yellow', 'FFFF00'],
                ['Green', '00FF00'],
                ['Aqua', '00FFFF'],
                ['Magenta', 'FF00FF'],
                ['Gray', 'C0C0C0'],
              ] as [string, string][]
            ).map(([n, hex]) => ({ value: hex, label: n, swatch: hex })),
          ]}
        />
        <div className="cd-rbn-sep" />
        <RbnSelect
          title="Paragraph style"
          width={142}
          value={curStyleSel}
          onChange={(v) => {
            if (v === 'Normal') ch?.setParagraph().updateAttributes('paragraph', { styleName: null }).run();
            else if (/^H[123]$/.test(v)) ch?.setHeading({ level: parseInt(v.slice(1), 10) as 1 | 2 | 3 }).run();
            else ch?.setParagraph().updateAttributes('paragraph', { styleName: v }).run();
          }}
          options={[
            { value: 'Normal', label: 'Normal' },
            { value: 'H1', label: 'Heading 1', preview: { fontSize: 14, fontWeight: 700 } },
            { value: 'H2', label: 'Heading 2', preview: { fontSize: 13, fontWeight: 700, color: '#1f3864' } },
            { value: 'H3', label: 'Heading 3', preview: { fontSize: 12.5, fontWeight: 700, color: '#2f6fed' } },
            { value: 'Quote', label: 'Quote', preview: { fontStyle: 'italic', color: '#595959' } },
            { value: 'IntenseQuote', label: 'Intense Quote', preview: { fontStyle: 'italic', fontWeight: 600, color: '#c00000' } },
            { value: 'ListParagraph', label: 'List Paragraph', preview: { paddingLeft: 6 } },
          ]}
        />
        <div className="cd-rbn-group">
          {ribbonBtn('⯇', () => ch?.setTextAlign('left').run(), { on: editor?.isActive({ textAlign: 'left' }), title: 'Align left', style: { letterSpacing: 0 } })}
          {ribbonBtn('≡', () => ch?.setTextAlign('center').run(), { on: editor?.isActive({ textAlign: 'center' }), title: 'Center' })}
          {ribbonBtn('⯈', () => ch?.setTextAlign('right').run(), { on: editor?.isActive({ textAlign: 'right' }), title: 'Align right' })}
          {ribbonBtn('▤', () => ch?.setTextAlign('justify').run(), { on: editor?.isActive({ textAlign: 'justify' }), title: 'Justify' })}
        </div>
        <div className="cd-rbn-sep" />
        <div className="cd-rbn-group">
          {ribbonBtn(
            '•⁄',
            () => {
              ch?.toggleBulletList().run();
              if (srcRef.current) void ensureNumbering(srcRef.current);
            },
            { on: editor?.isActive('bulletList'), title: 'Bulleted list' }
          )}
          {ribbonBtn(
            '1.',
            () => {
              ch?.toggleOrderedList().run();
              if (srcRef.current) void ensureNumbering(srcRef.current);
            },
            { on: editor?.isActive('orderedList'), title: 'Numbered list' }
          )}
        </div>
        <div className="cd-rbn-sep" />
        <div className="cd-rbn-group">
          {ribbonBtn('🔗', () => {
            const href = window.prompt('Link URL (https://…)');
            if (href) {
              const linkAttrs = { href, relId: null } as unknown as { href: string };
              ch?.extendMarkRange('link').setLink(linkAttrs).run();
            }
          }, 'Add link')}
          {ribbonBtn('✂︎⛓', () => ch?.unsetLink().run(), 'Remove link')}
          {ribbonBtn('🖼', () => imgInputRef.current?.click(), 'Insert image')}
          {ribbonBtn('▦', () => ch?.insertTable({ rows: 3, cols: 3, withHeaderRow: false }).run(), 'Insert table')}
        </div>
        {inTable && (
          <>
            <div className="cd-rbn-sep" />
            <div className="cd-rbn-group">
              <span className="cd-rbn-note">Table:</span>
              {ribbonBtn('+Row', () => ch?.addRowAfter().run(), 'Add row below')}
              {ribbonBtn('+Col', () => ch?.addColumnAfter().run(), 'Add column right')}
              {ribbonBtn('−Row', () => ch?.deleteRow().run(), 'Delete row')}
              {ribbonBtn('−Col', () => ch?.deleteColumn().run(), 'Delete column')}
            </div>
          </>
        )}
        <div style={{ flex: 1 }} />
        <span className="docx-wordcount">{words} words</span>
        {ribbonBtn('🖨', () => void selfRef.current?.print(), 'Print or export to PDF')}
      </div>

      <input
        ref={colorInputRef}
        type="color"
        style={{ position: 'absolute', width: 0, height: 0, opacity: 0, pointerEvents: 'none' }}
        onChange={(e) => ch?.setColor(e.target.value.replace('#', '').toUpperCase()).run()}
      />

      <input
        ref={imgInputRef}
        type="file"
        accept=".png,.jpg,.jpeg,.gif,.webp,.bmp"
        style={{ display: 'none' }}
        onChange={async (e) => {
          const f = e.target.files?.[0];
          e.target.value = '';
          const s = srcRef.current;
          if (!f || !s || !editorRef.current) return;
          try {
            const ext = (f.name.split('.').pop() || 'png').toLowerCase();
            const raw = new Uint8Array(await f.arrayBuffer());
            let width = 320;
            let height = 240;
            try {
              const bmp = await createImageBitmap(f);
              width = Math.min(460, bmp.width);
              height = Math.round((width / bmp.width) * bmp.height);
              bmp.close();
            } catch {
              /* keep defaults */
            }
            const relId = await addImage(s, raw, ext, width, height);
            let src: string | null = null;
            try {
              src = URL.createObjectURL(f);
            } catch {
              src = null;
            }
            editorRef.current.commands.focus('end');
            editorRef.current.commands.insertContent([
              { type: 'image', attrs: { src: src || '', relId, wpx: width, hpx: height, alt: f.name } },
            ]);
          } catch (err) {
            propsRef.current.onStatus(`Image failed: ${(err as Error).message}`);
          }
        }}
      />

      <div className="cd-docx-surface">
        <EditorContent editor={editor} />
      </div>

      <div ref={printHostRef} className="cd-print-host" />
    </div>
  );
});

export default DocxLens;
