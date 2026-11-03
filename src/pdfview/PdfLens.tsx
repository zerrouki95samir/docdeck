import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';
import { pdfjsLib, TextLayer as PdfTextLayer, type TextLayerCtor } from './pdfSetup';
import {
  newBlankPdf,
  applyPageRotation,
  deletePages,
  insertBlankPage,
  mergePdfs,
  flushEdits,
  listFormFields,
  alive,
  u8,
  type Stamp,
  type StampType,
  type FormFieldInfo,
} from './pdfOps';
import type { LensHandle, LensProps } from '../lenses';

type PdfDocLike = {
  numPages: number;
  getPage(n: number): Promise<PdfPageLike>;
  destroy(): Promise<void>;
};
type PdfPageLike = {
  getViewport(o: { scale: number }): PdfViewportLike;
  render(o: { canvasContext: CanvasRenderingContext2D; viewport: unknown }): {
    promise: Promise<void>;
    cancel: () => void;
  };
  getTextContent(): Promise<{ items: Array<{ str?: string }> }>;
};
type PdfViewportLike = {
  width: number;
  height: number;
  convertToPdfPoint(x: number, y: number): number[];
  convertToViewportPoint(x: number, y: number): number[];
};

type Gesture =
  | { kind: 'rect' | 'ellipse' | 'line' | 'highlight'; a: Pt; b: Pt; ui: { x0: number; y0: number; x1: number; y1: number } }
  | { kind: 'freehand'; pts: Pt[]; ui: Pt[] }
  | { kind: 'text'; pos: Pt; ui: Pt };
type Pt = { x: number; y: number };

export const PdfLens = forwardRef<LensHandle, LensProps>(function PdfLens(props, ref) {
  const propsRef = useRef(props);
  propsRef.current = props;

  const wrapRef = useRef<HTMLDivElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const tlRef = useRef<HTMLDivElement>(null);
  const docRef = useRef<PdfDocLike | null>(null);
  const bytesRef = useRef<Uint8Array | null>(null);
  const nameRef = useRef('Untitled.pdf');
  const pathRef = useRef<string | null>(null);
  const dirtyRef = useRef(false);
  const pageRef = useRef(0);
  const vpRef = useRef<PdfViewportLike | null>(null);
  const renderCtl = useRef(0);
  const textCache = useRef<Map<number, Promise<string>>>(new Map());
  const pendingStamps = useRef<Stamp[]>([]);
  const pendingForm = useRef<Record<string, string>>({});
  const fieldTimer = useRef<number | null>(null);
  const thumbToken = useRef(0);
  const b64cache = (u8: Uint8Array): string => {
    let s = '';
    for (let i = 0; i < u8.length; i += 0x4000) {
      const chunk = u8.subarray(i, i + 0x4000);
      s += String.fromCharCode.apply(null, Array.from(chunk) as unknown as number[]);
    }
    return btoa(s);
  };

  const [pages, setPages] = useState(0);
  const [page, setPage] = useState(0);
  const [fitW, setFitW] = useState(true);
  const [scale, setScale] = useState(1);
  const [tool, setTool] = useState<StampType | 'pointer'>('pointer');
  const [color, setColor] = useState('#d0342c');
  const [panel, setPanel] = useState<'none' | 'thumbs' | 'search' | 'fields'>('thumbs');
  const [query, setQuery] = useState('');
  const [results, setResults] = useState<Array<{ page: number; excerpt: string }>>([]);
  const [stampsUi, setStampsUi] = useState<Stamp[]>([]);
  const [fields, setFields] = useState<FormFieldInfo[]>([]);
  const [fieldValues, setFieldValues] = useState<Record<string, string>>({});
  const [gesture, setGesture] = useState<Gesture | null>(null);
  const [thumbTick, setThumbTick] = useState(0);

  /* =================================================================
     pdf.js
     ================================================================= */

  const renderPage = async (idx: number) => {
    const doc = docRef.current;
    const canvas = canvasRef.current;
    const wrap = wrapRef.current;
    if (!doc || !canvas || !wrap) return;
    if (idx < 0 || idx > doc.numPages - 1) return;
    const ctl = ++renderCtl.current;

    const pageObj = await doc.getPage(idx + 1);
    const base = pageObj.getViewport({ scale: 1 });
    const wrapW = wrap.clientWidth || 900;
    const eff = fitW
      ? Math.max(0.05, Math.min(6, (wrapW - 48) / base.width))
      : Math.max(0.2, Math.min(5, scale));
    const cssVp = pageObj.getViewport({ scale: eff });
    const dpr = Math.min(2, window.devicePixelRatio || 1);
    const devVp = pageObj.getViewport({ scale: eff * dpr });

    canvas.width = Math.floor(devVp.width);
    canvas.height = Math.floor(devVp.height);
    canvas.style.width = `${Math.floor(cssVp.width)}px`;
    canvas.style.height = `${Math.floor(cssVp.height)}px`;
    wrap.style.setProperty('--scale-factor', String(eff));

    try {
      await pageObj.render({ canvasContext: canvas.getContext('2d')!, viewport: devVp }).promise;
    } catch {
      return;
    }
    if (ctl !== renderCtl.current) return;
    vpRef.current = cssVp;

    const tl = tlRef.current;
    if (tl) {
      tl.innerHTML = '';
      tl.style.width = `${Math.floor(cssVp.width)}px`;
      tl.style.height = `${Math.floor(cssVp.height)}px`;
      try {
        if (PdfTextLayer) {
          const content = await pageObj.getTextContent();
          const layer = new (PdfTextLayer as unknown as TextLayerCtor)({
            textContentSource: content,
            container: tl,
            viewport: cssVp,
          });
          await layer.render();
        }
      } catch {
        /* text layer optional */
      }
    }

    pageRef.current = idx;
    setPage(idx);
  };

  const activate = async (bytes: Uint8Array) => {
    if (docRef.current) {
      try {
        void (await docRef.current.destroy());
      } catch {
        /* ignore */
      }
      docRef.current = null;
    }
    const doc = (await pdfjsLib.getDocument({ data: alive(bytes) }).promise) as unknown as PdfDocLike;
    docRef.current = doc;
    textCache.current.clear();
    setPages(doc.numPages);
    await renderPage(pageRef.current);
    setThumbTick((t) => t + 1);
  };

  const getPageText = (i: number): Promise<string> => {
    const doc = docRef.current;
    if (!doc) return Promise.resolve('');
    const memo = textCache.current;
    if (memo.has(i)) return memo.get(i)!;
    const fut = (async () => {
      const p = await doc.getPage(i + 1);
      const content = await p.getTextContent();
      return (content.items ?? []).map((it) => it.str ?? '').join(' ');
    })();
    memo.set(i, fut);
    return fut;
  };

  /* =================================================================
     stamps: pending overlay → baked into bytes on save/change
     ================================================================= */

  const addStamp = (st: Stamp) => {
    pendingStamps.current.push(st);
    setStampsUi([...pendingStamps.current]);
    dirtyRef.current = true;
  };

  /* flush pending stamps+forms into the doc bytes (one load+save) */
  const bake = async () => {
    const bytes = bytesRef.current;
    if (!bytes) return;
    const has = pendingStamps.current.length > 0 || Object.keys(pendingForm.current).length > 0;
    if (!has) return;
    const next = await flushEdits(bytes, pendingStamps.current, pendingForm.current);
    bytesRef.current = next;
    pendingStamps.current = [];
    pendingForm.current = {};
    setStampsUi([]);
    dirtyRef.current = true;
    await activate(next);
    await refreshFields();
  };

  const rebake = async (next: Uint8Array) => {
    bytesRef.current = next;
    dirtyRef.current = true;
    await activate(next);
  };

  /* =================================================================
     page operations (eager bake)
     ================================================================= */

  const rotatePageAt = async (i: number) => {
    try {
      await rebake(await applyPageRotation(bytesRef.current!, i, 90));
    } catch (err) {
      propsRef.current.onStatus(`Rotate failed: ${(err as Error).message}`);
    }
  };

  const deletePageAt = async (i: number) => {
    const doc = docRef.current;
    if (!doc) return;
    if (doc.numPages <= 1) {
      propsRef.current.onStatus('Cannot delete the only page');
      return;
    }
    try {
      pendingStamps.current = pendingStamps.current
        .filter((s) => s.page !== i)
        .map((s) => (s.page > i ? { ...s, page: s.page - 1 } : s));
      setStampsUi([...pendingStamps.current]);
      pageRef.current = Math.max(0, i - 1);
      await rebake(await deletePages(bytesRef.current!, [i]));
    } catch (err) {
      propsRef.current.onStatus(`Delete failed: ${(err as Error).message}`);
    }
  };

  const opInsertBlank = async () => {
    try {
      const after = pageRef.current;
      pendingStamps.current = pendingStamps.current.map((s) =>
        s.page > after ? { ...s, page: s.page + 1 } : s
      );
      setStampsUi([...pendingStamps.current]);
      pageRef.current = after + 1;
      await rebake(await insertBlankPage(bytesRef.current!, after, 612, 792));
    } catch (err) {
      propsRef.current.onStatus(`Insert failed: ${(err as Error).message}`);
    }
  };

  const opRotate = async () => {
    await rotatePageAt(pageRef.current);
  };

  const opDelete = async () => {
    await deletePageAt(pageRef.current);
  };

  const opMerge = async () => {
    const picked = await window.cellDesk!.openDialog('pdf');
    if (!picked || !picked.length) return;
    const p = picked[0];
    propsRef.current.onBusy('Merging…');
    try {
      const other = await window.cellDesk!.readFile(p);
      const merged = await mergePdfs(bytesRef.current!, [u8(new Uint8Array(other))]);
      await rebake(merged);
      propsRef.current.onStatus(`Merged ${base(p)}`);
    } catch (err) {
      propsRef.current.onStatus(`Merge failed: ${(err as Error).message}`);
    }
    propsRef.current.onBusy(null);
  };

  /* =================================================================
     gestures
     ================================================================= */

  const vpPt = async (e: React.MouseEvent) => {
    const vp = vpRef.current;
    const canvas = canvasRef.current;
    if (!vp || !canvas) return null;
    const r = canvas.getBoundingClientRect();
    const x = e.clientX - r.left;
    const y = e.clientY - r.top;
    const p = vp.convertToPdfPoint(x, y);
    return { pdf: { x: p[0], y: p[1] }, ui: { x, y } };
  };

  const onCanvasDown = async (e: React.MouseEvent) => {
    if (tool === 'pointer') return;
    e.preventDefault();
    const hit = await vpPt(e);
    if (!hit) return;
    if (tool === 'text') {
      setGesture({ kind: 'text', pos: hit.pdf, ui: { x: hit.ui.x, y: hit.ui.y } });
      return;
    }
    if (tool === 'freehand') setGesture({ kind: 'freehand', pts: [hit.pdf], ui: [hit.ui] });
    else
      setGesture({
        kind: tool,
        a: hit.pdf,
        b: hit.pdf,
        ui: { x0: hit.ui.x, y0: hit.ui.y, x1: hit.ui.x, y1: hit.ui.y },
      });
  };

  const onCanvasMove = async (e: React.MouseEvent) => {
    if (!gesture || gesture.kind === 'text') return;
    const hit = await vpPt(e);
    if (!hit) return;
    if (gesture.kind === 'freehand')
      setGesture({ kind: 'freehand', pts: [...gesture.pts, hit.pdf], ui: [...gesture.ui, hit.ui] });
    else setGesture({ ...gesture, b: hit.pdf, ui: { ...gesture.ui, x1: hit.ui.x, y1: hit.ui.y } });
  };

  const onCanvasUp = () => {
    if (!gesture || gesture.kind === 'text') return;
    const g = gesture;
    setGesture(null);
    const common = { page: pageRef.current, color };
    if (g.kind === 'freehand') {
      if (g.pts.length >= 2) addStamp({ ...common, type: 'freehand', pts: g.pts, width: 2.5 });
      return;
    }
    const a = g.a;
    const b = g.b;
    if (Math.hypot(b.x - a.x, b.y - a.y) < 3) return;
    addStamp({ ...common, type: g.kind, pts: [a, b], width: 2 });
  };

  const commitText = (text: string) => {
    if (gesture?.kind !== 'text') return;
    const g = gesture;
    setGesture(null);
    if (text.trim())
      addStamp({
        page: pageRef.current,
        type: 'text',
        pts: [g.pos],
        text: text.trim(),
        fontSize: 14,
        color,
      });
  };

  /* =================================================================
     search
     ================================================================= */

  const runSearch = async () => {
    const q = query.trim().toLowerCase();
    if (!q) return;
    propsRef.current.onBusy('Searching…');
    const out: Array<{ page: number; excerpt: string }> = [];
    for (let i = 0; i < pages; i++) {
      const norm = (await getPageText(i)).toLowerCase().replace(/\s+/g, ' ');
      let at = norm.indexOf(q);
      while (at !== -1 && out.length < 80) {
        const from = Math.max(0, at - 28);
        out.push({ page: i, excerpt: `…${norm.slice(from, at + q.length + 40).replace(/\s+/g, ' ')}…` });
        at = norm.indexOf(q, at + q.length);
      }
    }
    setResults(out);
    propsRef.current.onBusy(null);
    propsRef.current.onStatus(
      out.length ? `${out.length} hit${out.length === 1 ? '' : 's'} for “${query}”` : `No hits for “${query}”`
    );
    if (out.length) await renderPage(out[0].page);
  };

  /* =================================================================
     form fields — baked in live (values live in the document)
     ================================================================= */

  const refreshFields = async () => {
    if (!bytesRef.current) return;
    try {
      const info = await listFormFields(bytesRef.current);
      setFields(info);
      const vals: Record<string, string> = {};
      info.forEach((f) => {
        vals[f.name] = f.value != null ? String(f.value) : f.type === 'checkbox' ? 'false' : '';
      });
      setFieldValues(vals);
    } catch {
      setFields([]);
    }
  };

  const onFieldChange = (name: string, value: string) => {
    setFieldValues((prev) => ({ ...prev, [name]: value }));
    const form = { ...pendingForm.current, [name]: value };
    pendingForm.current = form;
    dirtyRef.current = true;
    debouncedBakeForm();
  };

  const debouncedBakeForm = () => {
    if (fieldTimer.current) window.clearTimeout(fieldTimer.current);
    fieldTimer.current = window.setTimeout(() => {
      void bake();
    }, 500);
  };

  useEffect(() => {
    if (fieldTimer.current) window.clearTimeout(fieldTimer.current);
    if (!Object.keys(pendingForm.current).length) return;
    fieldTimer.current = window.setTimeout(() => {
      void bake();
    }, 500);
    return () => {
      if (fieldTimer.current) window.clearTimeout(fieldTimer.current);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fieldValues]);

  /* =================================================================
     thumbnails
     ================================================================= */

  useEffect(() => {
    if (panel !== 'thumbs' || !docRef.current || !pages) return;
    const token = ++thumbToken.current;
    let cancelled = false;
    (async () => {
      const host = document.getElementById('pdf-thumb-panel');
      if (!host) return;
      host.querySelectorAll('[data-thumb]').forEach((el) => {
        (el as HTMLElement).dataset.rendered = '';
        while (el.firstChild) el.removeChild(el.firstChild);
      });
      const doc = docRef.current!;
      for (let i = 0; i < doc.numPages; i++) {
        if (cancelled || token !== thumbToken.current) return;
        const slot = host.querySelector<HTMLElement>(`[data-thumb="${i}"]`);
        if (!slot) continue;
        try {
          const p = await doc.getPage(i + 1);
          const base = p.getViewport({ scale: 1 });
          const dpr = Math.min(2, window.devicePixelRatio || 1);
          const vp = p.getViewport({ scale: (132 / base.width) * dpr });
          const c = document.createElement('canvas');
          c.width = Math.floor(vp.width);
          c.height = Math.floor(vp.height);
          c.style.width = '132px';
          c.style.height = 'auto';
          await p.render({ canvasContext: c.getContext('2d')!, viewport: vp }).promise;
          slot.appendChild(c);
          await new Promise((r) => setTimeout(r, 6));
        } catch {
          /* skip */
        }
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [thumbTick, panel, pages]);

  useEffect(() => {
    if (docRef.current) void renderPage(pageRef.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fitW, scale]);

  useEffect(() => {
    const wrap = wrapRef.current;
    if (!wrap) return;
    const ro = new ResizeObserver(() => {
      if (docRef.current && fitW) void renderPage(pageRef.current);
    });
    ro.observe(wrap);
    return () => ro.disconnect();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fitW]);

  /* =================================================================
     imperative API
     ================================================================= */

  useImperativeHandle<LensHandle, LensHandle>(ref, () => ({
    kind: 'pdf',
    async open(name: string, bytes: Uint8Array, path: string | null) {
      pendingStamps.current = [];
      pendingForm.current = {};
      setStampsUi([]);
      nameRef.current = name;
      pathRef.current = path;
      dirtyRef.current = false;
      pageRef.current = 0;
      setPanel('thumbs');
      try {
        bytesRef.current = u8(bytes);
        await activate(bytesRef.current);
        await refreshFields();
      } catch (err) {
        const msg = (err as Error).message || 'unknown error';
        propsRef.current.onStatus(`Could not read PDF: ${msg}`);
        throw new Error(msg);
      }
      propsRef.current.onTitle(name);
    },
    async save(asNew: boolean) {
      if (!bytesRef.current) return false;
      let target = pathRef.current;
      if (!target || asNew) {
        target = await window.cellDesk!.saveDialog({
          defaultPath: pathRef.current || nameRef.current,
          filters: [{ name: 'PDF Document', extensions: ['pdf'] }],
        });
        if (!target) return false;
      }
      propsRef.current.onBusy('Saving…');
      try {
        await bake();
        if (fieldTimer.current) window.clearTimeout(fieldTimer.current);
        const finalBytes = bytesRef.current!;
        await window.cellDesk!.writeFile(target, finalBytes);
        pathRef.current = target;
        nameRef.current = base(target);
        dirtyRef.current = false;
        propsRef.current.onTitle(nameRef.current);
        propsRef.current.onBusy(null);
        propsRef.current.onStatus(`Saved ${nameRef.current}`);
        return true;
      } catch (err) {
        propsRef.current.onStatus(`Save failed: ${(err as Error).message}`);
        propsRef.current.onBusy(null);
        return false;
      }
    },
    async makeNew() {
      pendingStamps.current = [];
      pendingForm.current = {};
      setStampsUi([]);
      bytesRef.current = await newBlankPdf();
      nameRef.current = 'Untitled.pdf';
      pathRef.current = null;
      dirtyRef.current = false;
      pageRef.current = 0;
      setFields([]);
      setFieldValues({});
      setPanel('thumbs');
      try {
        await activate(bytesRef.current);
      } catch (err) {
        propsRef.current.onStatus(`Could not create PDF: ${(err as Error).message}`);
        throw err;
      }
      propsRef.current.onTitle(nameRef.current);
    },
    async exportAs() {
      propsRef.current.onStatus('The document is already a PDF');
    },
    async print() {
      propsRef.current.onStatus('PDF printing lands later');
    },
    isDirty() {
      return dirtyRef.current;
    },
    fileName() {
      return nameRef.current;
    },
  }));

  useEffect(() => {
    if (props.active === false) return; // only the active tab publishes dev hooks
    (window as unknown as Record<string, unknown>).__cellDeskDevPdf = {
      open: async (name: string, b64: string) => {
        const bin = atob(b64);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        const api = (ref as React.MutableRefObject<LensHandle | null>)?.current;
        await api?.open(name, arr, null);
        return { pages: docRef.current?.numPages ?? 0, title: nameRef.current };
      },
      save: async () => {
        await bake();
        await new Promise((r) => setTimeout(r, 30));
        return b64cache(bytesRef.current!);
      },
      stamp: (spec: Partial<Stamp> & { type: StampType; pts: Pt[] }) => {
        addStamp({ ...(spec as Stamp), page: spec.page ?? pageRef.current });
        return { pending: pendingStamps.current.length };
      },
      rotate: async (i: number) => {
        await rotatePageAt(i);
        return { pages: docRef.current?.numPages ?? 0 };
      },
      del: async (i: number) => {
        await deletePageAt(i);
        return { pages: docRef.current?.numPages ?? 0 };
      },
      insert: async (i: number) => {
        const nxt = await insertBlankPage(bytesRef.current!, i, 612, 792);
        await rebake(nxt);
        return { pages: docRef.current?.numPages ?? 0 };
      },
      search: async (q: string) => {
        const lower = q.toLowerCase();
        const out: Array<{ page: number; excerpt: string }> = [];
        for (let i = 0; i < (docRef.current?.numPages ?? 0); i++) {
          const norm = (await getPageText(i)).toLowerCase().replace(/\s+/g, ' ');
          const at = norm.indexOf(lower);
          if (at !== -1) out.push({ page: i, excerpt: norm.slice(Math.max(0, at - 20), at + lower.length + 30) });
        }
        setResults(out);
        return out;
      },
      fields: async () => await listFormFields(bytesRef.current!),
      setField: async (name: string, value: string) => {
        const nxt = await flushEdits(bytesRef.current!, [], { [name]: value });
        await rebake(nxt);
        return true;
      },
      view: () => ({
        page: pageRef.current + 1,
        pages: docRef.current?.numPages ?? 0,
        dirty: dirtyRef.current,
      }),
      goto: async (idx: number) => {
        await renderPage(idx - 1);
        return true;
      },
    };
    return () => {
      delete (window as unknown as Record<string, unknown>).__cellDeskDevPdf;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.active]);

  /* =================================================================
     render
     ================================================================= */

  return (
    <div className="lens-pdf">
      <div className="pdf-toolbar">
        <div className="pdf-tb-group">
          <button onClick={() => void renderPage(pageRef.current - 1)} disabled={page <= 0}>
            ‹
          </button>
          <span className="pdf-pageno">
            {pages ? page + 1 : 0}/{pages}
          </span>
          <button onClick={() => void renderPage(pageRef.current + 1)} disabled={page >= pages - 1}>
            ›
          </button>
        </div>
        <div className="pdf-tb-group">
          <select
            value={fitW ? 'fit' : String(scale)}
            onChange={(e) => {
              const v = e.target.value;
              if (v === 'fit') setFitW(true);
              else {
                setFitW(false);
                setScale(Number(v));
              }
            }}
          >
            <option value="fit">fit width</option>
            <option value="0.5">50%</option>
            <option value="0.75">75%</option>
            <option value="1">100%</option>
            <option value="1.5">150%</option>
            <option value="2">200%</option>
            <option value="3">300%</option>
          </select>
        </div>
        <div className="pdf-tb-group">
          <button className={tool === 'pointer' ? 'tool-on' : ''} onClick={() => setTool('pointer')} title="Select text">
            ▤
          </button>
          <button className={tool === 'highlight' ? 'tool-on' : ''} onClick={() => setTool('highlight')} title="Highlight">
            ▨
          </button>
          <button className={tool === 'rect' ? 'tool-on' : ''} onClick={() => setTool('rect')} title="Rectangle">
            ▭
          </button>
          <button className={tool === 'ellipse' ? 'tool-on' : ''} onClick={() => setTool('ellipse')} title="Ellipse">
            ◯
          </button>
          <button className={tool === 'line' ? 'tool-on' : ''} onClick={() => setTool('line')} title="Line">
            ／
          </button>
          <button className={tool === 'freehand' ? 'tool-on' : ''} onClick={() => setTool('freehand')} title="Scribble">
            ✎
          </button>
          <button className={tool === 'text' ? 'tool-on' : ''} onClick={() => setTool('text')} title="Type text">
            𝐓
          </button>
          <input
            className="pdf-color"
            type="color"
            value={color}
            onChange={(e) => setColor(e.target.value)}
            title="Stamp color"
          />
        </div>
        <div className="pdf-tb-group">
          <button onClick={() => void opRotate()} title="Rotate page 90°">↻</button>
          <button onClick={() => void opInsertBlank()} title="Insert blank page after current">＋</button>
          <button onClick={() => void opDelete()} disabled={pages <= 1} title="Delete page">
            🗑
          </button>
          <button onClick={() => void opMerge()} title="Merge another PDF…">⇥</button>
        </div>
        <div className="pdf-tb-group pdf-search-group">
          <input
            value={query}
            placeholder="Search text…"
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === 'Enter') void runSearch();
            }}
          />
        </div>
        <div className="pdf-tb-group">
          {(['thumbs', 'search', 'fields', 'none'] as const).map((p) => (
            <button key={p} className={panel === p ? 'tool-on' : ''} onClick={() => setPanel(p)}>
              {p === 'thumbs' ? 'Pages' : p === 'search' ? 'Hits' : p === 'fields' ? 'Forms' : '×'}
            </button>
          ))}
        </div>
      </div>

      <div className={`pdf-body ${tool !== 'pointer' ? 'tool-active' : ''}`}>
        {panel !== 'none' && (
          <div className="pdf-side">
            {panel === 'thumbs' && (
              <div id="pdf-thumb-panel" className="pdf-thumbs">
                {new Array(pages).fill(0).map((_, i) => (
                  <div
                    key={i}
                    data-thumb={`${i}`}
                    className={`pdf-thumb ${i === page ? 'thumb-on' : ''}`}
                    onClick={() => void renderPage(i)}
                  >
                    <div className="pdf-thumb-tools">
                      <button
                        title="Rotate page"
                        onClick={async (e) => {
                          e.stopPropagation();
                          await rotatePageAt(i);
                        }}
                      >
                        ↻
                      </button>
                      <button
                        disabled={(docRef.current?.numPages ?? pages) <= 1}
                        title="Delete page"
                        onClick={async (e) => {
                          e.stopPropagation();
                          await deletePageAt(i);
                        }}
                      >
                        ✕
                      </button>
                    </div>
                    <span className="pdf-thumb-no">{i + 1}</span>
                  </div>
                ))}
              </div>
            )}

            {panel === 'search' && (
              <div className="pdf-search-results">
                {results.length === 0 && (
                  <div className="pdf-side-empty">No hits yet — search above, press Enter.</div>
                )}
                {results.map((r, i) => (
                  <div key={i} className="pdf-hit" onClick={() => void renderPage(r.page)}>
                    <span className="pdf-hit-page">p{r.page + 1}</span>
                    <span className="pdf-hit-text">{r.excerpt}</span>
                  </div>
                ))}
              </div>
            )}

            {panel === 'fields' && (
              <div className="pdf-fields">
                {fields.length === 0 && <div className="pdf-side-empty">No fillable form fields.</div>}
                {fields.map((f) => (
                  <div key={f.name} className="pdf-field-row" title={f.name}>
                    <label>
                      <span className="pdf-field-name">{f.name}</span>
                      {f.type === 'text' && (
                        <input
                          value={fieldValues[f.name] ?? ''}
                          onChange={(e) => onFieldChange(f.name, e.target.value)}
                        />
                      )}
                      {f.type === 'checkbox' && (
                        <input
                          type="checkbox"
                          checked={fieldValues[f.name] === 'true'}
                          onChange={(e) => onFieldChange(f.name, e.target.checked ? 'true' : 'false')}
                        />
                      )}
                      {(f.type === 'dropdown' || f.type === 'radio') && (
                        <select value={fieldValues[f.name] ?? ''} onChange={(e) => onFieldChange(f.name, e.target.value)}>
                          <option value="">—</option>
                          {(f.options || []).map((o) => (
                            <option key={o} value={o}>
                              {o}
                            </option>
                          ))}
                        </select>
                      )}
                      {f.type === 'other' && <input disabled placeholder="(unsupported)" />}
                    </label>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        <div
          className="pdf-canvas-wrap"
          ref={wrapRef}
          onMouseDown={(e) => void onCanvasDown(e)}
          onMouseMove={(e) => void onCanvasMove(e)}
          onMouseUp={() => void onCanvasUp()}
        >
          <div className="pdf-canvas-outer">
            <div className="pdf-page-slot">
              <canvas ref={canvasRef} className="pdf-canvas" />
              <div ref={tlRef} className="textLayer" />
              <div className="pdf-stamp-layer">
                {stampsUi
                  .filter((s) => s.page === page)
                  .map((s, i) => (
                    <StampView
                      key={i}
                      st={s}
                      vp={vpRef.current}
                                          />
                  ))}
              </div>
              {gesture && gesture.kind !== 'text' && vpRef.current && (
                <StampView st={{ page, type: gesture.kind, pts: gesturePts(gesture), color } as Stamp} vp={vpRef.current} live />
              )}
              {gesture?.kind === 'text' && (
                <textarea
                  className="pdf-floating-input"
                  autoFocus
                  rows={2}
                  placeholder="Type… Enter to place"
                  onKeyDown={(e) => {
                    const ta = e.target as HTMLTextAreaElement;
                    if (e.key === 'Enter' && !e.shiftKey) {
                      e.preventDefault();
                      commitText(ta.value);
                    } else if (e.key === 'Escape') setGesture(null);
                  }}
                  onBlur={(e) => commitText((e.target as HTMLTextAreaElement).value)}
                  style={{ left: gesture.ui.x, top: gesture.ui.y }}
                />
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  function gesturePts(g: Gesture): Pt[] {
    if (g.kind === 'freehand') return g.pts;
    if (g.kind === 'text') return [g.pos];
    return [g.a, g.b];
  }
});

function StampView({ st, vp, live }: { st: Stamp; vp: unknown; live?: boolean }) {
  const viewport = vp as PdfViewportLike | null;
  if (!viewport?.convertToViewportPoint) return null;
  const pts = st.pts.map((p) => viewport.convertToViewportPoint(p.x, p.y));
  if (!pts.length) return null;
  const common: React.CSSProperties = { position: 'absolute', pointerEvents: 'none' };
  switch (st.type) {
    case 'highlight':
    case 'rect': {
      const a = pts[0];
      const b = pts[pts.length - 1];
      return (
        <div
          style={{
            ...common,
            left: Math.min(a[0], b[0]),
            top: Math.min(a[1], b[1]),
            width: Math.abs(b[0] - a[0]),
            height: Math.abs(b[1] - a[1]),
            background:
              st.type === 'highlight'
                ? hexToRgba(st.color ?? '#ff9f00', 0.34)
                : 'transparent',
            border: st.type === 'rect' ? `${st.width ?? 2}px solid ${st.color ?? '#1e4db7'}` : 'none',
            mixBlendMode: 'multiply',
            opacity: live ? 0.9 : 1,
          }}
        />
      );
    }
    case 'ellipse': {
      const a = pts[0];
      const b = pts[pts.length - 1];
      return (
        <div
          style={{
            ...common,
            left: Math.min(a[0], b[0]),
            top: Math.min(a[1], b[1]),
            width: Math.abs(b[0] - a[0]),
            height: Math.abs(b[1] - a[1]),
            border: `${st.width ?? 2}px solid ${st.color ?? '#1e4db7'}`,
            borderRadius: '50%',
          }}
        />
      );
    }
    case 'line':
    case 'freehand': {
      const d = pts.map((p, i) => `${i ? 'L' : 'M'}${p[0]},${p[1]}`).join(' ');
      return (
        <svg style={{ ...common, left: 0, top: 0, width: '100%', height: '100%' }}>
          <path d={d} fill="none" stroke={st.color ?? '#1e4db7'} strokeWidth={st.width ?? 2} strokeLinecap="round" />
        </svg>
      );
    }
    case 'text':
      return (
        <div
          style={{
            ...common,
            left: pts[0][0],
            top: pts[0][1] - (st.fontSize ?? 14) * 1.05,
            fontSize: st.fontSize ?? 14,
            color: st.color ?? '#1e4db7',
            fontFamily: 'Helvetica, Arial, sans-serif',
            whiteSpace: 'pre',
            transform: 'translateY(-15%)',
          }}
        >
          {st.text ?? ''}
        </div>
      );
    default:
      return null;
  }
}

function hexToRgba(hex: string, a: number): string {
  const h = hex.replace('#', '');
  const v = h.length === 3 ? h.split('').map((c) => c + c).join('') : h;
  const n = parseInt(v, 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${a})`;
}

function base(p: string): string {
  return p.split('/').pop() || p;
}

export default PdfLens;
