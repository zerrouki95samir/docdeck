import { useCallback, useEffect, useRef, useState } from 'react';
import { SheetLens } from './sheetview/SheetLens';
import { PdfLens } from './pdfview/PdfLens';
import { DocxLens } from './docx/DocxLens';
import Home from './home/Home';
import type { DocKind, LensHandle, LensProps } from './lenses';

type PendingLoad =
  | { op: 'open'; kind: DocKind; name: string; bytes: Uint8Array; path: string | null }
  | { op: 'new'; kind: DocKind };

interface DocTab {
  id: string;
  kind: DocKind;
  name: string;
  path: string | null;
  dirty: boolean;
}

function extKind(path: string): DocKind | null {
  const m = path.toLowerCase().match(/\.([a-z0-9]+)$/);
  const ext = m?.[1] || '';
  if (['xlsx', 'xlsm', 'xls', 'csv', 'tsv', 'uni'].includes(ext)) return 'sheet';
  if (ext === 'pdf') return 'pdf';
  if (ext === 'docx') return 'docx';
  return null;
}

const APP_NAME = 'DocDeck';
const KIND_LABEL: Record<DocKind, string> = { sheet: 'Spreadsheet', docx: 'Word', pdf: 'PDF' };
const DEFAULT_EXT: Record<DocKind, string> = { sheet: 'xlsx', docx: 'docx', pdf: 'pdf' };
const LENS_KINDS = ['sheet', 'docx', 'pdf'] as const;

/* ---------- chrome icons ---------- */

function ChromeIcon({ k }: { k: DocKind | 'home' }) {
  const size = 15;
  const s = { width: size, height: size } as const;
  if (k === 'home')
    return (
      <svg {...s} viewBox="0 0 24 24" fill="none" aria-hidden>
        <path d="M4 11.2 12 4.5l8 6.7v7.3a1.5 1.5 0 0 1-1.5 1.5h-13A1.5 1.5 0 0 1 4 18.5v-7.3Z" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" />
        <path d="M9.5 20v-6h5v6" stroke="currentColor" strokeWidth="1.7" strokeLinejoin="round" />
      </svg>
    );
  if (k === 'sheet')
    return (
      <svg {...s} viewBox="0 0 24 24" fill="none" aria-hidden>
        <rect x="3.5" y="3.5" width="17" height="17" rx="2.5" stroke="currentColor" strokeWidth="1.7" />
        <path d="M3.5 9.5h17M3.5 15h17M10 3.5v17" stroke="currentColor" strokeWidth="1.7" />
      </svg>
    );
  if (k === 'docx')
    return (
      <svg {...s} viewBox="0 0 24 24" fill="none" aria-hidden>
        <path d="M6 2.75h7.5L19 8.25v11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-14.5a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.7" />
        <path d="M13 3v5.5h6M7.5 13h7M7.5 16.5h4.5" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
      </svg>
    );
  return (
    <svg {...s} viewBox="0 0 24 24" fill="none" aria-hidden>
      <path d="M6 2.75h7.5L19 8.25v11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2v-14.5a2 2 0 0 1 2-2Z" stroke="currentColor" strokeWidth="1.7" />
      <path d="M13 3v5.5h6M7.6 17V11.4h1.7a1.9 1.9 0 0 1 0 3.8H7.6m8.2 1.8v-5.6h2" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
    </svg>
  );
}

export function App() {
  const [tabs, setTabs] = useState<DocTab[]>([]);
  const [activeId, setActiveId] = useState<string | null>(null);
  const tabsRef = useRef(tabs);
  tabsRef.current = tabs;
  const activeIdRef = useRef(activeId);
  activeIdRef.current = activeId;

  const lensRefs = useRef(new Map<string, LensHandle>());
  const pendingOpsRef = useRef(new Map<string, PendingLoad>());
  const tabSeqRef = useRef(0);

  const [recents, setRecents] = useState<string[]>([]);
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [status, setStatus] = useState('');
  const [, setTick] = useState(0);
  const bump = useCallback(() => setTick((t) => t + 1), []);

  const lensAt = useCallback((id: string) => lensRefs.current.get(id) ?? null, []);
  const activeTab = tabs.find((t) => t.id === activeId) ?? null;

  const refreshTitle = useCallback(() => {
    const id = activeIdRef.current;
    const t = tabsRef.current.find((x) => x.id === id);
    if (!t) {
      window.cellDesk?.setTitle(APP_NAME, false);
      document.title = APP_NAME;
      return;
    }
    const lens = lensAt(t.id);
    const dirty = Boolean(lens?.isDirty());
    const title = `${dirty ? '• ' : ''}${t.name} — ${APP_NAME}`;
    document.title = title;
    window.cellDesk?.setTitle(t.name, dirty);
  }, [lensAt]);

  useEffect(() => {
    refreshTitle();
  }, [tabs, activeId, refreshTitle]);

  /* ---------- pending routing: mount first, load right after ---------- */

  useEffect(() => {
    for (const [id, p] of pendingOpsRef.current) {
      const lens = lensRefs.current.get(id);
      if (!lens) continue;
      pendingOpsRef.current.delete(id);
      void (async () => {
        try {
          if (p.op === 'open') {
            await lens.open(p.name, p.bytes, p.path);
            if (p.path) void window.cellDesk?.recents.add(p.path).then(() => bump());
            setStatus(`Opened ${p.name}`);
          } else {
            await lens.makeNew();
            setStatus(`New ${KIND_LABEL[p.kind]} created`);
          }
        } catch (err) {
          setStatus(`Could not load: ${(err as Error).message}`);
        }
        setTabs((ts) => ts.map((x) => (x.id === id ? { ...x, dirty: Boolean(lensRefs.current.get(id)?.isDirty()) } : x)));
        bump();
      })();
    }
  });

  /* ---------- tab lifecycle ---------- */

  const createTab = useCallback((kind: DocKind, p: PendingLoad, path: string | null) => {
    const id = `t${++tabSeqRef.current}`;
    const name = p.op === 'open' ? p.name : `Untitled.${DEFAULT_EXT[kind]}`;
    setTabs((ts) => [...ts, { id, kind, name, path, dirty: false }]);
    pendingOpsRef.current.set(id, p);
    activeIdRef.current = id;
    setActiveId(id);
    bump();
    return id;
  }, []);

  const closeTab = useCallback(
    async (id: string) => {
      const tab = tabsRef.current.find((t) => t.id === id);
      const lens = lensRefs.current.get(id);
      if (tab && lens?.isDirty()) {
        const res = await window.cellDesk?.leaveDialog();
        if (res === 2) return;
        if (res === 0 && !(await lens.save(false))) return;
      }
      lensRefs.current.delete(id);
      pendingOpsRef.current.delete(id);
      const rest = tabsRef.current.filter((t) => t.id !== id);
      setTabs(rest);
      if (activeIdRef.current === id) {
        const next = rest.length ? rest[rest.length - 1].id : null;
        activeIdRef.current = next;
        setActiveId(next);
      }
      bump();
    },
    []
  );

  const switchTab = useCallback(
    (id: string) => {
      if (activeIdRef.current === id) return;
      activeIdRef.current = id;
      setActiveId(id);
      bump();
    },
    []
  );

  const goHome = useCallback(() => {
    activeIdRef.current = null;
    setActiveId(null);
    bump();
  }, []);

  const switchKind = useCallback(
    (k: DocKind) => {
      const t = tabsRef.current.find((x) => x.kind === k);
      if (t) switchTab(t.id);
      else createTab(k, { op: 'new', kind: k }, null);
    },
    [switchTab, createTab]
  );

  /* ---------- open / create ---------- */

  const doOpenPath = useCallback(
    async (path: string) => {
      const k = extKind(path);
      if (!k) {
        setStatus(`Unsupported file type: ${path.split('/').pop() || path}`);
        return;
      }
      const existing = tabsRef.current.find((t) => t.path === path);
      if (existing) {
        switchTab(existing.id);
        return;
      }
      setBusy('Opening…');
      try {
        const bytes = await window.cellDesk!.readFile(path);
        const name = path.split('/').pop() || path;
        createTab(k, { op: 'open', kind: k, name, bytes: new Uint8Array(bytes), path }, path);
      } catch (err) {
        setStatus(`Could not open file: ${(err as Error).message}`);
      } finally {
        setBusy(null);
      }
    },
    [createTab, switchTab]
  );

  const openViaDialog = useCallback(async () => {
    const paths = await window.cellDesk?.openDialog();
    if (Array.isArray(paths)) for (const p of paths) await doOpenPath(p);
    else if (paths) await doOpenPath(paths);
  }, [doOpenPath]);

  /* ---------- wiring ---------- */

  const wiredRef = useRef(false);
  useEffect(() => {
    if (wiredRef.current) return;
    wiredRef.current = true;

    window.cellDesk!.queryDirty(() => tabsRef.current.some((t) => Boolean(lensAt(t.id)?.isDirty())));

    const offs: Array<() => void> = [];
    offs.push(
      window.cellDesk!.onMenu(async (action) => {
        try {
          switch (action) {
            case 'new-sheet':
              createTab('sheet', { op: 'new', kind: 'sheet' }, null);
              break;
            case 'new-pdf':
              createTab('pdf', { op: 'new', kind: 'pdf' }, null);
              break;
            case 'new-docx':
              createTab('docx', { op: 'new', kind: 'docx' }, null);
              break;
            case 'open':
              await openViaDialog();
              break;
            case 'go-home':
              goHome();
              break;
            case 'save': {
              const id = activeIdRef.current;
              if (id) await lensAt(id)?.save(false);
              setTabs((ts) => ts.map((t) => (t.id === id ? { ...t, dirty: Boolean(lensAt(t.id)?.isDirty()) } : t)));
              break;
            }
            case 'save-as': {
              const id = activeIdRef.current;
              if (id) {
                await lensAt(id)?.save(true);
                setTabs((ts) =>
                  ts.map((t) =>
                    t.id === id && lensAt(t.id) ? { ...t, name: lensAt(t.id)!.fileName(), dirty: Boolean(lensAt(t.id)?.isDirty()) } : t
                  )
                );
              }
              break;
            }
            case 'export-csv': {
              const id = activeIdRef.current;
              if (id) await lensAt(id)?.exportAs('csv');
              break;
            }
            case 'save-uni': {
              const id = activeIdRef.current;
              if (id) await lensAt(id)?.exportAs('uni');
              break;
            }
            case 'print': {
              const id = activeIdRef.current;
              if (id) await lensAt(id)?.print();
              break;
            }
            case 'save-then-exit': {
              for (const t of tabsRef.current) {
                const lens = lensAt(t.id);
                if (lens?.isDirty()) {
                  const ok = (await lens.save(false)) === true;
                  if (!ok) return;
                }
              }
              await window.cellDesk!.confirmClose();
              break;
            }
          }
        } catch (err) {
          setStatus((err as Error).message);
        }
        bump();
      })
    );

    offs.push(
      window.cellDesk!.onOpenPath((p) => void doOpenPath(p).then(bump)),
      window.cellDesk!.onRecentsChanged(() => {
        void window.cellDesk!.recents.list().then(setRecents);
      })
    );

    void window.cellDesk!.recents.list().then(setRecents).catch(() => setRecents([]));

    document.title = APP_NAME;
    window.cellDesk?.setTitle(APP_NAME, false);

    return () => {
      offs.forEach((off) => off?.());
      wiredRef.current = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!status) return;
    const t = setTimeout(() => {
      setStatus('');
    }, 4000);
    return () => clearTimeout(t);
  }, [status]);

  /* ---------- drag & drop ---------- */

  useEffect(() => {
    const dragDepth = { n: 0 };
    const hasFiles = (e: DragEvent) => Array.from(e.dataTransfer?.types || []).includes('Files');
    const enter = (e: DragEvent) => {
      if (!hasFiles(e)) return;
      e.preventDefault();
      dragDepth.n += 1;
      setDragging(true);
    };
    const over = (e: DragEvent) => {
      if (hasFiles(e)) e.preventDefault();
    };
    const leave = () => {
      dragDepth.n = Math.max(0, dragDepth.n - 1);
      if (dragDepth.n === 0) setDragging(false);
    };
    const drop = async (e: Event) => {
      const de = e as DragEvent;
      de.preventDefault();
      dragDepth.n = 0;
      setDragging(false);
      const files = Array.from(de.dataTransfer?.files || []);
      if (!files.length) return;
      let openedCount = 0;
      for (const file of files.slice(0, 8)) {
        const k = extKind(file.name);
        if (!k) continue;
        const bytes = new Uint8Array(await file.arrayBuffer());
        createTab(k, { op: 'open', kind: k, name: file.name, bytes, path: null }, null);
        openedCount += 1;
      }
      if (openedCount > 0) setBusy(null);
      else setStatus('Unsupported file type — try .xlsx, .docx or .pdf');
      bump();
    };
    window.addEventListener('dragenter', enter as EventListener);
    window.addEventListener('dragover', over as EventListener);
    window.addEventListener('dragleave', leave);
    window.addEventListener('drop', drop);
    return () => {
      window.removeEventListener('dragenter', enter);
      window.removeEventListener('dragover', over);
      window.removeEventListener('dragleave', leave);
      window.removeEventListener('drop', drop as EventListener);
    };
  }, [createTab, bump]);

  /* ---------- dev hooks ---------- */

  useEffect(() => {
    (window as unknown as Record<string, unknown>).__cellDeskDevApp = {
      openB64: async (name: string, b64: string, fileKind: DocKind) => {
        const bin = atob(b64);
        const arr = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
        createTab(fileKind, { op: 'open', kind: fileKind, name, bytes: arr, path: null }, null);
        return fileKind;
      },
    };
    return () => {
      delete (window as unknown as Record<string, unknown>).__cellDeskDevApp;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [createTab]);

  /* =================================================================
     render
     ================================================================= */

  const mkLensProps = useCallback(
    (t: DocTab): LensProps => ({
      active: activeId === t.id,
      onDirty: () => {
        const lens = lensRefs.current.get(t.id);
        setTabs((ts) => ts.map((x) => (x.id === t.id ? { ...x, dirty: Boolean(lens?.isDirty()) } : x)));
      },
      onTitle: (name: string) => {
        setTabs((ts) => ts.map((x) => (x.id === t.id ? { ...x, name } : x)));
      },
      onStatus: setStatus,
      onBusy: setBusy,
    }),
    [activeId]
  );

  const resume = tabs.map((t) => ({ kind: t.kind, name: t.name, dirty: t.dirty }));

  return (
    <div className="celldeck-root">
      <div className="lens-chrome">
        <button className="chrome-home-btn" onClick={goHome} title="Workspace Home">
          <ChromeIcon k="home" />
          <span>Home</span>
        </button>
        <span className="chrome-divider" />
        <div className="chrome-tabs">
          {tabs.map((t, i) => (
            <div
              key={t.id}
              className={`chrome-tab ${t.id === activeId ? 'tab-on' : ''}`}
              onClick={() => switchTab(t.id)}
              title={!t.dirty ? t.name : `• ${t.name} — unsaved changes`}
            >
              <span className={`chrome-kind_tag k-${t.kind}`} title={KIND_LABEL[t.kind]}>
                <ChromeIcon k={t.kind} />
              </span>
              <span className="chrome-tab-name">{t.name}</span>
              {t.dirty ? <span className="chrome-tab-dot" title="Unsaved changes" /> : null}
              <button
                className="chrome-tab-close"
                title="Close this document"
                onClick={(e) => {
                  e.stopPropagation();
                  void closeTab(t.id);
                }}
              >
                ×
              </button>
            </div>
          ))}
        </div>
        <button className="chrome-plus-btn" onClick={() => void openViaDialog()} title="Open a document">
          +
        </button>
        <span className="chrome-spacer" />
        <div className="chrome-seg">
          {LENS_KINDS.map((k) => (
            <button
              key={k}
              className={`chrome-seg-btn ${activeTab?.kind === k ? 'seg-on' : ''}`}
              onClick={() => switchKind(k)}
              title={`New ${KIND_LABEL[k]}`}
            >
              <ChromeIcon k={k} />
              <span>{KIND_LABEL[k]}</span>
            </button>
          ))}
        </div>
      </div>

      {activeId === null ? (
        <Home
          onCreate={(k) => switchKind(k)}
          onOpen={() => void openViaDialog()}
          onOpenPath={(p) => void doOpenPath(p)}
          recents={recents}
          resume={resume}
          onResume={(k) => switchKind(k)}
        />
      ) : null}

      {tabs.map((t) => (
        <div key={t.id} className={`lens-pane ${t.id === activeId ? '' : 'lens-hidden'}`}>
          {t.kind === 'sheet' ? (
            <SheetLens
              ref={(h) => {
                if (h) lensRefs.current.set(t.id, h);
                else lensRefs.current.delete(t.id);
              }}
              {...mkLensProps(t)}
            />
          ) : null}
          {t.kind === 'pdf' ? (
            <PdfLens
              ref={(h) => {
                if (h) lensRefs.current.set(t.id, h);
                else lensRefs.current.delete(t.id);
              }}
              {...mkLensProps(t)}
            />
          ) : null}
          {t.kind === 'docx' ? (
            <DocxLens
              ref={(h) => {
                if (h) lensRefs.current.set(t.id, h);
                else lensRefs.current.delete(t.id);
              }}
              {...mkLensProps(t)}
            />
          ) : null}
        </div>
      ))}

      {busy && (
        <div className="celldeck-busy">
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <div className="celldeck-spinner" />
            <span>{busy}</span>
          </div>
        </div>
      )}
      {status && <div className="celldeck-status">{status}</div>}
      {dragging && (
        <div className="celldeck-dropzone">
          <div className="celldeck-dropzone-box">
            <div className="celldeck-dropzone-icon">⤓</div>
            Drop to open in DocDeck
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
