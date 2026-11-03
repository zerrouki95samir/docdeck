import { forwardRef, useEffect, useImperativeHandle, useRef } from 'react';
import type { FUniver } from '@univerjs/presets';
import { initUniver } from '../univerSetup';
import { parseSpreadsheet, serializeSpreadsheet, extOfPath } from '../io/workbookIO';
import type { LensHandle, LensProps } from '../lenses';

const IGNORE_COMMANDS = [
  'sheet.command.set-zoom-ratio',
  'univer.command.set-menu',
  'sheet.command.scroll-command',
  'sheet.command.set-scroll',
];

interface InnerFile {
  name: string;
  path: string | null;
}

export const SheetLens = forwardRef<LensHandle, LensProps>(function SheetLens(props, ref) {
  const hostRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<FUniver | null>(null);
  const fileIdRef = useRef<string | null>(null);
  const fileRef = useRef<InnerFile>({ name: 'Untitled.xlsx', path: null });
  const dirtyRef = useRef(false);
  const propsRef = useRef(props);
  propsRef.current = props;

  const snapshotSerialize = async (target: string) => {
    const api = apiRef.current;
    const fwb = api?.getActiveWorkbook();
    if (!api || !fwb) throw new Error('spreadsheet not ready');
    const snapshot = (await Promise.resolve(fwb.save())) as never;
    return await serializeSpreadsheet(snapshot, fwb.getActiveSheet().getSheetId(), target);
  };

  useImperativeHandle<LensHandle, LensHandle>(ref, () => ({
    kind: 'sheet',
    async open(name: string, bytes: Uint8Array, path: string | null) {
      const api = initReady();
      const snapshot = await parseSpreadsheet(name, bytes);
      setSnapshot(snapshot as unknown, { name, path });
      void api;
    },
    async save(asNew: boolean) {
      const f = fileRef.current;
      let path = f.path;
      if (!path || asNew) {
        const base = f.path ? f.name : 'Untitled.xlsx';
        path = await window.cellDesk!.saveDialog({
          defaultPath: base,
          filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }],
        });
        if (!path) return false;
      }
      const data = (await snapshotSerialize(path)) as Uint8Array;
      await window.cellDesk!.writeFile(path, data);
      fileRef.current = { name: basename(path), path };
      propsRef.current.onTitle(basename(path));
      return true;
    },
    async makeNew() {
      initReady();
      setSnapshot({}, { name: 'Untitled.xlsx', path: null });
    },
    async exportAs(kind) {
      const f = fileRef.current;
      const ext = kind === 'csv' ? 'csv' : 'uni';
      const base = f.path ? stripExtOnly(f.name) : stripExtOnly(f.name);
      const filters =
        kind === 'csv'
          ? [{ name: 'CSV (Current Sheet)', extensions: ['csv'] }]
          : [{ name: 'DocDeck Snapshot', extensions: ['uni'] }];
      const target = await window.cellDesk!.saveDialog({
        defaultPath: `${base}.${ext}`,
        filters,
      });
      if (!target) return;
      const data = (await snapshotSerialize(target)) as Uint8Array;
      await window.cellDesk!.writeFile(target, data);
      if (kind === 'uni') {
        fileRef.current = { name: basename(target), path: target };
        propsRef.current.onTitle(basename(target));
      }
      propsRef.current.onStatus(`Saved ${basename(target)}`);
    },
    async print() {
      propsRef.current.onStatus('Printing spreadsheets is not supported yet');
    },
    isDirty() {
      return dirtyRef.current;
    },
    fileName() {
      return fileRef.current.name;
    },
  }));
  // eslint-disable-next-line react-hooks/exhaustive-deps

  const initReady = (): FUniver => {
    const api = apiRef.current;
    if (!api) throw new Error('spreadsheet engine not ready');
    return api;
  };

  const sheetDirtyListener = () => {
    dirtyRef.current = true;
    propsRef.current.onDirty();
  };

  const setSnapshot = (snapshot: unknown, meta: InnerFile) => {
    const api = initReady();
    const oldId = fileIdRef.current;
    const fwb = api.createWorkbook(snapshot as never);
    if (oldId) {
      try {
        api.disposeUnit(oldId);
      } catch {
        /* old workbook already gone */
      }
    }
    fileIdRef.current = fwb.getId();

    fwb.onCommandExecuted((command) => {
      const id = (command as { id?: string }).id || '';
      if (!id.startsWith('sheet.command.') && !id.startsWith('sheet.mutation') && !id.startsWith('doc.mutation')) return;
      if (IGNORE_COMMANDS.some((g) => id === g) || id.includes('zoom') || id.includes('scroll')) return;
      sheetDirtyListener();
    });

    fileRef.current = meta;
    dirtyRef.current = false;
    propsRef.current.onTitle(meta.name);
  };

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return;
    const container = document.createElement('div');
    container.style.height = '100%';
    container.style.width = '100%';
    host.append(container);

    const { univerAPI } = initUniver(container);
    apiRef.current = univerAPI;

    setSnapshot({}, { name: 'Untitled.xlsx', path: null });

    return () => {
      try {
        (univerAPI as unknown as { dispose?: () => void }).dispose?.();
      } catch {
        /* ignore */
      }
      delete (window as unknown as Record<string, unknown>).__cellDeskDev;
      container.remove();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* dev hooks — only the active tab publishes them (multi-tab safety) */
  useEffect(() => {
    if (props.active === false) return;
    (window as unknown as Record<string, unknown>).__cellDeskDev = {
      openBuffer: async (name: string, bytes: ArrayBuffer) => {
        const snapshot = await parseSpreadsheet(name, new Uint8Array(bytes));
        setSnapshot(snapshot as unknown, { name, path: null });
        return {
          sheets: Object.keys(snapshot.sheets || {}).length,
          title: fileRef.current.name,
          a1: (snapshot.sheets?.[snapshot.sheetOrder![0] as string]?.cellData?.['0']?.['0']?.v ?? null) as unknown,
        };
      },
      saveBytes: async (targetName: string) => (await snapshotSerialize(targetName)) as Uint8Array,
    };
    return () => {
      delete (window as unknown as Record<string, unknown>).__cellDeskDev;
    };
  }, [props.active]);

  return <div ref={hostRef} className="lens-host" />;
});

function basename(p: string): string {
  return p.split('/').pop() || p;
}

function stripExtOnly(name: string): string {
  const dot = name.lastIndexOf('.');
  return dot > 0 ? name.slice(0, dot) : name;
}

export { extOfPath };
