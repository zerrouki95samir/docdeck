export interface DocDeckApi {
  platform: string;
  openDialog: (filterKey?: 'pdf' | 'image') => Promise<string[]>;
  saveDialog: (opts?: {
    defaultPath?: string;
    filters?: Array<{ name: string; extensions: string[] }>;
  }) => Promise<string | null>;
  unsavedDialog: () => Promise<0 | 1 | 2>;
  readFile: (path: string) => Promise<ArrayBuffer>;
  writeFile: (path: string, data: Uint8Array) => Promise<boolean>;
  setTitle: (name: string, dirty: boolean) => void;
  confirmClose: () => Promise<void>;
  leaveDialog: () => Promise<0 | 1 | 2>;
  recents: {
    list: () => Promise<string[]>;
    add: (path: string) => Promise<boolean>;
  };
  onMenu: (cb: (action: string) => void) => () => void;
  onOpenPath: (cb: (path: string) => void) => () => void;
  onRecentsChanged: (cb: () => void) => () => void;
  queryDirty: (fn: () => boolean) => void;
}

declare global {
  interface Window {
    cellDesk?: DocDeckApi;
  }
}

export {};
