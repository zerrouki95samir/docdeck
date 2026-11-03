export type DocKind = 'sheet' | 'pdf' | 'docx';

export interface SaveDialogOpts {
  defaultPath: string;
  filters: Array<{ name: string; extensions: string[] }>;
}

export interface LensProps {
  active?: boolean;
  onDirty: () => void;
  onTitle: (name: string) => void;
  onStatus: (msg: string) => void;
  onBusy: (msg: string | null) => void;
}

export interface LensHandle {
  kind: DocKind;
  open(name: string, bytes: Uint8Array, path: string | null): Promise<void>;
  save(asNew: boolean): Promise<boolean>;
  makeNew(): Promise<void>;
  exportAs(kind: 'csv' | 'uni'): Promise<void>;
  print(): Promise<void>;
  isDirty(): boolean;
  fileName(): string;
}
