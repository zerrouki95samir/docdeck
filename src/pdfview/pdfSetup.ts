import * as pdfjsLib from 'pdfjs-dist';
import workerUrl from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import 'pdfjs-dist/web/pdf_viewer.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = workerUrl;

/* pdf.js v4+ exposes TextLayer at the root; keep it optional at runtime so older
   deployments degrade to no text selection instead of failing to open. */
type TextLayerCtor = new (opts: {
  textContentSource: unknown;
  container: HTMLElement;
  viewport: unknown;
}) => { render(): Promise<void> };

const TextLayerClass: TextLayerCtor | undefined = (pdfjsLib as unknown as { TextLayer?: TextLayerCtor }).TextLayer;

export { pdfjsLib, TextLayerClass as TextLayer, type TextLayerCtor };
