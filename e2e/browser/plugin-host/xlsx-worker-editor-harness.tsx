import { createRoot } from 'react-dom/client';
import JSZip from 'jszip';
import { XlsxEditor, type XlsxWorkerEditorApi } from '@betteroffice/xlsx-react';
import { editableWorkbookSessionBackend } from '../../../packages/xlsx-react/src/worker/useEditableSessionWorkbook';
import type { WorkerEditorProbe } from './xlsx-worker-editor-probe';

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((done) => { resolve = done; });
  return { promise, resolve };
}

async function workbook() {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
  zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Sheet" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
  zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="sans-serif"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>`);
  zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="24"/>
  <cols><col min="1" max="2" width="24" customWidth="1"/></cols>
  <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>initial</t></is></c><c r="B1"><v>1</v></c></row></sheetData>
</worksheet>`);
  return zip.generateAsync({ type: 'uint8array' });
}

const paints = new WeakMap<HTMLCanvasElement, Set<string>>();
const nativeClear = CanvasRenderingContext2D.prototype.clearRect;
CanvasRenderingContext2D.prototype.clearRect = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeClear>) {
  nativeClear.apply(this, args);
  if (this.canvas.closest('#worker-editor-root')) paints.set(this.canvas, new Set());
};
const nativeText = CanvasRenderingContext2D.prototype.fillText;
CanvasRenderingContext2D.prototype.fillText = function (this: CanvasRenderingContext2D, ...args: Parameters<typeof nativeText>) {
  nativeText.apply(this, args);
  paints.get(this.canvas)?.add(args[0]);
};

const rootElement = document.getElementById('worker-editor-root')!;
const root = createRoot(rootElement);
const commitOrder: WorkerEditorProbe['commitOrder'] = [];
const attach = editableWorkbookSessionBackend.attach;
editableWorkbookSessionBackend.attach = (options) => {
  const edit = options.peer.editCell;
  options.peer.editCell = (...args) => {
    commitOrder.push({ kind: 'mutator-entry', text: args[3] });
    return edit(...args);
  };
  return attach(options);
};
const nativeFrame = globalThis.requestAnimationFrame.bind(globalThis);
const nativeCancel = globalThis.cancelAnimationFrame.bind(globalThis);
const heldFrames = new Map<number, FrameRequestCallback>();
let holdingPreview = false;
let previewFrame: number | null = null;
globalThis.requestAnimationFrame = (callback) => {
  const id = nativeFrame((time) => {
    if (holdingPreview && rootElement.querySelector('[data-testid="xlsx-commit-preview"]')) {
      if (previewFrame === null) previewFrame = time;
      else if (previewFrame !== time) { heldFrames.set(id, callback); return; }
    }
    callback(time);
  });
  return id;
};
globalThis.cancelAnimationFrame = (id) => { heldFrames.delete(id); nativeCancel(id); };
const errors: string[] = [];
const previews: string[] = [];
const previewFrames: string[] = [];
const observer = new MutationObserver(() => {
  const preview = rootElement.querySelector('[data-testid="xlsx-commit-preview"]');
  if (!preview?.textContent) return;
  const text = preview.textContent;
  previews.push(text);
  requestAnimationFrame(() => {
    if (preview.isConnected && preview.textContent === text) previewFrames.push(text);
  });
});
observer.observe(rootElement, { subtree: true, childList: true, characterData: true });
let current = deferred<XlsxWorkerEditorApi>();
let api: XlsxWorkerEditorApi | null = null;
const open = (bytes: Uint8Array) => {
  root.render(<XlsxEditor file={bytes} experimentalWorkerOpen onError={(error) => errors.push(error.message)}
    onReady={(value) => { api = value; current.resolve(value); }} />);
};
const ready = (async () => {
  await document.fonts.ready;
  open(await workbook());
  const editor = await current.promise;
  await editor.whenHydrated();
})();

window.__xlsxWorkerEditor = {
  ready, errors, previews, previewFrames, commitOrder,
  holdPreview() { commitOrder.length = 0; previewFrame = null; holdingPreview = true; },
  previewHeld() { return heldFrames.size > 0; },
  releasePreview(text) {
    if (commitOrder.length > 0) throw new Error('Mutation preceded the preview screenshot');
    if (rootElement.querySelector('[data-testid="xlsx-commit-preview"]')?.textContent !== text) {
      throw new Error('Pending text changed before the preview screenshot');
    }
    commitOrder.push({ kind: 'painted-preview', text });
    holdingPreview = false;
    for (const callback of heldFrames.values()) nativeFrame(callback);
    heldFrames.clear();
  },
  paintedTexts() {
    const canvas = rootElement.querySelector<HTMLCanvasElement>('[data-testid="xlsx-scroll"] canvas');
    return canvas ? [...(paints.get(canvas) ?? [])] : [];
  },
  async cell() { return (await api?.cellAsync(0, 0, 0))?.input ?? null; },
  async saveAndReopen() {
    const bytes = await api!.save();
    if (!bytes) throw new Error('Workbook save returned no bytes');
    api = null;
    current = deferred<XlsxWorkerEditorApi>();
    open(bytes);
    const editor = await current.promise;
    await editor.whenHydrated();
    return bytes.byteLength;
  },
  async undo() {
    const result = await api!.commands.execute('undo', null);
    return result.ok;
  },
};
