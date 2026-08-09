import initWasmModule, { VsdxDocument, VsdxRenderer, rendererVersion } from './generated/vsdx_wasm.js';
import type { InitInput } from './generated/vsdx_wasm.js';
import type { CellFormulaReceipt, CollaborationUpdateOrigin, DiagramSnapshot, FormulaShapeDraft, HistoryResult, HitTestResult, PageDisplayList, ShapeReceipt, VsdxFontFace } from '../types';

export type WasmInitInput = InitInput | Promise<InitInput>;
export interface OpenDiagramOptions { clientId?: number; fonts?: ReadonlyArray<VsdxFontFace>; }
export interface DiagramHandle {
  readonly clientId: number;
  snapshot(): DiagramSnapshot;
  registerFont(face: VsdxFontFace): number;
  layoutPage(pageIndex: number): PageDisplayList;
  hitTest(x: number, y: number): HitTestResult | null;
  mediaBytes(assetId: string): Uint8Array;
  setCellFormula(pageId: string, shapeId: string, locator: { section?: string; rowIndex?: number; rowName?: string; cellName: string }, formula: string): CellFormulaReceipt;
  moveShape(pageId: string, shapeId: string, xFormula: string, yFormula: string): [CellFormulaReceipt, CellFormulaReceipt];
  resizeShape(pageId: string, shapeId: string, widthFormula: string, heightFormula: string): [CellFormulaReceipt, CellFormulaReceipt];
  reorderShape(pageId: string, shapeId: string, toIndex: number): ShapeReceipt;
  reorderPage(pageId: string, toIndex: number): ShapeReceipt;
  addShape(pageId: string, draft: FormulaShapeDraft): ShapeReceipt;
  canUndo(): boolean; canRedo(): boolean; undo(): HistoryResult; redo(): HistoryResult;
  encodeStateVector(): Uint8Array; encodeStateAsUpdate(): Uint8Array; encodeDiff(vector: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): DiagramSnapshot;
  onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
  dispose(): void;
}
let initialized = false;
let initialization: Promise<void> | undefined;
export function initWasm(input: WasmInitInput = new URL('./generated/vsdx_wasm_bg.wasm', import.meta.url)): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(() => { initialized = true; }, error => { initialization = undefined; throw toError(error); });
  return initialization;
}
export function isWasmAvailable(): boolean { return typeof WebAssembly === 'object'; }
export function wasmVersion(): string { requireInitialized(); return rendererVersion(); }
export function openDiagram(bytes: Uint8Array, options: OpenDiagramOptions = {}): DiagramHandle {
  requireInitialized();
  const doc = construct(() => VsdxDocument.openCollaborative(bytes, options.clientId ?? clientId()));
  const renderer = construct(() => new VsdxRenderer());
  for (const face of options.fonts ?? []) construct(() => renderer.registerFont(face.family, face.bold ?? false, face.italic ?? false, face.bytes));
  const listeners = new Map<number, (update: Uint8Array, origin: CollaborationUpdateOrigin) => void>();
  const queued: Array<{ update: Uint8Array; origin: CollaborationUpdateOrigin }> = [];
  let nextListener = 0, disposed = false, observing = false, depth = 0, flushing = false;
  const assertAlive = () => { if (disposed) throw new Error('diagram handle is disposed'); };
  const resync = () => { queued.length = 0; doc.encodeStateAsUpdate(); };
  const drain = () => {
    if (!observing || disposed) return;
    for (;;) {
      const event = doc.drainUpdateEvent();
      if (event.byteLength === 0) return;
      if (event.byteLength === 1 && event[0] === 2) { resync(); continue; }
      const origin = event[0];
      if (origin !== 0 && origin !== 1) throw new Error(`vsdx wasm returned unknown update origin ${origin}`);
      queued.push({ update: event.slice(1), origin: origin === 0 ? 'local' : 'remote' });
    }
  };
  const flush = () => {
    if (disposed || flushing || depth !== 0) return;
    flushing = true;
    try { while (!disposed && queued.length) { const event = queued.shift(); if (!event) break; for (const [id, listener] of [...listeners]) if (listeners.get(id) === listener) try { listener(event.update.slice(), event.origin); } catch {} } }
    finally { flushing = false; if (disposed) queued.length = 0; }
  };
  const wasm = <T>(operation: () => T, drainUpdates = false): T => {
    assertAlive(); depth++;
    try { const result = operation(); if (drainUpdates) drain(); return result; }
    finally { depth--; if (depth === 0) flush(); }
  };
  const json = <T>(operation: () => string, drainUpdates = false): T => JSON.parse(wasm(operation, drainUpdates)) as T;
  return {
    get clientId() { return wasm(() => doc.clientId); }, snapshot: () => json(() => doc.snapshotJson()),
    registerFont: face => wasm(() => renderer.registerFont(face.family, face.bold ?? false, face.italic ?? false, face.bytes)),
    layoutPage: pageIndex => { const list = json<PageDisplayList>(() => renderer.layoutPageJson(doc, pageIndex)); if (list.contractVersion !== 3) throw new Error(`unsupported VSDX display-list contract version ${list.contractVersion}`); return list; },
    hitTest: (x, y) => json<HitTestResult | null>(() => renderer.hitTestJson(x, y)), mediaBytes: assetId => wasm(() => doc.mediaBytes(assetId).slice()),
    setCellFormula: (pageId, shapeId, locator, formula) => json(() => doc.setCellFormulaJson(JSON.stringify({ pageId, shapeId, locator, formula })), true),
    moveShape: (pageId, shapeId, xFormula, yFormula) => json(() => doc.moveShapeJson(JSON.stringify({ pageId, shapeId, xFormula, yFormula })), true),
    resizeShape: (pageId, shapeId, widthFormula, heightFormula) => json(() => doc.resizeShapeJson(JSON.stringify({ pageId, shapeId, widthFormula, heightFormula })), true),
    reorderShape: (pageId, shapeId, toIndex) => json(() => doc.reorderShapeJson(JSON.stringify({ pageId, shapeId, toIndex })), true),
    reorderPage: (pageId, toIndex) => json(() => doc.reorderPageJson(JSON.stringify({ pageId, toIndex })), true),
    addShape: (pageId, draft) => json(() => doc.addShapeJson(JSON.stringify({ pageId, draft })), true),
    canUndo: () => wasm(() => doc.canUndo()), canRedo: () => wasm(() => doc.canRedo()), undo: () => json(() => doc.undoJson(), true), redo: () => json(() => doc.redoJson(), true),
    encodeStateVector: () => wasm(() => doc.encodeStateVector().slice()), encodeStateAsUpdate: () => wasm(() => doc.encodeStateAsUpdate().slice()), encodeDiff: vector => wasm(() => doc.encodeDiff(vector.slice()).slice()), applyUpdate: update => json(() => doc.applyUpdateJson(update.slice()), true),
    onUpdate(listener) { assertAlive(); if (typeof listener !== 'function') throw new TypeError('update listener must be a function'); const id = nextListener++; listeners.set(id, listener); if (!observing) { wasm(() => doc.startUpdateObservation()); observing = true; } return () => { listeners.delete(id); if (!listeners.size && observing && !disposed) { queued.length = 0; wasm(() => doc.clearUpdateObservation()); observing = false; } }; },
    dispose() { if (disposed) return; disposed = true; listeners.clear(); queued.length = 0; let error: unknown; if (observing) try { doc.clearUpdateObservation(); } catch (caught) { error = caught; } try { renderer.free(); } catch (caught) { error ??= caught; } try { doc.free(); } catch (caught) { error ??= caught; } if (error) throw toError(error); },
  };
}
function requireInitialized(): void { if (!initialized) throw new Error('vsdx wasm is not initialized; call initWasm() first'); }
function clientId(): number { if (!globalThis.crypto?.getRandomValues) throw new Error('crypto.getRandomValues is required to generate a collaboration client ID'); const values = new Uint32Array(2); let value = 0; do { crypto.getRandomValues(values); value = (values[0] & 0x1fffff) * 0x1_0000_0000 + values[1]; } while (!value); return value; }
function construct<T>(operation: () => T): T { try { return operation(); } catch (error) { throw toError(error); } }
function toError(error: unknown): Error { return error instanceof Error ? error : new Error(typeof error === 'string' ? error : String(error)); }
