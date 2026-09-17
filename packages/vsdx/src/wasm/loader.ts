import initWasmModule, { VsdxDocument, VsdxRenderer, rendererVersion } from './generated/vsdx_wasm.js';
import type { InitInput } from './generated/vsdx_wasm.js';
import type { CellLocator, CellFormulaReceipt, CollaborationUpdateOrigin, ConnectedShapeReceipt, ConnectorGlue, ConnectorRoutePoint, ConnectorRouteReceipt, DiagramSnapshot, DocumentMaster, RawValidationIssue, FormulaShapeDraft, FormulaShapeTreeDraft, HistoryResult, HitTestResult, PageDisplayList, PageLayer, ShapeReceipt, ShapeTreeGlue, TextReceipt, ValidationIssue, VsdxFontFace } from '../types';

export type WasmInitInput = InitInput | Promise<InitInput>;
export interface OpenDiagramOptions { clientId?: number; fonts?: ReadonlyArray<VsdxFontFace>; initialUpdate?: Uint8Array; }
export interface CollaborationResync { update: Uint8Array; }
export interface ShapeMove { pageId: string; shapeId: string; xFormula: string; yFormula: string; }
export interface ShapeDelete { pageId: string; shapeId: string; }
export interface CellFormulaWrite { pageId: string; shapeId: string; cellName: string; formula: string; }
export interface DiagramHandle {
  readonly clientId: number;
  snapshot(): DiagramSnapshot;
  masters(): DocumentMaster[];
  registerFont(face: VsdxFontFace): number;
  layoutPage(pageIndex: number): PageDisplayList;
  exportPdf(): Uint8Array;
  exportSvg(): string[];
  exportPng(pageIndex: number, scale?: number): Uint8Array;
  pageLayers(pageIndex: number): PageLayer[];
  validate(): ValidationIssue[];
  validatePage(pageIndex: number): ValidationIssue[];
  setLayerVisible(pagePartPath: string, layerIndex: number, visible: boolean): void;
  clearLayerVisibility(): void;
  hitTest(x: number, y: number): HitTestResult | null;
  mediaBytes(assetId: string): Uint8Array;
  setCellFormula(pageId: string, shapeId: string, locator: CellLocator, formula: string): CellFormulaReceipt;
  setCellFormulas(writes: ReadonlyArray<CellFormulaWrite>): CellFormulaReceipt[];
  setControlHandle(pageId: string, shapeId: string, row: string, xFormula: string | null, yFormula: string | null): CellFormulaReceipt[];
  moveShape(pageId: string, shapeId: string, xFormula: string, yFormula: string): [CellFormulaReceipt, CellFormulaReceipt];
  moveShapes(moves: ReadonlyArray<ShapeMove>): Array<[CellFormulaReceipt, CellFormulaReceipt]>;
  setShapeBounds(pageId: string, shapeId: string, xFormula: string, yFormula: string, widthFormula: string, heightFormula: string): [CellFormulaReceipt, CellFormulaReceipt, CellFormulaReceipt, CellFormulaReceipt];
  resizeLocPin(pageId: string, shapeId: string, width: number, height: number): { x: number; y: number };
  resizeShape(pageId: string, shapeId: string, widthFormula: string, heightFormula: string): [CellFormulaReceipt, CellFormulaReceipt];
  deleteShapes(deletes: ReadonlyArray<ShapeDelete>): ShapeReceipt[];
  reorderShape(pageId: string, shapeId: string, toIndex: number): ShapeReceipt;
  reorderPage(pageId: string, toIndex: number): ShapeReceipt;
  addShape(pageId: string, draft: FormulaShapeDraft): ShapeReceipt;
  addShapeWithText(pageId: string, draft: FormulaShapeDraft, text: string): ShapeReceipt;
  addShapeTree(pageId: string, draft: FormulaShapeTreeDraft): ShapeReceipt;
  subtreeGlue(pageId: string, shapeId: string): ShapeTreeGlue[];
  addConnector(pageId: string, draft: FormulaShapeDraft, from: ConnectorGlue, to: ConnectorGlue): ShapeReceipt;
  addFreeConnector(pageId: string, draft: FormulaShapeDraft, from: ConnectorGlue): ShapeReceipt;
  addConnectedShape(pageId: string, shapeDraft: FormulaShapeDraft, connectorDraft: FormulaShapeDraft, from: ConnectorGlue, toCell?: string): ConnectedShapeReceipt;
  setConnectorRoute(pageId: string, shapeId: string, points: ConnectorRoutePoint[]): ConnectorRouteReceipt;
  deleteShape(pageId: string, shapeId: string): ShapeReceipt;
  shapeText(pageId: string, shapeId: string): string;
  setShapeText(pageId: string, shapeId: string, text: string): TextReceipt;
  save(): Uint8Array;
  canUndo(): boolean; canRedo(): boolean; undo(): HistoryResult; redo(): HistoryResult;
  encodeStateVector(): Uint8Array; encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array; encodeDiff(vector: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): DiagramSnapshot;
  onUpdate(listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void): () => void;
  onResync(listener: (resync: CollaborationResync) => void): () => void;
  dispose(): void;
}
const MAX_QUEUED_UPDATES = 1024;
const MAX_QUEUED_UPDATE_BYTES = 8 * 1024 * 1024;

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
  let renderer!: VsdxRenderer;
  try {
    if (options.initialUpdate) construct(() => doc.applyUpdateJson(options.initialUpdate!.slice()));
    renderer = construct(() => new VsdxRenderer());
    for (const face of options.fonts ?? []) construct(() => renderer!.registerFont(face.family, face.bold ?? false, face.italic ?? false, face.bytes));
  } catch (error) {
    try { renderer?.free(); } catch {}
    try { doc.free(); } catch {}
    throw error;
  }
  const hitIds = new Map<string, string>();
  const listeners = new Map<number, (update: Uint8Array, origin: CollaborationUpdateOrigin) => void>();
  const resyncListeners = new Map<number, (resync: CollaborationResync) => void>();
  const queued: Array<{ kind: 'update'; update: Uint8Array; origin: CollaborationUpdateOrigin } | { kind: 'resync' }> = [];
  let nextListener = 0, disposed = false, observing = false, depth = 0, flushing = false;
  let queuedBytes = 0, resyncPending = false;
  const clearQueued = () => { queued.length = 0; queuedBytes = 0; resyncPending = false; };
  const assertAlive = () => { if (disposed) throw new Error('diagram handle is disposed'); };
  const resync = () => { clearQueued(); resyncPending = true; queued.push({ kind: 'resync' }); };
  const drain = () => {
    if (!observing || disposed) return;
    for (;;) {
      const event = doc.drainUpdateEvent();
      if (event.byteLength === 0) return;
      if (event.byteLength === 1 && event[0] === 2) { resync(); continue; }
      const origin = event[0];
      if (origin !== 0 && origin !== 1) throw new Error(`vsdx wasm returned unknown update origin ${origin}`);
      if (resyncPending) continue;
      if (queued.length >= MAX_QUEUED_UPDATES || event.byteLength - 1 > MAX_QUEUED_UPDATE_BYTES - queuedBytes) { resync(); continue; }
      queuedBytes += event.byteLength - 1;
      queued.push({ kind: 'update', update: event.slice(1), origin: origin === 0 ? 'local' : 'remote' });
    }
  };
  const flush = () => {
    if (disposed || flushing || depth !== 0) return;
    flushing = true;
    try {
      while (!disposed && queued.length) {
        const event = queued.shift();
        if (!event) break;
        if (event.kind === 'update') {
          queuedBytes -= event.update.byteLength;
          for (const [id, listener] of [...listeners]) {
            if (listeners.get(id) === listener) try { listener(event.update.slice(), event.origin); } catch {}
          }
        } else {
          resyncPending = false;
          const update = doc.encodeStateAsUpdate().slice();
          for (const [id, listener] of [...resyncListeners]) {
            if (resyncListeners.get(id) === listener) try { listener({ update: update.slice() }); } catch {}
          }
        }
      }
    } finally { flushing = false; if (disposed) clearQueued(); }
  };
  const wasm = <T>(operation: () => T, drainUpdates = false): T => {
    assertAlive(); depth++;
    try { const result = operation(); if (drainUpdates) drain(); return result; }
    finally { depth--; if (depth === 0) flush(); }
  };
  const json = <T>(operation: () => string, drainUpdates = false): T => JSON.parse(wasm(operation, drainUpdates)) as T;
  return {
    clientId: doc.clientId, snapshot: () => json(() => doc.snapshotJson()),
    masters: () => {
      const masters = json<DocumentMaster[]>(() => renderer.masterPreviewsJson(doc));
      for (const master of masters) {
        if (master.display && master.display.contractVersion !== 7) throw new Error(`unsupported VSDX display-list contract version ${master.display.contractVersion}`);
      }
      return masters;
    },
    registerFont: face => wasm(() => renderer.registerFont(face.family, face.bold ?? false, face.italic ?? false, face.bytes)),
    pageLayers: pageIndex => json(() => renderer.pageLayersJson(doc, pageIndex)),
    validate: () => mapValidationIssues(json<RawValidationIssue[]>(() => renderer.validateJson(doc)), json<DiagramSnapshot>(() => doc.snapshotJson())),
    validatePage: (pageIndex) => {
      const snapshot = json<DiagramSnapshot>(() => doc.snapshotJson());
      const page = snapshot.pages[pageIndex];
      if (!page) throw new Error('page index is outside the document');
      return mapValidationIssues(json<RawValidationIssue[]>(() => renderer.validatePageJson(doc, pageIndex)), snapshot).filter((issue) => issue.pageId === page.id);
    },
    setLayerVisible: (pagePartPath, layerIndex, visible) => wasm(() => renderer.setLayerVisible(pagePartPath, layerIndex, visible)),
    clearLayerVisibility: () => wasm(() => renderer.clearLayerVisibility()),
    layoutPage: pageIndex => {
      hitIds.clear();
      const list = json<PageDisplayList>(() => renderer.layoutPageJson(doc, pageIndex));
      if (list.contractVersion !== 7) throw new Error(`unsupported VSDX display-list contract version ${list.contractVersion}`);
      const page = json<DiagramSnapshot>(() => doc.snapshotJson()).pages[pageIndex];
      const shapes = [...page.shapes];
      while (shapes.length) { const shape = shapes.pop()!; hitIds.set(`${page.sourcePartPath}:${shape.sourceId}`, shape.id); shapes.push(...shape.children); }
      return list;
    },
    exportPdf: () => wasm(() => renderer.exportPdf(doc).slice()),
    exportSvg: () => json<string[]>(() => renderer.exportSvgJson(doc)),
    exportPng: (pageIndex, scale = 1) => {
      if (!Number.isInteger(pageIndex) || pageIndex < 0) throw new Error('VSDX page index must be a non-negative integer');
      if (!Number.isFinite(scale) || scale <= 0) throw new Error('VSDX PNG scale must be a positive number');
      return wasm(() => renderer.exportPng(doc, pageIndex, scale).slice());
    },
    hitTest: (x, y) => {
      const hit = json<HitTestResult | null>(() => renderer.hitTestJson(x, y));
      const shapeId = hit && hitIds.get(hit.shapeId);
      return hit && shapeId ? { ...hit, shapeId } : null;
    }, mediaBytes: assetId => wasm(() => doc.mediaBytes(assetId).slice()),
    setCellFormula: (pageId, shapeId, locator, formula) => json(() => doc.setCellFormulaJson(JSON.stringify({ pageId, shapeId, locator, formula })), true),
    setCellFormulas: (writes) => json(() => doc.setCellFormulasJson(JSON.stringify({ writes: [...writes] })), true),
    setControlHandle: (pageId, shapeId, row, xFormula, yFormula) => json(() => doc.setControlHandleJson(JSON.stringify({ pageId, shapeId, row, xFormula, yFormula })), true),
    moveShape: (pageId, shapeId, xFormula, yFormula) => json(() => doc.moveShapeJson(JSON.stringify({ pageId, shapeId, xFormula, yFormula })), true),
    moveShapes: (moves) => json(() => doc.moveShapesJson(JSON.stringify({ moves: [...moves] })), true),
    setShapeBounds: (pageId, shapeId, xFormula, yFormula, widthFormula, heightFormula) => json(() => doc.setShapeBoundsJson(JSON.stringify({ pageId, shapeId, xFormula, yFormula, widthFormula, heightFormula })), true),
    resizeLocPin: (pageId, shapeId, width, height) => { const [x, y] = wasm(() => doc.resizeLocPin(pageId, shapeId, width, height)); return { x, y }; },
    resizeShape: (pageId, shapeId, widthFormula, heightFormula) => json(() => doc.resizeShapeJson(JSON.stringify({ pageId, shapeId, widthFormula, heightFormula })), true),
    reorderShape: (pageId, shapeId, toIndex) => json(() => doc.reorderShapeJson(JSON.stringify({ pageId, shapeId, toIndex })), true),
    reorderPage: (pageId, toIndex) => json(() => doc.reorderPageJson(JSON.stringify({ pageId, toIndex })), true),
    addShape: (pageId, draft) => json(() => doc.addShapeJson(JSON.stringify({ pageId, draft })), true),
    addShapeWithText: (pageId, draft, text) => json(() => doc.addShapeWithTextJson(JSON.stringify({ pageId, draft, text })), true),
    addShapeTree: (pageId, draft) => json(() => doc.addShapeTreeJson(JSON.stringify({ pageId, draft })), true),
    subtreeGlue: (pageId, shapeId) => json(() => doc.subtreeGlueJson(JSON.stringify({ pageId, shapeId }))),
    addConnector: (pageId, draft, from, to) => json(() => doc.addConnectorJson(JSON.stringify({ pageId, draft, from, to })), true),
    addFreeConnector: (pageId, draft, from) => json(() => doc.addFreeConnectorJson(JSON.stringify({ pageId, draft, from })), true),
    addConnectedShape: (pageId, shapeDraft, connectorDraft, from, toCell) => json(() => doc.addConnectedShapeJson(JSON.stringify({ pageId, shapeDraft, connectorDraft, from, toCell })), true),
    setConnectorRoute: (pageId, shapeId, points) => json(() => doc.setConnectorRouteJson(JSON.stringify({ pageId, shapeId, points })), true),
    deleteShape: (pageId, shapeId) => json(() => doc.deleteShapeJson(JSON.stringify({ pageId, shapeId })), true),
    deleteShapes: (deletes) => json(() => doc.deleteShapesJson(JSON.stringify({ deletes: [...deletes] })), true),
    shapeText: (pageId, shapeId) => json(() => doc.shapeTextJson(JSON.stringify({ pageId, shapeId }))),
    setShapeText: (pageId, shapeId, text) => json(() => doc.setShapeTextJson(JSON.stringify({ pageId, shapeId, text })), true),
    save: () => wasm(() => doc.save().slice()),
    canUndo: () => wasm(() => doc.canUndo()), canRedo: () => wasm(() => doc.canRedo()), undo: () => json(() => doc.undoJson(), true), redo: () => json(() => doc.redoJson(), true),
    encodeStateVector: () => wasm(() => doc.encodeStateVector().slice()), encodeStateAsUpdate: vector => wasm(() => (vector === undefined ? doc.encodeStateAsUpdate() : doc.encodeDiff(vector.slice())).slice()), encodeDiff: vector => wasm(() => doc.encodeDiff(vector.slice()).slice()), applyUpdate: update => json(() => doc.applyUpdateJson(update.slice()), true),
    onUpdate(listener) { assertAlive(); if (typeof listener !== 'function') throw new TypeError('update listener must be a function'); const id = nextListener++; listeners.set(id, listener); if (!observing) { wasm(() => doc.startUpdateObservation()); observing = true; } return () => { listeners.delete(id); if (!listeners.size && !resyncListeners.size && observing && !disposed) { clearQueued(); wasm(() => doc.clearUpdateObservation()); observing = false; } }; },
    onResync(listener) { assertAlive(); if (typeof listener !== 'function') throw new TypeError('resync listener must be a function'); const id = nextListener++; resyncListeners.set(id, listener); if (!observing) { wasm(() => doc.startUpdateObservation()); observing = true; } return () => { resyncListeners.delete(id); if (!listeners.size && !resyncListeners.size && observing && !disposed) { clearQueued(); wasm(() => doc.clearUpdateObservation()); observing = false; } }; },
    dispose() { if (disposed) return; disposed = true; listeners.clear(); resyncListeners.clear(); clearQueued(); let error: unknown; if (observing) try { doc.clearUpdateObservation(); } catch (caught) { error = caught; } try { renderer.free(); } catch (caught) { error ??= caught; } try { doc.free(); } catch (caught) { error ??= caught; } if (error) throw toError(error); },
  };
}
function requireInitialized(): void { if (!initialized) throw new Error('vsdx wasm is not initialized; call initWasm() first'); }
function clientId(): number { if (!globalThis.crypto?.getRandomValues) throw new Error('crypto.getRandomValues is required to generate a collaboration client ID'); const values = new Uint32Array(2); let value = 0; do { crypto.getRandomValues(values); value = (values[0] & 0x1fffff) * 0x1_0000_0000 + values[1]; } while (!value); return value; }
function mapValidationIssues(raw: RawValidationIssue[], snapshot: DiagramSnapshot): ValidationIssue[] {
  const pages = new Map(snapshot.pages.map((page) => [page.sourcePartPath, page]));
  const shapes = new Map<string, Map<number, string>>();
  for (const page of snapshot.pages) {
    const index = new Map<number, string>();
    const work = [...page.shapes];
    while (work.length) {
      const shape = work.pop()!;
      if (!index.has(shape.sourceId)) index.set(shape.sourceId, shape.id);
      work.push(...shape.children);
    }
    shapes.set(page.sourcePartPath, index);
  }
  const out: ValidationIssue[] = [];
  for (const issue of raw) {
    const page = pages.get(issue.pagePart);
    const shapeId = shapes.get(issue.pagePart)?.get(issue.shapeId);
    if (!page || !shapeId) continue;
    const otherShapeId = issue.otherShapeId === null ? null : shapes.get(issue.pagePart)?.get(issue.otherShapeId) ?? null;
    out.push({ id: issue.id, rule: issue.rule, severity: issue.severity, pageId: page.id, shapeId, otherShapeId, endpoint: issue.endpoint, row: issue.row });
  }
  return out;
}
function construct<T>(operation: () => T): T { try { return operation(); } catch (error) { throw toError(error); } }
function toError(error: unknown): Error { return error instanceof Error ? error : new Error(typeof error === 'string' ? error : String(error)); }
