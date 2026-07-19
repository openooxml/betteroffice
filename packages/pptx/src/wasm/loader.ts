import initWasmModule, {
  parsePptxJson,
  PptxDocument,
  PptxRenderer,
  rendererVersion,
} from './generated/pptx_wasm.js';
import type { InitInput } from './generated/pptx_wasm.js';
import type {
  CollaborationReplica,
  CollaborationUpdateOrigin,
} from '../collaboration/types';
import type {
  DeckSnapshot,
  HistoryResult,
  HitTestResult,
  PptxFontFace,
  ShapeDraft,
  ShapeReceipt,
  SlideDisplayList,
  SlideReceipt,
  StorySnapshot,
  TextReceipt,
  TextStyle,
  TextStylePatch,
  TransformReceipt,
} from '../types';

export type WasmInitInput = InitInput | Promise<InitInput>;

export interface OpenPresentationOptions {
  clientId?: number;
  fonts?: ReadonlyArray<PptxFontFace>;
  initialUpdate?: Uint8Array;
}

export interface PresentationHandle extends CollaborationReplica {
  readonly clientId: number;
  snapshot(): DeckSnapshot;
  story(storyId: string): StorySnapshot;
  registerFont(face: PptxFontFace): number;
  layoutSlide(slideIndex: number): SlideDisplayList;
  hitTest(x: number, y: number): HitTestResult | null;
  mediaBytes(partPath: string): Uint8Array;
  insertText(storyId: string, index: number, text: string, style?: TextStyle): TextReceipt;
  deleteText(storyId: string, start: number, end: number): TextReceipt;
  formatText(storyId: string, start: number, end: number, patch: TextStylePatch): TextReceipt;
  insertParagraphBreak(storyId: string, index: number): TextReceipt;
  insertSlide(index: number, layoutPartPath?: string): SlideReceipt;
  deleteSlide(slideId: string): SlideReceipt;
  moveSlide(slideId: string, toIndex: number): SlideReceipt;
  addTextBox(slideId: string, draft: ShapeDraft): ShapeReceipt;
  removeShape(slideId: string, shapeId: string): ShapeReceipt;
  moveShape(slideId: string, shapeId: string, x: number, y: number): TransformReceipt;
  resizeShape(slideId: string, shapeId: string, width: number, height: number): TransformReceipt;
  canUndo(): boolean;
  canRedo(): boolean;
  undo(): HistoryResult;
  redo(): HistoryResult;
  encodeStateVector(): Uint8Array;
  encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array;
  encodeDiff(remoteStateVector: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): DeckSnapshot;
  onUpdate(
    listener: (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  ): () => void;
  dispose(): void;
}

let initialized = false;
let initialization: Promise<void> | undefined;

export function initWasm(
  input: WasmInitInput = new URL('./generated/pptx_wasm_bg.wasm', import.meta.url)
): Promise<void> {
  if (initialized) return Promise.resolve();
  if (initialization) return initialization;
  initialization = initWasmModule({ module_or_path: input }).then(
    () => {
      initialized = true;
    },
    (error: unknown) => {
      initialization = undefined;
      throw toError(error);
    }
  );
  return initialization;
}

export function isWasmAvailable(): boolean {
  return typeof WebAssembly === 'object';
}

export function wasmVersion(): string {
  requireInitialized();
  return rendererVersion();
}

export function inspectPresentation(bytes: Uint8Array): unknown {
  requireInitialized();
  return call(() => parsePptxJson(bytes));
}

export function openPresentation(
  bytes: Uint8Array,
  options: OpenPresentationOptions = {}
): PresentationHandle {
  requireInitialized();
  const collaborationClientId = options.clientId ?? clientId();
  const doc = construct(() =>
    options.initialUpdate === undefined
      ? PptxDocument.openCollaborative(bytes, collaborationClientId)
      : PptxDocument.openCollaborativeFromUpdate(
          options.initialUpdate.slice(),
          collaborationClientId
        )
  );
  const renderer = construct(() => new PptxRenderer());
  for (const face of options.fonts ?? []) registerFont(renderer, face);
  const listeners = new Map<
    number,
    (update: Uint8Array, origin: CollaborationUpdateOrigin) => void
  >();
  const pendingUpdates: Array<{
    update: Uint8Array;
    origin: CollaborationUpdateOrigin;
  }> = [];
  let nextListenerId = 0;
  let disposed = false;
  let observerInstalled = false;
  let wasmCallDepth = 0;
  let flushingUpdates = false;

  const assertAlive = (): void => {
    if (disposed) throw new Error('presentation handle is disposed');
  };

  const flushUpdates = (): void => {
    if (disposed || flushingUpdates || wasmCallDepth !== 0) return;
    flushingUpdates = true;
    try {
      while (!disposed && pendingUpdates.length > 0) {
        const event = pendingUpdates.shift();
        if (!event) break;
        for (const [id, listener] of [...listeners]) {
          if (disposed) return;
          if (listeners.get(id) !== listener) continue;
          try {
            listener(event.update.slice(), event.origin);
          } catch {}
        }
      }
    } finally {
      flushingUpdates = false;
      if (disposed) pendingUpdates.length = 0;
    }
  };

  const drainWasmUpdates = (): void => {
    if (!observerInstalled || disposed) return;
    while (true) {
      const encoded = doc.drainUpdateEvent();
      if (encoded.byteLength === 0) return;
      const origin = encoded[0];
      if (origin !== 0 && origin !== 1) {
        throw new Error(`pptx wasm returned unknown update origin ${origin}`);
      }
      pendingUpdates.push({
        update: encoded.slice(1),
        origin: origin === 0 ? 'local' : 'remote',
      });
    }
  };

  const wasmCall = <T,>(operation: () => T, drainUpdates = false): T => {
    assertAlive();
    wasmCallDepth += 1;
    try {
      let result: T | undefined;
      let failure: unknown;
      let failed = false;
      try {
        result = operation();
      } catch (error) {
        failure = error;
        failed = true;
      }
      if (drainUpdates) {
        try {
          drainWasmUpdates();
        } catch (error) {
          if (!failed) {
            failure = error;
            failed = true;
          }
        }
      }
      if (failed) throw toError(failure);
      return result as T;
    } finally {
      wasmCallDepth -= 1;
      if (wasmCallDepth === 0) flushUpdates();
    }
  };

  const jsonWasmCall = <T,>(operation: () => string, drainUpdates = false): T =>
    wasmCall(() => JSON.parse(operation()) as T, drainUpdates);

  const ensureUpdateObserver = (): void => {
    if (observerInstalled) return;
    wasmCall(() => doc.startUpdateObservation());
    observerInstalled = true;
  };

  const clearUnusedUpdateObserver = (): void => {
    if (!observerInstalled || listeners.size > 0 || disposed) return;
    pendingUpdates.length = 0;
    wasmCall(() => doc.clearUpdateObservation());
    observerInstalled = false;
  };

  const handle: PresentationHandle = {
    get clientId(): number {
      return wasmCall(() => doc.clientId);
    },
    snapshot(): DeckSnapshot {
      return jsonWasmCall(() => doc.snapshotJson());
    },
    story(storyId: string): StorySnapshot {
      return jsonWasmCall(() => doc.storyJson(JSON.stringify({ storyId })));
    },
    registerFont(face: PptxFontFace): number {
      return wasmCall(() => registerFont(renderer, face));
    },
    layoutSlide(slideIndex: number): SlideDisplayList {
      return jsonWasmCall(() => renderer.layoutSlideJson(doc, slideIndex));
    },
    hitTest(x: number, y: number): HitTestResult | null {
      return jsonWasmCall(() => renderer.hitTestJson(x, y));
    },
    mediaBytes(partPath: string): Uint8Array {
      return wasmCall(() => doc.mediaBytes(partPath).slice());
    },
    insertText(storyId, index, text, style = {}): TextReceipt {
      return jsonWasmCall(
        () => doc.insertTextJson(JSON.stringify({ storyId, index, text, style })),
        true
      );
    },
    deleteText(storyId, start, end): TextReceipt {
      return jsonWasmCall(
        () => doc.deleteTextJson(JSON.stringify({ storyId, start, end })),
        true
      );
    },
    formatText(storyId, start, end, patch): TextReceipt {
      return jsonWasmCall(
        () => doc.formatTextJson(JSON.stringify({ storyId, start, end, patch })),
        true
      );
    },
    insertParagraphBreak(storyId, index): TextReceipt {
      return jsonWasmCall(
        () => doc.insertParagraphBreakJson(JSON.stringify({ storyId, index })),
        true
      );
    },
    insertSlide(index, layoutPartPath): SlideReceipt {
      return jsonWasmCall(
        () =>
          doc.insertSlideJson(JSON.stringify({ index, layoutPartPath: layoutPartPath ?? null })),
        true
      );
    },
    deleteSlide(slideId): SlideReceipt {
      return jsonWasmCall(() => doc.deleteSlideJson(JSON.stringify({ slideId })), true);
    },
    moveSlide(slideId, toIndex): SlideReceipt {
      return jsonWasmCall(
        () => doc.moveSlideJson(JSON.stringify({ slideId, toIndex })),
        true
      );
    },
    addTextBox(slideId, draft): ShapeReceipt {
      return jsonWasmCall(() => doc.addTextBoxJson(JSON.stringify({ slideId, draft })), true);
    },
    removeShape(slideId, shapeId): ShapeReceipt {
      return jsonWasmCall(
        () => doc.removeShapeJson(JSON.stringify({ slideId, shapeId })),
        true
      );
    },
    moveShape(slideId, shapeId, x, y): TransformReceipt {
      return jsonWasmCall(
        () => doc.moveShapeJson(JSON.stringify({ slideId, shapeId, x, y })),
        true
      );
    },
    resizeShape(slideId, shapeId, width, height): TransformReceipt {
      return jsonWasmCall(
        () => doc.resizeShapeJson(JSON.stringify({ slideId, shapeId, width, height })),
        true
      );
    },
    canUndo(): boolean {
      return wasmCall(() => doc.canUndo());
    },
    canRedo(): boolean {
      return wasmCall(() => doc.canRedo());
    },
    undo(): HistoryResult {
      return jsonWasmCall(() => doc.undoJson(), true);
    },
    redo(): HistoryResult {
      return jsonWasmCall(() => doc.redoJson(), true);
    },
    encodeStateVector(): Uint8Array {
      return wasmCall(() => doc.encodeStateVector().slice());
    },
    encodeStateAsUpdate(remoteStateVector?: Uint8Array): Uint8Array {
      return wasmCall(() =>
        remoteStateVector === undefined
          ? doc.encodeStateAsUpdate().slice()
          : doc.encodeDiff(remoteStateVector.slice()).slice()
      );
    },
    encodeDiff(remoteStateVector): Uint8Array {
      return wasmCall(() => doc.encodeDiff(remoteStateVector.slice()).slice());
    },
    applyUpdate(update): DeckSnapshot {
      return jsonWasmCall(() => doc.applyUpdateJson(update.slice()), true);
    },
    onUpdate(listener): () => void {
      assertAlive();
      if (typeof listener !== 'function') throw new TypeError('update listener must be a function');
      const id = nextListenerId++;
      listeners.set(id, listener);
      try {
        ensureUpdateObserver();
      } catch (error) {
        listeners.delete(id);
        throw error;
      }
      let subscribed = true;
      return () => {
        if (!subscribed) return;
        subscribed = false;
        listeners.delete(id);
        clearUnusedUpdateObserver();
      };
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      listeners.clear();
      pendingUpdates.length = 0;
      let disposalError: unknown;
      if (observerInstalled) {
        try {
          doc.clearUpdateObservation();
        } catch (error) {
          disposalError = error;
        }
        observerInstalled = false;
      }
      try {
        renderer.free();
      } catch (error) {
        disposalError ??= error;
      }
      try {
        doc.free();
      } catch (error) {
        disposalError ??= error;
      }
      if (disposalError !== undefined) throw toError(disposalError);
    },
  };
  return handle;
}

function registerFont(renderer: PptxRenderer, face: PptxFontFace): number {
  try {
    return renderer.registerFont(face.family, face.bold ?? false, face.italic ?? false, face.bytes);
  } catch (error) {
    throw toError(error);
  }
}

function requireInitialized(): void {
  if (!initialized) throw new Error('pptx wasm is not initialized; call initWasm() first');
}

function clientId(): number {
  const random = globalThis.crypto;
  if (!random || typeof random.getRandomValues !== 'function') {
    throw new Error('crypto.getRandomValues is required to generate a collaboration client ID');
  }
  const values = new Uint32Array(2);
  let value: number;
  do {
    random.getRandomValues(values);
    value = (values[0] & 0x1fffff) * 0x1_0000_0000 + values[1];
  } while (value === 0);
  return value;
}

function construct<T>(operation: () => T): T {
  try {
    return operation();
  } catch (error) {
    throw toError(error);
  }
}

function jsonCall<T>(operation: () => string): T {
  try {
    return JSON.parse(operation()) as T;
  } catch (error) {
    throw toError(error);
  }
}

function call<T>(operation: () => string): T {
  return jsonCall(operation);
}

function toError(error: unknown): Error {
  if (error instanceof Error) return error;
  return new Error(typeof error === 'string' ? error : String(error));
}
