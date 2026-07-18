import initWasmModule, {
  parsePptxJson,
  PptxDocument,
  PptxRenderer,
  rendererVersion,
} from './generated/pptx_wasm.js';
import type { InitInput } from './generated/pptx_wasm.js';
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
  UpdateEvent,
} from '../types';

export type WasmInitInput = InitInput | Promise<InitInput>;

export interface OpenPresentationOptions {
  clientId?: number;
  fonts?: ReadonlyArray<PptxFontFace>;
}

export interface PresentationHandle {
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
  encodeStateAsUpdate(): Uint8Array;
  encodeDiff(remoteStateVector: Uint8Array): Uint8Array;
  applyUpdate(update: Uint8Array): DeckSnapshot;
  startUpdateObservation(): void;
  drainUpdateEvent(): UpdateEvent | null;
  clearUpdateObservation(): void;
  dispose(): void;
}

let initialized = false;
let initialization: Promise<void> | undefined;
let fallbackClientId = 1;

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
  const doc = construct(() => PptxDocument.openCollaborative(bytes, options.clientId ?? clientId()));
  const renderer = construct(() => new PptxRenderer());
  let disposed = false;
  for (const face of options.fonts ?? []) registerFont(renderer, face);

  const active = (): [PptxDocument, PptxRenderer] => {
    if (disposed) throw new Error('presentation handle is disposed');
    return [doc, renderer];
  };

  return {
    get clientId(): number {
      return active()[0].clientId;
    },
    snapshot(): DeckSnapshot {
      return jsonCall(() => active()[0].snapshotJson());
    },
    story(storyId: string): StorySnapshot {
      return jsonCall(() => active()[0].storyJson(JSON.stringify({ storyId })));
    },
    registerFont(face: PptxFontFace): number {
      return registerFont(active()[1], face);
    },
    layoutSlide(slideIndex: number): SlideDisplayList {
      const [document, slideRenderer] = active();
      return jsonCall(() => slideRenderer.layoutSlideJson(document, slideIndex));
    },
    hitTest(x: number, y: number): HitTestResult | null {
      return jsonCall(() => active()[1].hitTestJson(x, y));
    },
    mediaBytes(partPath: string): Uint8Array {
      return binaryCall(() => active()[0].mediaBytes(partPath));
    },
    insertText(storyId, index, text, style = {}): TextReceipt {
      return jsonCall(() =>
        active()[0].insertTextJson(JSON.stringify({ storyId, index, text, style }))
      );
    },
    deleteText(storyId, start, end): TextReceipt {
      return jsonCall(() => active()[0].deleteTextJson(JSON.stringify({ storyId, start, end })));
    },
    formatText(storyId, start, end, patch): TextReceipt {
      return jsonCall(() =>
        active()[0].formatTextJson(JSON.stringify({ storyId, start, end, patch }))
      );
    },
    insertParagraphBreak(storyId, index): TextReceipt {
      return jsonCall(() =>
        active()[0].insertParagraphBreakJson(JSON.stringify({ storyId, index }))
      );
    },
    insertSlide(index, layoutPartPath): SlideReceipt {
      return jsonCall(() =>
        active()[0].insertSlideJson(
          JSON.stringify({ index, layoutPartPath: layoutPartPath ?? null })
        )
      );
    },
    deleteSlide(slideId): SlideReceipt {
      return jsonCall(() => active()[0].deleteSlideJson(JSON.stringify({ slideId })));
    },
    moveSlide(slideId, toIndex): SlideReceipt {
      return jsonCall(() => active()[0].moveSlideJson(JSON.stringify({ slideId, toIndex })));
    },
    addTextBox(slideId, draft): ShapeReceipt {
      return jsonCall(() => active()[0].addTextBoxJson(JSON.stringify({ slideId, draft })));
    },
    removeShape(slideId, shapeId): ShapeReceipt {
      return jsonCall(() => active()[0].removeShapeJson(JSON.stringify({ slideId, shapeId })));
    },
    moveShape(slideId, shapeId, x, y): TransformReceipt {
      return jsonCall(() =>
        active()[0].moveShapeJson(JSON.stringify({ slideId, shapeId, x, y }))
      );
    },
    resizeShape(slideId, shapeId, width, height): TransformReceipt {
      return jsonCall(() =>
        active()[0].resizeShapeJson(JSON.stringify({ slideId, shapeId, width, height }))
      );
    },
    canUndo(): boolean {
      return active()[0].canUndo();
    },
    canRedo(): boolean {
      return active()[0].canRedo();
    },
    undo(): HistoryResult {
      return jsonCall(() => active()[0].undoJson());
    },
    redo(): HistoryResult {
      return jsonCall(() => active()[0].redoJson());
    },
    encodeStateVector(): Uint8Array {
      return binaryCall(() => active()[0].encodeStateVector());
    },
    encodeStateAsUpdate(): Uint8Array {
      return binaryCall(() => active()[0].encodeStateAsUpdate());
    },
    encodeDiff(remoteStateVector): Uint8Array {
      return binaryCall(() => active()[0].encodeDiff(remoteStateVector));
    },
    applyUpdate(update): DeckSnapshot {
      return jsonCall(() => active()[0].applyUpdateJson(update));
    },
    startUpdateObservation(): void {
      try {
        active()[0].startUpdateObservation();
      } catch (error) {
        throw toError(error);
      }
    },
    drainUpdateEvent(): UpdateEvent | null {
      const encoded = binaryCall(() => active()[0].drainUpdateEvent());
      if (encoded.length === 0) return null;
      return {
        origin: encoded[0] === 0 ? 'local' : 'remote',
        update: encoded.slice(1),
      };
    },
    clearUpdateObservation(): void {
      active()[0].clearUpdateObservation();
    },
    dispose(): void {
      if (disposed) return;
      disposed = true;
      renderer.free();
      doc.free();
    },
  };
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
  if (typeof globalThis.crypto?.getRandomValues === 'function') {
    const values = new Uint32Array(2);
    globalThis.crypto.getRandomValues(values);
    const value = values[0] * 0x100000 + (values[1] & 0xfffff);
    if (value > 0) return value;
  }
  fallbackClientId = (fallbackClientId % 0x7ffffffe) + 1;
  return fallbackClientId;
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

function binaryCall(operation: () => Uint8Array): Uint8Array {
  try {
    return operation();
  } catch (error) {
    throw toError(error);
  }
}

function toError(error: unknown): Error {
  if (error instanceof Error) return error;
  return new Error(typeof error === 'string' ? error : String(error));
}
