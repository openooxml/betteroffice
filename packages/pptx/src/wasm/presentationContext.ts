import type { PptxDocument, PptxRenderer } from './generated/pptx_wasm.js';
import type { PresentationHandle } from './loader';

export const presentationContexts = new WeakMap<PresentationHandle, {
  doc: PptxDocument;
  renderer: PptxRenderer;
  wasmCall<T>(operation: () => T, drainUpdates?: boolean): T;
  jsonWasmCall<T>(operation: () => string, drainUpdates?: boolean): T;
  peer: boolean;
}>();
