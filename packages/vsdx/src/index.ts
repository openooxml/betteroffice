export { paintPage, sizeCanvasForPage } from './render/canvas';
export type { CanvasImageResolver, PageCanvasLike, PaintPageOptions } from './render/canvas';
export { initWasm, isWasmAvailable, openDiagram, wasmVersion } from './wasm/loader';
export type { CollaborationResync, DiagramHandle, OpenDiagramOptions, WasmInitInput } from './wasm/loader';
export type * from './types';
