/* tslint:disable */
/* eslint-disable */

export class PptxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    addTextBoxJson(args: string): string;
    applyUpdateJson(update: Uint8Array): string;
    canRedo(): boolean;
    canUndo(): boolean;
    clearUpdateObservation(): void;
    deleteSlideJson(args: string): string;
    deleteTextJson(args: string): string;
    drainUpdateEvent(): Uint8Array;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    formatTextJson(args: string): string;
    insertParagraphBreakJson(args: string): string;
    insertSlideJson(args: string): string;
    insertTextJson(args: string): string;
    mediaBytes(part_path: string): Uint8Array;
    moveShapeJson(args: string): string;
    moveSlideJson(args: string): string;
    static openCollaborative(bytes: Uint8Array, client_id: number): PptxDocument;
    static openCollaborativeFromUpdate(update: Uint8Array, client_id: number): PptxDocument;
    redoJson(): string;
    removeShapeJson(args: string): string;
    resizeShapeJson(args: string): string;
    snapshotJson(): string;
    startUpdateObservation(): void;
    storyJson(args: string): string;
    undoJson(): string;
    static version(): string;
    readonly clientId: number;
}

export class PptxRenderer {
    free(): void;
    [Symbol.dispose](): void;
    hitTestJson(x: number, y: number): string;
    layoutSlideJson(document: PptxDocument, slide_index: number): string;
    constructor();
    registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
}

export function compileSlideJson(slide_json: string): string;

export function parsePptxJson(data: Uint8Array): string;

export function rendererVersion(): string;

/**
 * Rezip from a JS object `{ [path]: Uint8Array }` into a DOCX byte array.
 */
export function rezip_docx(entries: any): Uint8Array;

/**
 * Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
 */
export function unzip_docx(data: Uint8Array): any;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_pptxrenderer_free: (a: number, b: number) => void;
    readonly compileSlideJson: (a: number, b: number) => [number, number, number, number];
    readonly parsePptxJson: (a: number, b: number) => [number, number, number, number];
    readonly pptxrenderer_hitTestJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_new: () => number;
    readonly pptxrenderer_registerFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly rendererVersion: () => [number, number];
    readonly __wbg_pptxdocument_free: (a: number, b: number) => void;
    readonly pptxdocument_addTextBoxJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_canRedo: (a: number) => number;
    readonly pptxdocument_canUndo: (a: number) => number;
    readonly pptxdocument_clearUpdateObservation: (a: number) => void;
    readonly pptxdocument_clientId: (a: number) => number;
    readonly pptxdocument_deleteSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_drainUpdateEvent: (a: number) => [number, number];
    readonly pptxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly pptxdocument_encodeStateVector: (a: number) => [number, number];
    readonly pptxdocument_formatTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertParagraphBreakJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_mediaBytes: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly pptxdocument_openCollaborativeFromUpdate: (a: number, b: number, c: number) => [number, number, number];
    readonly pptxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_removeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_resizeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_snapshotJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly pptxdocument_storyJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_version: () => [number, number];
    readonly rezip_docx: (a: any) => [number, number, number, number];
    readonly unzip_docx: (a: number, b: number) => [number, number, number];
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_start: () => void;
}

export type SyncInitInput = BufferSource | WebAssembly.Module;

/**
 * Instantiates the given `module`, which can either be bytes or
 * a precompiled `WebAssembly.Module`.
 *
 * @param {{ module: SyncInitInput }} module - Passing `SyncInitInput` directly is deprecated.
 *
 * @returns {InitOutput}
 */
export function initSync(module: { module: SyncInitInput } | SyncInitInput): InitOutput;

/**
 * If `module_or_path` is {RequestInfo} or {URL}, makes a request and
 * for everything else, calls `WebAssembly.instantiate` directly.
 *
 * @param {{ module_or_path: InitInput | Promise<InitInput> }} module_or_path - Passing `InitInput` directly is deprecated.
 *
 * @returns {Promise<InitOutput>}
 */
export default function __wbg_init (module_or_path?: { module_or_path: InitInput | Promise<InitInput> } | InitInput | Promise<InitInput>): Promise<InitOutput>;
