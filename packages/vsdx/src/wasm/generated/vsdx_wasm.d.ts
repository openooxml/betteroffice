/* tslint:disable */
/* eslint-disable */

export class VsdxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    addConnectedShapeJson(args: string): string;
    addConnectorJson(args: string): string;
    addFreeConnectorJson(args: string): string;
    addShapeJson(args: string): string;
    addShapeTreeJson(args: string): string;
    addShapeWithTextJson(args: string): string;
    applyUpdateJson(update: Uint8Array): string;
    canRedo(): boolean;
    canUndo(): boolean;
    clearUpdateObservation(): void;
    deleteShapeJson(args: string): string;
    deleteShapesJson(args: string): string;
    /**
     * Returns `[2]` after overflow; discard queued observations and resync from a state vector.
     */
    drainUpdateEvent(): Uint8Array;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    mediaBytes(part_path: string): Uint8Array;
    moveShapeJson(args: string): string;
    moveShapesJson(args: string): string;
    static openCollaborative(bytes: Uint8Array, client_id: number): VsdxDocument;
    static openCollaborativeFromUpdate(update: Uint8Array, client_id: number): VsdxDocument;
    probeCellWritesJson(args: string): string;
    redoJson(): string;
    reorderPageJson(args: string): string;
    reorderShapeJson(args: string): string;
    resizeLocPin(page_id: string, shape_id: string, width: number, height: number): Float64Array;
    resizeShapeJson(args: string): string;
    save(): Uint8Array;
    setCellFormulaJson(args: string): string;
    setCellFormulasJson(args: string): string;
    setConnectorRouteJson(args: string): string;
    setControlHandleJson(args: string): string;
    setShapeBoundsJson(args: string): string;
    setShapeDataJson(args: string): string;
    setShapeTextJson(args: string): string;
    shapeTextJson(args: string): string;
    snapshotJson(): string;
    startUpdateObservation(): void;
    subtreeGlueJson(args: string): string;
    undoJson(): string;
    static version(): string;
    readonly clientId: number;
}

export class VsdxRenderer {
    free(): void;
    [Symbol.dispose](): void;
    clearLayerVisibility(): void;
    exportPdf(document: VsdxDocument): Uint8Array;
    exportPng(document: VsdxDocument, page_index: number, scale: number): Uint8Array;
    exportSvgJson(document: VsdxDocument): string;
    hitTestJson(x: number, y: number): string;
    layoutPageJson(document: VsdxDocument, page_index: number): string;
    /**
     * Lays every document master out once per materialized package.
     */
    masterPreviewsJson(document: VsdxDocument): string;
    constructor();
    pageLayersJson(document: VsdxDocument, page_index: number): string;
    registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
    setLayerVisible(page_part: string, index: number, visible: boolean): void;
    validateJson(document: VsdxDocument): string;
    validatePageJson(document: VsdxDocument, page_index: number): string;
}

export function parseVsdxJson(data: Uint8Array): string;

export function rendererVersion(): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_vsdxrenderer_free: (a: number, b: number) => void;
    readonly parseVsdxJson: (a: number, b: number) => [number, number, number, number];
    readonly rendererVersion: () => [number, number];
    readonly vsdxrenderer_clearLayerVisibility: (a: number) => void;
    readonly vsdxrenderer_exportPdf: (a: number, b: number) => [number, number, number, number];
    readonly vsdxrenderer_exportPng: (a: number, b: number, c: number, d: number) => [number, number, number, number];
    readonly vsdxrenderer_exportSvgJson: (a: number, b: number) => [number, number, number, number];
    readonly vsdxrenderer_hitTestJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxrenderer_layoutPageJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxrenderer_masterPreviewsJson: (a: number, b: number) => [number, number, number, number];
    readonly vsdxrenderer_new: () => number;
    readonly vsdxrenderer_pageLayersJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxrenderer_registerFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly vsdxrenderer_setLayerVisible: (a: number, b: number, c: number, d: number, e: number) => void;
    readonly vsdxrenderer_validateJson: (a: number, b: number) => [number, number, number, number];
    readonly vsdxrenderer_validatePageJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly __wbg_vsdxdocument_free: (a: number, b: number) => void;
    readonly vsdxdocument_addConnectedShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_addConnectorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_addFreeConnectorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_addShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_addShapeTreeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_addShapeWithTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_canRedo: (a: number) => number;
    readonly vsdxdocument_canUndo: (a: number) => number;
    readonly vsdxdocument_clearUpdateObservation: (a: number) => void;
    readonly vsdxdocument_clientId: (a: number) => number;
    readonly vsdxdocument_deleteShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_deleteShapesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_drainUpdateEvent: (a: number) => [number, number];
    readonly vsdxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly vsdxdocument_encodeStateVector: (a: number) => [number, number];
    readonly vsdxdocument_mediaBytes: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_moveShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_moveShapesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly vsdxdocument_openCollaborativeFromUpdate: (a: number, b: number, c: number) => [number, number, number];
    readonly vsdxdocument_probeCellWritesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly vsdxdocument_reorderPageJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_reorderShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_resizeLocPin: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number, number];
    readonly vsdxdocument_resizeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_save: (a: number) => [number, number, number, number];
    readonly vsdxdocument_setCellFormulaJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_setCellFormulasJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_setConnectorRouteJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_setControlHandleJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_setShapeBoundsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_setShapeDataJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_setShapeTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_shapeTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_snapshotJson: (a: number) => [number, number, number, number];
    readonly vsdxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly vsdxdocument_subtreeGlueJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly vsdxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly vsdxdocument_version: () => [number, number];
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __externref_table_dealloc: (a: number) => void;
    readonly __wbindgen_free: (a: number, b: number, c: number) => void;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
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
