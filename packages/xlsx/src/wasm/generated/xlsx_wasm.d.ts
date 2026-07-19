/* tslint:disable */
/* eslint-disable */

/**
 * a workbook handle exposed to js; wraps the pure `Session`.
 */
export class XlsxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    /**
     * accept a proposal as one agent transaction; returns the edit envelope
     * plus `proposalId`, or a `stale: ...` error when the base moved.
     */
    acceptProposalJson(args: string): string;
    /**
     * apply a raw op list as one user transaction; returns `SheetInfo` json.
     */
    applyOpsJson(transaction_json: string): string;
    applyUpdateJson(update: Uint8Array): string;
    calculationStatusJson(): string;
    /**
     * the editable representation of one cell.
     */
    cellJson(args: string): string;
    /**
     * Stop observation and discard queued events.
     */
    clearUpdateObservation(): void;
    /**
     * serialized `DisplayList` for a serialized `Viewport`.
     */
    displayListJson(viewport_json: string): string;
    /**
     * Poll one event: origin byte (`0` local, `1` remote), then update; empty means none.
     */
    drainUpdateEvent(): Uint8Array;
    /**
     * enter one cell edit; returns updated `SheetInfo` json.
     */
    editCellJson(args: string): string;
    /**
     * enter a batch of cell edits as one undo step; returns `SheetInfo` json.
     */
    editCellsJson(args: string): string;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    /**
     * the pending proposals: `{"proposals":[...]}`.
     */
    listProposalsJson(): string;
    /**
     * open a workbook from raw `.xlsx` bytes.
     */
    static open(bytes: Uint8Array): XlsxDocument;
    /**
     * Open a replica with a positive, safe-integer client ID.
     */
    static openCollaborative(bytes: Uint8Array, client_id: number): XlsxDocument;
    /**
     * register an agent proposal (preview only); returns the stored `Proposal` json.
     */
    proposeJson(args: string): string;
    /**
     * a rectangular block of cells for clipboard copy.
     */
    rangeCellsJson(args: string): string;
    /**
     * redo the last undone transaction; same shape as `undoJson`.
     */
    redoJson(): string;
    /**
     * reject a proposal by id; returns `{"removed":bool}`.
     */
    rejectProposalJson(args: string): string;
    /**
     * render the current sheet viewport to png bytes (raster feature only).
     */
    renderPng(viewport_json: string): Uint8Array;
    /**
     * render an a1 range (default: used range) at an optional scale to png.
     */
    renderRangePng(args: string): Uint8Array;
    /**
     * serialize the current workbook back to `.xlsx` bytes.
     */
    saveBytes(): Uint8Array;
    /**
     * switch the active sheet by index.
     */
    setActiveSheet(index: number): void;
    /**
     * serialized `SheetInfo`: sheet names, active index, content extent.
     */
    sheetInfoJson(): string;
    /**
     * Start queuing origin-prefixed Yrs update events for polling.
     */
    startUpdateObservation(): void;
    /**
     * undo the last transaction; returns `{"applied":bool,"sheetInfo":{...}}`.
     */
    undoJson(): string;
    /**
     * crate version string.
     */
    static version(): string;
    readonly clientId: number;
}

/**
 * Rezip from a JS object `{ [path]: Uint8Array }` into a DOCX byte array.
 */
export function rezip_docx(entries: any): Uint8Array;

export function sanitizeOoxml(data: Uint8Array, expected_format: string): Uint8Array;

/**
 * Unzip a DOCX; returns a JS object `{ [path]: Uint8Array }`.
 */
export function unzip_docx(data: Uint8Array): any;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_xlsxdocument_free: (a: number, b: number) => void;
    readonly xlsxdocument_acceptProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyOpsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_calculationStatusJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_cellJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_clearUpdateObservation: (a: number) => void;
    readonly xlsxdocument_clientId: (a: number) => number;
    readonly xlsxdocument_displayListJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_drainUpdateEvent: (a: number) => [number, number, number, number];
    readonly xlsxdocument_editCellJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_editCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly xlsxdocument_encodeStateVector: (a: number) => [number, number];
    readonly xlsxdocument_listProposalsJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_open: (a: number, b: number) => [number, number, number];
    readonly xlsxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly xlsxdocument_proposeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_rangeCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_rejectProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_renderPng: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_renderRangePng: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_saveBytes: (a: number) => [number, number, number, number];
    readonly xlsxdocument_setActiveSheet: (a: number, b: number) => [number, number];
    readonly xlsxdocument_sheetInfoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly xlsxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_version: () => [number, number];
    readonly rezip_docx: (a: any) => [number, number, number, number];
    readonly sanitizeOoxml: (a: number, b: number, c: number, d: number) => [number, number, number, number];
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
