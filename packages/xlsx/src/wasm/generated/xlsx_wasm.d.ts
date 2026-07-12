/* tslint:disable */
/* eslint-disable */

/**
 * a workbook handle exposed to js. wraps the pure `Session`.
 */
export class XlsxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    /**
     * accept a proposal, applying it as one agent transaction; returns the edit
     * envelope plus `proposalId`, or a `stale: ...` error when the base moved.
     */
    acceptProposalJson(args: string): string;
    /**
     * apply a raw op list as one user transaction; returns `SheetInfo` json.
     */
    applyOpsJson(transaction_json: string): string;
    /**
     * the editable representation of one cell.
     */
    cellJson(args: string): string;
    /**
     * serialized `DisplayList` for a serialized `Viewport`.
     */
    displayListJson(viewport_json: string): string;
    /**
     * enter one cell edit; returns updated `SheetInfo` json.
     */
    editCellJson(args: string): string;
    /**
     * enter a batch of cell edits as one undo step; returns `SheetInfo` json.
     */
    editCellsJson(args: string): string;
    /**
     * the pending proposals: `{"proposals":[...]}`.
     */
    listProposalsJson(): string;
    /**
     * open a workbook from raw `.xlsx` bytes.
     */
    static open(bytes: Uint8Array): XlsxDocument;
    /**
     * register an agent proposal (preview only; the workbook is untouched);
     * returns the stored `Proposal` json.
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
     * render the current sheet viewport to png bytes. only exported when the
     * `raster` feature is compiled in; the js loader feature-detects it.
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
     * undo the last transaction; returns `{"applied":bool,"sheetInfo":{...}}`.
     */
    undoJson(): string;
    /**
     * crate version string.
     */
    static version(): string;
}

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_xlsxdocument_free: (a: number, b: number) => void;
    readonly xlsxdocument_acceptProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_applyOpsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_cellJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_displayListJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_editCellJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_editCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_listProposalsJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_open: (a: number, b: number) => [number, number, number];
    readonly xlsxdocument_proposeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_rangeCellsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_rejectProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_renderPng: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_renderRangePng: (a: number, b: number, c: number) => [number, number, number, number];
    readonly xlsxdocument_saveBytes: (a: number) => [number, number, number, number];
    readonly xlsxdocument_setActiveSheet: (a: number, b: number) => [number, number];
    readonly xlsxdocument_sheetInfoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly xlsxdocument_version: () => [number, number];
    readonly __wbindgen_externrefs: WebAssembly.Table;
    readonly __wbindgen_malloc: (a: number, b: number) => number;
    readonly __wbindgen_realloc: (a: number, b: number, c: number, d: number) => number;
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
