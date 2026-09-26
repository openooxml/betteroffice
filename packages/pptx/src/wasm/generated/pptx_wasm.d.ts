/* tslint:disable */
/* eslint-disable */

export class PptxDocument {
    private constructor();
    free(): void;
    [Symbol.dispose](): void;
    acceptProposalJson(args: string): string;
    addCommentJson(args: string): string;
    addPictureJson(args: string): string;
    addShapeJson(args: string): string;
    addTextBoxJson(args: string): string;
    addTextBoxProfiledJson(args: string): string;
    addUndoBoundary(): void;
    /**
     * Anchors a caret offset so that later edits, undo and remote updates move it.
     */
    anchorCaretJson(args: string): string;
    /**
     * Applies a batch all-or-nothing: `{"ok":true,"baseVersion","version","applied","source",
     * "changedSlides","changedStories","receipts"}`.
     */
    applyEditsJson(request: string): string;
    applyUpdateJson(update: Uint8Array): string;
    bringShapeForwardJson(args: string): string;
    bringShapeToFrontJson(args: string): string;
    canRedo(): boolean;
    canUndo(): boolean;
    clearUpdateObservation(): void;
    commentsJson(): string;
    deleteSlideJson(args: string): string;
    deleteTextJson(args: string): string;
    deleteTextProfiledJson(args: string): string;
    /**
     * The session-scoped version token of the committed deck state.
     */
    documentVersion(): string;
    drainUpdateEvent(): Uint8Array;
    encodeDiff(remote_state_vector: Uint8Array): Uint8Array;
    encodeStateAsUpdate(): Uint8Array;
    encodeStateVector(): Uint8Array;
    /**
     * `{"text","within"?,"limit"?}` -> `{"ok":true,"version","matches","truncated"}`.
     */
    findTextJson(request: string): string;
    formatTextJson(args: string): string;
    insertParagraphBreakJson(args: string): string;
    insertSlideJson(args: string): string;
    insertSlideProfiledJson(args: string): string;
    insertTextJson(args: string): string;
    insertTextProfiledJson(args: string): string;
    listProposalsJson(): string;
    mediaBytes(part_path: string): Uint8Array;
    moveShapeJson(args: string): string;
    moveShapeProfiledJson(args: string): string;
    moveSlideJson(args: string): string;
    static openCollaborative(bytes: Uint8Array, client_id: number): PptxDocument;
    /**
     * `source` is the file the update was seeded from; when it matches the
     * recorded fingerprint the session keeps its part bytes and can save.
     * Any other bytes fall back to the bare update session, whose `saveBytes`
     * fails — joining a room must not depend on carrying the right file.
     */
    static openCollaborativeFromUpdate(update: Uint8Array, client_id: number, source?: Uint8Array | null): PptxDocument;
    previewProposalJson(args: string): string;
    proposeJson(args: string): string;
    /**
     * `{"slideIds"?}` -> `{"ok":true,"version","slides","stories"}`.
     */
    readContentJson(request: string): string;
    redoJson(): string;
    rejectProposalJson(args: string): string;
    removeCommentJson(args: string): string;
    removeShapeJson(args: string): string;
    replyToCommentJson(args: string): string;
    resizeShapeJson(args: string): string;
    /**
     * The current offset of an anchor from `anchorCaretJson`, or `null` once its story is gone.
     */
    resolveCaretAnchorJson(args: string): string;
    /**
     * Serializes the deck back to `.pptx` bytes, edits included.
     */
    saveBytes(): Uint8Array;
    searchTextJson(args: string): string;
    sendShapeBackwardJson(args: string): string;
    sendShapeToBackJson(args: string): string;
    setCommentFlavorJson(args: string): string;
    setCommentPositionJson(args: string): string;
    setCommentStatusJson(args: string): string;
    setParagraphAlignmentJson(args: string): string;
    setShapeAdjustJson(args: string): string;
    setShapeFillJson(args: string): string;
    setShapeRectJson(args: string): string;
    setShapeStrokeJson(args: string): string;
    setSlideNotesJson(args: string): string;
    setUndoCaptureMode(mode: string): void;
    snapshotJson(): string;
    startUpdateObservation(): void;
    storyJson(args: string): string;
    undoCaptureMode(): string;
    undoJson(): string;
    /**
     * `undoJson` timed at its undo, snapshot and serialize boundaries, as
     * `{"receipt": ..., "profile": {"undoMs", "snapshotMs", "serializeMs"}}`.
     */
    undoProfiledJson(): string;
    /**
     * Runs every check of `applyEditsJson`, staging included, without changing anything:
     * `{"ok":true,"baseVersion","wouldApply","previews"}`.
     */
    validateEditsJson(request: string): string;
    static version(): string;
    readonly clientId: number;
}

export class PptxRenderer {
    free(): void;
    [Symbol.dispose](): void;
    hitTestJson(x: number, y: number): string;
    layoutProposalDiffSlideJson(document: PptxDocument, id: string, slide_index: number): string;
    layoutProposalSlideJson(document: PptxDocument, id: string, slide_index: number): string;
    layoutSlideJson(document: PptxDocument, slide_index: number): string;
    /**
     * `layoutSlideJson` with its stages timed, returned as
     * `{"layout": ..., "profile": {"scopeMs", "layoutMs", "serializeMs"}}`.
     */
    layoutSlideProfiledJson(document: PptxDocument, slide_index: number): string;
    constructor();
    registerFallbackFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
    registerFont(family: string, bold: boolean, italic: boolean, bytes: Uint8Array): number;
}

export function compileSlideJson(slide_json: string): string;

export function decodeTiffPng(data: Uint8Array): Uint8Array;

export function parsePptxJson(data: Uint8Array): string;

export function rendererVersion(): string;

export type InitInput = RequestInfo | URL | Response | BufferSource | WebAssembly.Module;

export interface InitOutput {
    readonly memory: WebAssembly.Memory;
    readonly __wbg_pptxrenderer_free: (a: number, b: number) => void;
    readonly compileSlideJson: (a: number, b: number) => [number, number, number, number];
    readonly decodeTiffPng: (a: number, b: number) => [number, number, number, number];
    readonly parsePptxJson: (a: number, b: number) => [number, number, number, number];
    readonly pptxrenderer_hitTestJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_layoutProposalDiffSlideJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly pptxrenderer_layoutProposalSlideJson: (a: number, b: number, c: number, d: number, e: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_layoutSlideProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxrenderer_new: () => number;
    readonly pptxrenderer_registerFallbackFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly pptxrenderer_registerFont: (a: number, b: number, c: number, d: number, e: number, f: number, g: number) => [number, number, number];
    readonly rendererVersion: () => [number, number];
    readonly __wbg_pptxdocument_free: (a: number, b: number) => void;
    readonly pptxdocument_acceptProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addPictureJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addTextBoxJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addTextBoxProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_addUndoBoundary: (a: number) => void;
    readonly pptxdocument_anchorCaretJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_applyEditsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_applyUpdateJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_bringShapeForwardJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_bringShapeToFrontJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_canRedo: (a: number) => number;
    readonly pptxdocument_canUndo: (a: number) => number;
    readonly pptxdocument_clearUpdateObservation: (a: number) => void;
    readonly pptxdocument_clientId: (a: number) => number;
    readonly pptxdocument_commentsJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_deleteSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_deleteTextProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_documentVersion: (a: number) => [number, number];
    readonly pptxdocument_drainUpdateEvent: (a: number) => [number, number];
    readonly pptxdocument_encodeDiff: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_encodeStateAsUpdate: (a: number) => [number, number];
    readonly pptxdocument_encodeStateVector: (a: number) => [number, number];
    readonly pptxdocument_findTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_formatTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertParagraphBreakJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertSlideProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_insertTextProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_listProposalsJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_mediaBytes: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveShapeProfiledJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_moveSlideJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_openCollaborative: (a: number, b: number, c: number) => [number, number, number];
    readonly pptxdocument_openCollaborativeFromUpdate: (a: number, b: number, c: number, d: number, e: number) => [number, number, number];
    readonly pptxdocument_previewProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_proposeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_readContentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_redoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_rejectProposalJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_removeCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_removeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_replyToCommentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_resizeShapeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_resolveCaretAnchorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_saveBytes: (a: number) => [number, number, number, number];
    readonly pptxdocument_searchTextJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sendShapeBackwardJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_sendShapeToBackJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentFlavorJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentPositionJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setCommentStatusJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setParagraphAlignmentJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeAdjustJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeFillJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeRectJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setShapeStrokeJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setSlideNotesJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_setUndoCaptureMode: (a: number, b: number, c: number) => [number, number];
    readonly pptxdocument_snapshotJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_startUpdateObservation: (a: number) => [number, number];
    readonly pptxdocument_storyJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_undoCaptureMode: (a: number) => [number, number];
    readonly pptxdocument_undoJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_undoProfiledJson: (a: number) => [number, number, number, number];
    readonly pptxdocument_validateEditsJson: (a: number, b: number, c: number) => [number, number, number, number];
    readonly pptxdocument_version: () => [number, number];
    readonly __wbindgen_exn_store: (a: number) => void;
    readonly __externref_table_alloc: () => number;
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
