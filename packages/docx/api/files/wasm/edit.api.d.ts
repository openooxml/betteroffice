// @betteroffice/docx declaration wasm/edit.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Loader for the docx-edit wasm (the Rust yrs editing core + resident layout
 * engine). External-asset pattern — see ./loadWasmAsset.ts for the init
 * contract and URL-geometry invariant.
 *
 * Reached only via the `src/yrs/` facade's dynamic `import()`, so the glue
 * (and the wasm fetch) stay out of every non-editor bundle path. The facade's
 * async `createYrsSession` awaits {@link preloadEditWasm} before constructing
 * a session, which also covers the resident engine worker (it bootstraps
 * through the same facade inside the worker context).
 */
import { EditSession, decodeFontFace as decodePackedFontFace } from './engineEdit';
import { type WasmAsyncInput } from './loadWasmAsset';
export declare const editWasmState: import("./loadWasmAsset").WasmModuleState;
/** Load + instantiate the editing-core wasm (browser path). Idempotent. */
export declare function preloadEditWasm(input?: WasmAsyncInput): Promise<void>;
/** @internal */
export declare function editWasmModule(): Promise<WebAssembly.Module | null>;
/** @internal */
export declare function preloadEditWasmFrom(source: Promise<WebAssembly.Module | null>): Promise<void>;
/**
 * Caps the bytes the editing core in this thread may allocate at once; an
 * allocation past the cap fails as if the memory were full. Undefined removes it.
 */
export declare function setEditWasmHeapLimit(bytes: number | undefined): void;
/**
 * Constructs a raw wasm `EditSession` replica after ensuring the module is
 * initialized. The `src/yrs/` facade wraps this in the typed `YrsSession`
 * surface — nothing else should call it.
 */
export declare function createEditSession(clientId: number): EditSession;
/** Structured export of DOCX bytes: `{ok: true, content}` or `{ok: false, failure}` JSON. */
export declare function exportDocxStructuredJson(bytes: Uint8Array, options: string): string;
/** {@link exportDocxStructuredJson} rendered as Markdown. */
export declare function exportDocxMarkdownJson(bytes: Uint8Array, options: string): string;
/** Markdown rendering of structured content JSON. */
export declare function renderDocxMarkdownJson(content: string, options: string): string;
/** Content controls of DOCX bytes: `{ok: true, content}` or `{ok: false, failure}` JSON. */
export declare function listDocxContentControlsJson(bytes: Uint8Array, options: string): string;
/** The content controls of DOCX bytes matching a query, as {@link listDocxContentControlsJson}. */
export declare function findDocxContentControlsJson(bytes: Uint8Array, query: string, options: string): string;
/** Markdown rendering of a paged export's JSON, with optional page markers. */
export declare function renderDocxMarkdownWithPagesJson(content: string, options: string): string;
export type { EditSession };
export { decodePackedFontFace };
