// @betteroffice/docx declaration wasm/preflight.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { type WasmAsyncInput } from './loadWasmAsset';
export declare const preflightWasmState: import("./loadWasmAsset").WasmModuleState;
/** @experimental Load the font preflight wasm on demand. Idempotent. */
export declare function preloadPreflightWasm(input?: WasmAsyncInput): Promise<void>;
/** @experimental */
export interface FontPreflightFace {
    family: string;
    bold: boolean;
    italic: boolean;
    scripts: string[];
}
/** @experimental */
export interface FontPreflight {
    faces: FontPreflightFace[];
    /** False when some font rules could not be resolved. */
    complete: boolean;
}
/** @experimental Font faces needed by DOCX bytes, without opening the editing engine. */
export declare function fontPreflight(bytes: Uint8Array): Promise<FontPreflight>;
