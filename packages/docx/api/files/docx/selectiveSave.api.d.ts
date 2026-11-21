// @betteroffice/docx declaration docx/selectiveSave.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Rust-backed selective DOCX save compatibility wrapper. */
import type { Document } from '../types/document';
export interface SelectiveSaveOptions {
    changedParaIds: Set<string>;
    structuralChange: boolean;
    hasUntrackedChanges: boolean;
}
/**
 * Patch changed paragraphs through the Rust package writer. Returns `null` when
 * selective save is unsafe so callers can fall back to a full repack.
 */
export declare function attemptSelectiveSave(doc: Document, originalBuffer: ArrayBuffer, options: SelectiveSaveOptions): Promise<ArrayBuffer | null>;
