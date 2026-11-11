// @betteroffice/docx declaration docx/rustSaveFacade.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** Typed TypeScript boundary for the Rust package writer. */
import type { Document } from '../types/document';
import type { LegacyParagraphV1, ParagraphPatchV1 } from './paragraphPatch';
export interface RustSaveOptions {
    updateModifiedDate?: boolean;
    modifiedBy?: string;
}
export interface RustSaveDeterminism {
    seed: string;
    now: string;
}
/** Main-document body paragraphs a save replaces by source location, guarded by the part digest. */
export interface RustSourceParagraphs {
    partSha256: string;
    paragraphs: Array<{
        path: number[];
        block: number;
    }>;
}
export interface RustSelectiveSave {
    changedParaIds: Iterable<string>;
    /** Replace only these paragraphs and keep every other part's source bytes. */
    sourceParagraphs?: RustSourceParagraphs;
}
export interface RustSaveResult {
    buffer: ArrayBuffer;
    determinism: RustSaveDeterminism;
    warnings: string[];
}
/**
 * Paragraph IDs a session save applies: IDs for source paragraphs outside
 * the edited stories by part and `w:p` occurrence, and parts written as their
 * source bytes with only these IDs patched in.
 */
export interface RustParagraphIds {
    removedCommentIds?: string[];
    sectionEdits?: {
        onlySections: boolean;
        sha256: string;
        /** `paragraph` is a `w:p` occurrence, or with `bodyParagraphs` the position among the body's own paragraphs. */
        edits: Array<{
            paragraph: number | null;
            bodyParagraphs?: number;
            donor: number | null;
            properties: Record<string, unknown>;
        }>;
    };
    assignments: Array<{
        part: string;
        ordinal: number;
        paraId: string;
    }>;
    patchedParts: Array<{
        part: string;
        paraIds: Array<[number, string]>;
    }>;
    wholeParts?: Array<{
        part: string;
        sha256: string;
        changed: number[];
    }>;
    touchedParts?: Array<{
        part: string;
        touched: number[];
    }>;
    /**
     * Story parts written as their source bytes with only some paragraphs
     * re-serialized: those whose `sourceOrdinal` is in `changed`, or whose
     * written comments, revisions, notes or relationships differ from their
     * source. `paragraphs` lists every `sourceOrdinal` the model holds for the
     * part, `sha256` the source part it addresses, `removed` the source
     * paragraphs it no longer holds, and `inserted`, in writing order, where
     * each paragraph written without a `sourceOrdinal` goes.
     */
    splicedParts?: Array<{
        part: string;
        sha256: string;
        paragraphs: number[];
        changed: number[];
        paragraphPatches?: ParagraphPatchV1[];
        legacyParagraphs?: LegacyParagraphV1[];
        tablePatches?: Array<{
            ordinal: number;
            before: unknown;
            after: unknown;
        }>;
        inserted?: Array<{
            before: number;
        } | {
            after: number;
        }>;
        removed?: number[];
        formatOnly?: number[];
    }>;
}
/** Encode the public Document through the inverse allowlisted package wire. */
export declare function writeDocumentWithRust(document: Document, originalBuffer: ArrayBuffer, options?: RustSaveOptions, selective?: RustSelectiveSave, determinism?: RustSaveDeterminism, paragraphIds?: RustParagraphIds, skipMutations?: boolean, yieldSave?: () => Promise<void>): Promise<RustSaveResult>;
/** @internal */
export declare function assertSafeSaveTree(value: unknown, path: string, active?: WeakSet<object>): void;
