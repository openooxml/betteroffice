// @betteroffice/docx declaration yrs/saveYrsDocx.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { RepackOptions } from '../docx/rezip';
import type { Document } from '../types/document';
import type { YrsCommentInfo, YrsSession } from './index';
import type { DocxParagraphIdentitySnapshot, DocxParagraphSavePlan, DocxPersistedParagraphAnchor, DocxSessionParagraphAnchor } from './paragraphIdentity';
/** A saved paragraph's session anchor and the persisted anchor that resolves to it once reopened. */
export interface DocxSavedParagraph {
    session: DocxSessionParagraphAnchor;
    persisted: DocxPersistedParagraphAnchor;
}
/** A story part written whole with fewer structural elements than its source, see {@link DocxSavedDocument.structureDropped}. */
export interface DocxDroppedStructure {
    part: string;
    dropped: string[];
    /** Surviving paragraphs with a Word paragraph ID that hold less, up to 200 per part. */
    paragraphs: Array<{
        ordinal: number;
        paraId: string;
        text: string;
        dropped: string[];
    }>;
}
/** A paragraph a save wrote as its source bytes, see {@link DocxSavedDocument.keptUnchanged}. */
export interface DocxKeptParagraph {
    part: string;
    ordinal: number;
    paraId?: string;
    /** The start of the paragraph's text. */
    text: string;
    /** `part` when no safe paragraph splice applied and the whole part kept its source bytes. */
    scope: 'paragraph' | 'part';
    /** The source elements writing it would have dropped. */
    dropped: string[];
}
export interface DocxSavedDocument {
    bytes: Uint8Array;
    /**
     * Every session paragraph the bytes hold with a Word paragraph ID, in
     * document order. Stories that share a part are views of its paragraphs:
     * each view the part holds is listed, and its persisted anchor resolves to
     * the first story's view once reopened.
     */
    paragraphs: DocxSavedParagraph[];
    /**
     * Saved paragraphs whose Word paragraph ID the live session reassigned
     * while the save was written, such as by a duplicate repair after a remote
     * update: their persisted anchor finds them in these bytes, not in the
     * live session.
     */
    conflicts: DocxSavedParagraph[];
    /**
     * Paragraphs written as their source bytes although their run formatting changed, because
     * writing them from the session would have dropped structure the source holds (a hyperlink,
     * smart tag, field, drawing alternate). The session still shows their new formatting; the
     * saved file keeps the old. Absent when there are none.
     */
    keptUnchanged?: {
        structure: number;
        paragraphs: DocxKeptParagraph[];
    };
    /**
     * Story parts this save had to write whole, which hold fewer structural elements than their
     * source, such as after an edit to a table's rows that no paragraph splice can record. The edits
     * may account for them, as a deleted row does; absent when no part lost any.
     */
    structureDropped?: DocxDroppedStructure[];
}
/** @internal */
export interface DocxCommentSaveIdentity {
    id: number;
    paraId?: string;
    durableId?: string;
}
/** @internal */
export interface DocxCommentSaveBaseline {
    bytes: Uint8Array;
    comments: readonly YrsCommentInfo[];
}
export interface DocxSaveOptions extends RepackOptions {
    /** @internal */
    commentIdentities?: ReadonlyMap<string, DocxCommentSaveIdentity>;
    /** @internal */
    commentBaseline?: DocxCommentSaveBaseline;
    /** @internal */
    commentDeletion?: boolean;
}
/** The identities and paragraph ID plan a session save applies. @internal */
export interface DocxSessionSave {
    identities: DocxParagraphIdentitySnapshot;
    plan: DocxParagraphSavePlan;
}
/** Captures a save of `session`, in the task that projects the saved document. @internal */
export declare function captureSessionSave(session: YrsSession, identities?: DocxParagraphIdentitySnapshot): DocxSessionSave;
/**
 * Writes `document`, projected from `session` when `capture` was taken, and
 * records the Word paragraph IDs the bytes hold as saved. `patches` selects
 * the planned parts written as `originalBuffer`'s bytes with IDs patched in,
 * and `splices` the story parts written as its bytes with only the
 * paragraphs that need it re-serialized; each selects none unless
 * `originalBuffer` is the session's source package.
 * @internal
 */
export declare function writeSessionSave(session: YrsSession, document: Document, capture: DocxSessionSave, originalBuffer: ArrayBuffer, options?: DocxSaveOptions, patches?: (part: string) => boolean, skipMutations?: boolean, splices?: (part: string) => boolean, yieldSave?: () => Promise<void>): Promise<DocxSavedDocument>;
/**
 * Saves a session opened from DOCX bytes and reports each saved paragraph's
 * persisted anchor, read from the same projection the bytes are written
 * from and kept only where the written part holds its Word paragraph ID.
 * Parts whose stories are unchanged since the session opened keep
 * their source bytes, paragraph IDs aside. The comments the session holds
 * are saved with their anchors, replies and resolution. A source paragraph without a
 * Word paragraph ID saves without one unless the host called
 * `persistParagraphIds()` first; the IDs a save writes are recorded as
 * saved, so they keep them against unsaved claims from other replicas, once
 * reconciled with the live session.
 */
export declare function saveYrsDocx(session: YrsSession, options?: DocxSaveOptions): Promise<DocxSavedDocument>;
