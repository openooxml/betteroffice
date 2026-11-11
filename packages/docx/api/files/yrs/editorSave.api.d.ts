// @betteroffice/docx declaration yrs/editorSave.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** The editor's save: the session projected with its host's metadata and comments. @internal */
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import type { YrsSession } from './index';
import { type DocxSavedDocument } from './saveYrsDocx';
export { DirtyProjectionStories, dirtyProjectionStory, EditorDirtyStories, proposalProjectionStories, serialWorkerSaves, } from './dirtyProjectionStories';
/** What the editor's earlier saves of a session leave for its next save. @internal */
export interface EditorSaveRecord {
    commentIds?: ReadonlyMap<number, string>;
    /** A save wrote every part; every later save does too. */
    full: boolean;
    /** The bytes the last save wrote from the source package. */
    saved?: ArrayBuffer;
    /** Paragraphs the last save wrote as their source bytes, see `DocxSavedDocument.keptUnchanged`. */
    keptUnchanged?: DocxSavedDocument['keptUnchanged'];
    structureDropped?: DocxSavedDocument['structureDropped'];
}
/** @internal */
export declare function mergeDocxHostMetadata(full: Document, host: Document): Document;
/** Host metadata read by `mergeDocxHostMetadata`. @internal */
export declare function hostSaveMetadata(host: Document): Document;
/**
 * Paragraphs the session's last editor save wrote as their source bytes although their run
 * formatting changed, because writing them would have dropped structure; absent when none were.
 */
export declare function editorSaveKeptUnchanged(session: YrsSession): DocxSavedDocument['keptUnchanged'];
/** Story parts the session's last editor save wrote whole with less structure, see `DocxSavedDocument.structureDropped`. */
export declare function editorSaveStructureDropped(session: YrsSession): DocxSavedDocument['structureDropped'];
/**
 * Saves `projected`, the editor's projection of `session`, with the host's
 * `comments`. `record` holds the earlier saves; the session's own by default.
 * @internal
 */
export declare function saveEditorDocument(session: YrsSession, projected: Document, comments: Comment[], record?: EditorSaveRecord, yieldSave?: () => Promise<void>): Promise<ArrayBuffer>;
