// @betteroffice/docx declaration yrs/residentSave.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
import { DirtyProjectionStories, type EditorSaveRecord } from './editorSave';
import type { EditSession } from './wasm/index';
/** Earlier saves of the worker's document and the projection the next one starts from. @internal */
export interface ResidentSaveRecord extends EditorSaveRecord {
    base?: Document;
}
/** Saves peer marks when supplied, otherwise worker marks; clears worker marks on success. @internal */
export declare function saveResidentDocument(raw: EditSession, clientId: number, source: Uint8Array, hostJson: string, host: Document | undefined, comments: Comment[], record: ResidentSaveRecord, dirty: DirtyProjectionStories, stories?: readonly string[], yieldSave?: () => Promise<void>): Promise<ArrayBuffer>;
