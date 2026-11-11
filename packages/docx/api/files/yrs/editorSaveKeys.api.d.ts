// @betteroffice/docx declaration yrs/editorSaveKeys.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Comment } from '../types/content';
import type { Document } from '../types/document';
/** @internal */
export declare function editorSaveKeys(document: Document, comments?: readonly Comment[]): {
    metadata: string;
    comments: string;
    commentIds: number[];
};
/** @internal */
export declare function noteSeparatorKeys(document: Document): {
    footnotes: string;
    endnotes: string;
};
