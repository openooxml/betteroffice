// @betteroffice/docx declaration yrs/residentToolbarMetadata.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document } from '../types/document';
import { type ResolvedStyleOption } from '../utils/stylePreview';
export interface DocxToolbarMetadata {
    styles: ResolvedStyleOption[];
    fontFamilies: string[];
}
export declare function readResidentToolbarMetadata(hostJson: string): DocxToolbarMetadata;
/** The metadata a source open reports, read from a parsed document: its font table and every font its values name. */
export declare function readParsedToolbarMetadata(document: Document): DocxToolbarMetadata;
