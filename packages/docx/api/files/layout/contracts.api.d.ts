// @betteroffice/docx declaration layout/contracts.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document, HeaderFooter, SectionProperties } from '../types/document';
export declare const DEFAULT_PAGE_HEIGHT_PX = 1056;
export declare function resolveHeaderFooter(document: Document | null, section: SectionProperties | null | undefined): {
    header: HeaderFooter | null;
    footer: HeaderFooter | null;
    firstHeader: HeaderFooter | null;
    firstFooter: HeaderFooter | null;
};
