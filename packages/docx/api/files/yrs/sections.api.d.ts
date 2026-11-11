// @betteroffice/docx declaration yrs/sections.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { SectionProperties } from '../types/document';
export type SectionStartKind = 'nextPage' | 'continuous' | 'evenPage' | 'oddPage';
export interface SectionPropertyPatch {
    orientation?: 'portrait' | 'landscape';
    pageWidth?: number;
    pageHeight?: number;
    margins?: Partial<Record<'top' | 'bottom' | 'left' | 'right' | 'header' | 'footer' | 'gutter', number>>;
}
export interface SectionInfo {
    closingParagraphId: string;
    firstParagraph: number;
    lastParagraph: number;
    authored: SectionProperties;
    resolved: SectionProperties;
}
export declare class DocxSectionError extends Error {
    readonly code: string;
    constructor(code: string, message: string);
}
export declare function sectionPropertiesEqual(left: unknown, right: unknown): boolean;
