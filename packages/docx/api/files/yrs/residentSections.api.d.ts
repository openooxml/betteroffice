// @betteroffice/docx declaration yrs/residentSections.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document, SectionProperties } from '../types/document';
import type { ResidentEngineSession } from './residentEngineSession';
import type { YrsLoc, YrsParagraphAttrs } from './index';
export type DocxSectionMargins = Partial<Pick<SectionProperties, 'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom' | 'headerDistance' | 'footerDistance' | 'gutter'>>;
export interface ResidentSectionSnapshot {
    version: string;
    sections: {
        section: string;
        properties: SectionProperties;
    }[];
    paragraphs: {
        story: string;
        paraId: string;
        tabs: NonNullable<YrsParagraphAttrs['tabs']>;
    }[];
}
export declare function nativeTabStops(value: unknown): NonNullable<YrsParagraphAttrs['tabs']>;
export declare function validSectionMargins(value: unknown): value is DocxSectionMargins;
export declare function residentSectionPropertiesAt(engine: ResidentEngineSession, at: YrsLoc): SectionProperties | null;
export declare function sectionMarkKey(mark: unknown): string;
export declare function readSectionSnapshot(engine: ResidentEngineSession, base: Document, margins: Record<string, DocxSectionMargins>, seeded: ReadonlyMap<string, string>, nativeSections?: boolean): ResidentSectionSnapshot;
/** `document` with each section keyed by the body paragraph that closes it in the freshly seeded session. */
export declare function withSectionIds(document: Document, closing: readonly string[]): Document;
export declare function withSectionNames(document: Document, names: ReadonlyMap<string, string>): Document;
export declare function withSectionSnapshot(document: Document, snapshot: ResidentSectionSnapshot): Document;
