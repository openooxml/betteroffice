// @betteroffice/docx declaration yrs/residentProperties.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document, SectionProperties, Watermark, HeaderFooter, Relationship } from '../types/document';
import type { YrsHyperlinkAttrs, YrsLoc, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';
import type { DocxCommandSelection } from './residentCommandProtocol';
export interface ResidentPackageMetadata {
    body: Pick<Document['package']['document'], 'sections' | 'finalSectionProperties'>;
    headers?: [string, HeaderFooter][];
    footers?: [string, HeaderFooter][];
    relationships?: [string, Relationship][];
}
export type ResidentPropertyTarget = {
    kind: 'section';
    section: string;
} | {
    kind: 'watermark';
} | {
    kind: 'hyperlink';
    selection: DocxCommandSelection;
} | {
    kind: 'objectSelection';
    selection: DocxCommandSelection;
} | {
    kind: 'object';
    at?: YrsLoc;
    handle?: string;
} | {
    kind: 'tableSelection';
    selection: DocxCommandSelection;
};
export type ResidentPropertySnapshot = {
    kind: 'section';
    section: string;
    properties: SectionProperties;
} | {
    kind: 'watermark';
    watermark: Watermark | null;
} | {
    kind: 'hyperlink';
    range: YrsStoryRange;
    hyperlink: YrsHyperlinkAttrs | null;
} | {
    kind: 'object';
    handle: string;
    properties: Record<string, unknown>;
} | {
    kind: 'table';
    properties: {
        width?: number | null;
        widthType?: string | null;
        justification?: 'left' | 'center' | 'right' | null;
    };
};
export declare function packageMetadata(document: Document): ResidentPackageMetadata;
export declare function withPackageMetadata(document: Document, metadata: ResidentPackageMetadata): Document;
export declare function sectionProperties(document: Document, section: string): SectionProperties | null;
export declare function patchSection(document: Document, section: string, patch: Partial<SectionProperties>): Document | null;
export declare function resolvePropertyImage(engine: ResidentEngineSession, handle: string): YrsLoc | null;
export declare function readResidentProperties(engine: ResidentEngineSession, target: ResidentPropertyTarget): ResidentPropertySnapshot | null;
export declare function watermarkDocument(document: Document, watermark: Watermark | null): Document;
