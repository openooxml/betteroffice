// @betteroffice/docx declaration yrs/residentParts.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document, HeaderFooterType, HeaderFooter, Footnote, Endnote, Relationship } from '../types/document';
import type { ResidentEngineSession } from './residentEngineSession';
import type { RetainedFrame } from '../layout/render/frameDelta';
export interface DocxPartTarget {
    kind: 'header' | 'footer' | 'footnote' | 'endnote';
    section: string;
    page: number;
    variant?: HeaderFooterType;
    noteId?: number;
    expectVersion: string;
}
export interface ResidentPartTarget extends Omit<DocxPartTarget, 'expectVersion'> {
    story: string;
}
export interface ResidentPartGeometry {
    kind: DocxPartTarget['kind'];
    story: string;
    pageIndex: number;
    pageId: string;
    x: number;
    y: number;
    width: number;
    height: number;
}
export interface ResidentPartMetadata {
    body: Pick<Document['package']['document'], 'sections' | 'finalSectionProperties'>;
    headers?: [string, HeaderFooter][];
    footers?: [string, HeaderFooter][];
    footnotes?: Footnote[];
    endnotes?: Endnote[];
    relationships?: [string, Relationship][];
}
export declare function partMetadata(document: Document): ResidentPartMetadata;
export declare function withPartMetadata(document: Document, metadata: ReturnType<typeof partMetadata>): Document;
export declare function validPartTarget(value: unknown): value is DocxPartTarget;
export declare function resolvePart(engine: ResidentEngineSession, target: DocxPartTarget): ResidentPartTarget | null;
export declare function ensurePartDocument(engine: ResidentEngineSession, target: DocxPartTarget): Document | null;
export declare function removePartDocument(engine: ResidentEngineSession, target: DocxPartTarget): Document | null;
export declare function readPartGeometry(engine: ResidentEngineSession, frame: RetainedFrame): ResidentPartGeometry[];
