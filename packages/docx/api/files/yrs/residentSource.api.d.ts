// @betteroffice/docx declaration yrs/residentSource.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { Document } from '../types/document';
import type { DocxCommandInitRequest } from './residentCommandProtocol';
export interface DocxSourcePreviewOptions extends Partial<Pick<DocxCommandInitRequest, 'fonts' | 'fontRegistrationSlots' | 'layoutInput' | 'extras' | 'displayWindow' | 'progressiveFonts'>> {
    pages?: number;
}
export type DocxSourceOpenRequest = ({
    document: Document;
    bytes?: never;
} | {
    bytes: Uint8Array;
    document?: never;
}) & {
    preview?: boolean | DocxSourcePreviewOptions;
    /** Opens a display-only cut of the first body blocks of `bytes`; a later whole open replaces it. */
    cut?: {
        blocks: number;
        paragraphBudget?: number;
    };
    mediaTokens?: boolean;
    heapLimitBytes?: number;
    generation?: string;
    digest?: string;
    /**
     * One replica seeds from bytes; joiners pass `initialUpdate`. Two seeds merge by story last-writer-wins.
     * `clientClock` is the highest clock of `clientId` the caller has held. A seed would write that id's clocks again
     * and a joined state behind it would reuse later ones, so either opens under a fresh client id.
     */
    collaboration?: {
        clientId: number;
        initialUpdate?: Uint8Array;
        clientClock?: number;
    };
    /** Yrs state restored over the unseeded package of `bytes`, such as a command checkpoint. */
    state?: Uint8Array;
};
export declare function validateResidentSource(source: DocxSourceOpenRequest): void;
export declare function residentSourceMetadata(document: Document): Document;
/** The SVG `data:` image sources a parsed document names directly. */
export declare function residentSvgSources(document: Document): string[];
export declare function tokenizeResidentDocument(document: Document, sources: Map<string, string>): Document;
