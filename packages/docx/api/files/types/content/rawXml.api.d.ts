// @betteroffice/docx declaration types/content/rawXml.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export interface RawXml {
    type: 'rawXml';
    xml: string;
    /** The source occurrence of its first `w:p`, on documents parsed for an editing session. */
    sourceOrdinal?: number;
}
export declare function isRawXml(value: {
    type: string;
}): value is RawXml;
