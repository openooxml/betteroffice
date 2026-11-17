// @betteroffice/docx declaration yrs/findParagraphs.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/** A paragraph the legacy paragraph finder matched, with the handle the paragraph helpers take. */
export interface DocxParagraphMatch {
    paraId: string;
    match: string;
    before: string;
    after: string;
}
/** @internal */
export interface DocxFindParagraphsOptions {
    caseSensitive?: boolean;
    limit?: number;
}
/** @internal */
export declare function findParagraphs(reader: {
    storyIds(): string[];
    paragraphs(story: string): Array<{
        paraId: string;
        text: string;
    }>;
}, query: string, options?: DocxFindParagraphsOptions): DocxParagraphMatch[];
