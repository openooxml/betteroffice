// @betteroffice/docx declaration yrs/formattingError.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export declare class SuggestedFormattingOverContainerError extends Error {
    readonly code = "suggested-formatting-over-container";
    constructor(message: string);
}
export declare function formattingFailure(error: unknown): unknown;
