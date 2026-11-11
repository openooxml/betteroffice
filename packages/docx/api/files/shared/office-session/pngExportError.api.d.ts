// @betteroffice/docx declaration shared/office-session/pngExportError.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
export type PngExportCode = 'canvas-unavailable' | 'canvas-too-large' | 'font-unavailable' | 'picture-unavailable';
export declare class PngExportError extends Error {
    readonly code: PngExportCode;
    constructor(code: PngExportCode, message: string);
}
