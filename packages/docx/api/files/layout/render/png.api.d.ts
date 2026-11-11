// @betteroffice/docx declaration layout/render/png.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import { PngExportError } from '../../../../../shared/office-session/pngExportError';
import type { BundledFontProvider } from '../measure';
import type { DisplayPage } from './displayList';
/** The rasterizer's largest PNG side and area, in pixels. */
export declare const MAX_PAGE_PNG_SIDE = 16384;
export declare const MAX_PAGE_PNG_PIXELS = 16777216;
export type DocxPngTarget = {
    page: number;
};
/** A face a PNG export may also lay out with, as `openDocx` hosts register them. */
export interface DocxPngFontFace {
    family: string;
    bold?: boolean;
    italic?: boolean;
    bytes: Uint8Array;
}
export interface DocxPngOptions {
    scale?: number;
    signal?: AbortSignal;
    fonts?: ReadonlyArray<DocxPngFontFace>;
    /** The provider the editor showing this document measures with, where it has one. */
    fontProvider?: BundledFontProvider;
}
/** Whether this runtime can paint and encode a PNG on a canvas: one with a 2d context, not only a document. */
export declare function canvasPngAvailable(): boolean;
export declare function canvasUnavailable(): PngExportError;
export declare function pictureUnavailable(): PngExportError;
export declare function fontUnavailable(detail: string): PngExportError;
/** Checks a page export's scale and size the way every engine does, and returns the PNG size. */
export declare function pagePngRequest(page: Pick<DisplayPage, 'width' | 'height'>, options: Pick<DocxPngOptions, 'scale'>): {
    width: number;
    height: number;
    scale: number;
};
export declare function pngScale(value: unknown): number;
/** Checks a target the way the native renderer does. */
export declare function pngPage(target: unknown): number;
