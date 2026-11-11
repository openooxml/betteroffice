// @betteroffice/docx declaration layout/render/canvasImageResolver.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
/**
 * Image resolver for the canvas replay backend, shared by both adapters.
 *
 * v0 display lists carry the flow-block image `src` as the primitive's
 * relId — for embedded media that is a blob:/data: URL minted by the parser,
 * or a `media:{n}` token naming the package part, whose bytes `media` reads.
 * Shape picture fills resolve through the same gate via
 * `fillPaint.pictureSrc`. Only those sources are decoded; anything else
 * (notably remote http urls from external-mode relationships, or a raw
 * unresolved `rId`) resolves to null so opening a document never triggers a
 * network fetch (the no-zero-click-external-fetch security contract). A token
 * whose part this session cannot read paints a grey box. Decode
 * results are cached per source so repaints reuse decoded images.
 *
 */
import type { ImageResolver } from './canvasBackend';
export interface CanvasImageResolverOptions {
    /** The bytes and media type of a `media:{n}` source, or null. */
    media?: (token: string) => {
        bytes: Uint8Array;
        mimeType: string;
    } | null;
    /** Identifies the package `media` reads; a new value decodes tokens anew. */
    mediaScope?: () => number;
    /** Maximum cached decoded bytes; defaults to 64 MiB. */
    maxCacheBytes?: number;
}
export declare function createCanvasImageResolver(options?: CanvasImageResolverOptions): ImageResolver;
