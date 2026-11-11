// @betteroffice/docx declaration layout/render/canvasBackend.d.ts
// Regenerate: bun scripts/api-surface.ts --write docx
import type { DisplayList, DisplayPage, DisplayPrimitive, GlyphRunPrimitive } from './displayList';
import type { GlyphCache } from './glyphCache';
/**
 * Resolves an image relationship id to a drawable source. Media decode stays
 * in the host (browser decoders, object URLs, ImageBitmap caches) — the
 * backend only replays the resolved source at the primitive's rect. Return
 * null (or resolve to null) to skip the image.
 */
export type ImageResolver = (relId: string) => CanvasImageSource | Promise<CanvasImageSource | null> | null;
export interface DrawPageOptions {
    resolveImage?: ImageResolver;
    /** Glyph cache; unresolved outlines fall back to `fillText`. */
    glyphCache?: GlyphCache;
    /** The CSS family each document font family paints browser text with, where they differ. */
    fontFamilies?: ReadonlyMap<string, string>;
    /** Told when a glyph run's outlines fail and it falls back to `fillText`. */
    onGlyphFallback?: (error: unknown) => void;
}
/**
 * Minimal structural view of a canvas element for DPR sizing. Matches
 * HTMLCanvasElement without requiring one, so hosts and tests can pass
 * lookalikes.
 */
export interface PageCanvasLike {
    width: number;
    height: number;
    style: {
        width: string;
        height: string;
    };
}
export type PageCanvasBuffer = HTMLCanvasElement | OffscreenCanvas;
/**
 * Raster a page or damage band into a detached backing surface. Callers can prepare
 * every damaged page concurrently, then present the finished buffers in one
 * synchronous pass without ever exposing the renderer's initial clear.
 */
export declare function rasterizeDisplayPageToBackBuffer(canvas: PageCanvasBuffer, page: DisplayPage, options?: DrawPageOptions, devicePixelRatio?: number, zoom?: number, 
/** Device-pixel crop; its pixels equal the corresponding full-page raster. */
clip?: {
    x: number;
    y: number;
    width: number;
    height: number;
}): Promise<PageCanvasBuffer>;
/** Atomically replace one visible canvas with a fully rastered page buffer. */
export declare function presentDisplayPageBackBuffer(canvas: PageCanvasBuffer, buffer: PageCanvasBuffer, page: DisplayPage, zoom?: number): void;
/** Zero-copy atomic presentation for a worker-owned transferred canvas. */
export declare function presentOffscreenPageBackBuffer(canvas: OffscreenCanvas, buffer: OffscreenCanvas): void;
/**
 * Release a worker-owned transferred canvas's bitmap and keep the canvas usable.
 * Chromium never shows a canvas's later frames once it is resized to 0x0 while a
 * presented frame is in flight, so it shrinks to one pixel and presents an empty bitmap.
 */
export declare function releaseOffscreenPageCanvas(canvas: OffscreenCanvas): void;
/**
 * Present a page buffer with a caret line composited at present time. The
 * stroke goes through `stage`, so `buffer` keeps its clean raster and the
 * line can later be dropped by re-presenting `buffer` without re-rastering.
 * `caret` is a device-px rect.
 */
export declare function presentOffscreenPageBackBufferWithCaret(canvas: OffscreenCanvas, buffer: OffscreenCanvas, stage: OffscreenCanvas, caret: {
    x: number;
    y: number;
    width: number;
    height: number;
    color: string;
}): void;
/**
 * DPR- and zoom-aware canvas sizing: the bitmap is the page's logical px size
 * times `devicePixelRatio * zoom`, the CSS box is the page size times `zoom`
 * (so the page physically grows with the zoom control, mirroring the DOM
 * painter's `transform: scale`), and the context is scaled by the same factor
 * so all drawing keeps using page-local px coordinates. Rastering at
 * `zoom * DPR` — rather than CSS-scaling a `zoom=1` bitmap — is what keeps text
 * and vector outlines crisp when zoomed: glyph `Path2D` fills re-rasterize at
 * the higher backing resolution instead of being upsampled. `zoom` defaults to
 * 1, so the un-zoomed path is byte-identical to before.
 *
 * Setting canvas.width resets the context state, so call this before drawing
 * (and re-call after any resize/zoom change) — the scale applies to the fresh
 * transform.
 */
export declare function sizeCanvasForPage(canvas: PageCanvasLike, ctx: CanvasRenderingContext2D, page: Pick<DisplayPage, 'width' | 'height'>, devicePixelRatio: number, zoom?: number): void;
/** Raster every display page without requiring a mounted renderer surface. */
export declare function rasterizeDisplayListPages(list: DisplayList, options?: DrawPageOptions, devicePixelRatio?: number): Promise<HTMLCanvasElement[]>;
/** Replays watermark, header/footer, body and note layers. */
export declare function drawDisplayPage(ctx: CanvasRenderingContext2D, page: DisplayPage, options?: DrawPageOptions): Promise<void>;
/** Replays a single primitive; exported for hosts that paint incrementally. */
export declare function drawPrimitive(ctx: CanvasRenderingContext2D, primitive: DisplayPrimitive, options?: DrawPageOptions): Promise<void>;
/** Paints one glyph run synchronously, exactly as page rasterization does. */
export declare function drawGlyphRunPrimitive(ctx: CanvasRenderingContext2D, primitive: GlyphRunPrimitive, options?: DrawPageOptions): void;
/**
 * `font`, a CSS font shorthand or a bare family list, with each family that
 * `families` maps replaced by the mapped one.
 */
export declare function withFontFamilies(font: string, families: ReadonlyMap<string, string>): string;
