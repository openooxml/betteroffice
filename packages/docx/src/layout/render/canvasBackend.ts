// Canvas replay backend for the display list. Dumb glue by design: every
// geometry and style decision was already made by the layout engine that
// built the DisplayList — this module draws exactly what the primitives say,
// in paint order, and nothing else. Context + display list in, pixels out;
// no DOM queries, no measurement, no fallbacks.

import type {
  DisplayList,
  DisplayPage,
  DisplayPrimitive,
  TextRunPrimitive,
  GlyphRunPrimitive,
  RectPrimitive,
  LinePrimitive,
  ImagePrimitive,
  ShapePrimitive,
  PageBorderPrimitive,
  DisplayBorderStyle,
} from './displayList';
import type { GlyphCache } from './glyphCache';
import { glyphRunRect, textRunRect, type GeoRect } from './displayListGeometry';

/**
 * Resolves an image relationship id to a drawable source. Media decode stays
 * in the host (browser decoders, object URLs, ImageBitmap caches) — the
 * backend only replays the resolved source at the primitive's rect. Return
 * null (or resolve to null) to skip the image.
 */
export type ImageResolver = (
  relId: string
) => CanvasImageSource | Promise<CanvasImageSource | null> | null;

export interface DrawPageOptions {
  resolveImage?: ImageResolver;
  /**
   * Glyph-outline cache for `glyphRun` primitives (phase-2 shaped text). When
   * absent, or when an outline cannot be resolved, the run is painted with
   * `fillText` so text is never invisible. Owned by the host and reused across
   * pages/renders.
   */
  glyphCache?: GlyphCache;
}

/**
 * Minimal structural view of a canvas element for DPR sizing. Matches
 * HTMLCanvasElement without requiring one, so hosts and tests can pass
 * lookalikes.
 */
export interface PageCanvasLike {
  width: number;
  height: number;
  style: { width: string; height: string };
}

export type PageCanvasBuffer = HTMLCanvasElement | OffscreenCanvas;
type PageCanvasContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Raster a complete page into a detached backing surface. Callers can prepare
 * every damaged page concurrently, then present the finished buffers in one
 * synchronous pass without ever exposing the renderer's initial clear.
 */
export async function rasterizeDisplayPageToBackBuffer(
  canvas: PageCanvasBuffer,
  page: DisplayPage,
  options: DrawPageOptions = {},
  devicePixelRatio: number = 1,
  zoom: number = 1
): Promise<PageCanvasBuffer> {
  const ctx = canvas.getContext('2d') as PageCanvasContext | null;
  if (!ctx) throw new Error('Canvas 2D context is unavailable');
  const scale = devicePixelRatio * zoom;
  const width = page.width * scale;
  const height = page.height * scale;
  if (canvas.width !== width) canvas.width = width;
  if (canvas.height !== height) canvas.height = height;
  ctx.resetTransform();
  ctx.scale(scale, scale);
  await drawDisplayPage(ctx as unknown as CanvasRenderingContext2D, page, options);
  return canvas;
}

/** Atomically replace one visible canvas with a fully rastered page buffer. */
export function presentDisplayPageBackBuffer(
  canvas: PageCanvasBuffer,
  buffer: PageCanvasBuffer,
  page: DisplayPage,
  zoom: number = 1
): void {
  if (canvas.width !== buffer.width) canvas.width = buffer.width;
  if (canvas.height !== buffer.height) canvas.height = buffer.height;
  if (typeof HTMLCanvasElement !== 'undefined' && canvas instanceof HTMLCanvasElement) {
    canvas.style.width = `${page.width * zoom}px`;
    canvas.style.height = `${page.height * zoom}px`;
  }
  const ctx = canvas.getContext('2d') as PageCanvasContext | null;
  if (!ctx) throw new Error('Canvas 2D context is unavailable');
  ctx.resetTransform();
  ctx.globalCompositeOperation = 'copy';
  ctx.drawImage(buffer, 0, 0);
  ctx.globalCompositeOperation = 'source-over';
}

/** Zero-copy atomic presentation for a worker-owned transferred canvas. */
export function presentOffscreenPageBackBuffer(
  canvas: OffscreenCanvas,
  buffer: OffscreenCanvas
): void {
  if (canvas.width !== buffer.width) canvas.width = buffer.width;
  if (canvas.height !== buffer.height) canvas.height = buffer.height;
  const ctx = canvas.getContext('bitmaprenderer');
  if (!ctx) throw new Error('Offscreen bitmap renderer is unavailable');
  ctx.transferFromImageBitmap(buffer.transferToImageBitmap());
}

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
export function sizeCanvasForPage(
  canvas: PageCanvasLike,
  ctx: CanvasRenderingContext2D,
  page: Pick<DisplayPage, 'width' | 'height'>,
  devicePixelRatio: number,
  zoom: number = 1
): void {
  const backingScale = devicePixelRatio * zoom;
  canvas.width = page.width * backingScale;
  canvas.height = page.height * backingScale;
  canvas.style.width = `${page.width * zoom}px`;
  canvas.style.height = `${page.height * zoom}px`;
  ctx.scale(backingScale, backingScale);
}

/** Raster every display page without requiring a mounted renderer surface. */
export async function rasterizeDisplayListPages(
  list: DisplayList,
  options: DrawPageOptions = {},
  devicePixelRatio: number = 2
): Promise<HTMLCanvasElement[]> {
  const canvases: HTMLCanvasElement[] = [];
  for (const page of list.pages) {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) continue;
    sizeCanvasForPage(canvas, ctx, page, devicePixelRatio);
    await drawDisplayPage(ctx, page, options);
    canvases.push(canvas);
  }
  return canvases;
}

/**
 * Replays one DisplayPage onto a 2D context, honoring primitive paint order.
 * Header/footer bands replay after the body, mirroring the DOM painter's
 * append order in renderPage (content, then header, then footer); their
 * primitives are already in page coordinates so they go through the same
 * replay path. Async only because image sources may resolve lazily; each
 * image is awaited in sequence so later primitives still paint above it.
 * Every primitive sets the full context state it depends on, so no
 * save/restore per primitive is needed (images use save/restore for their
 * rotation transform only).
 */
export async function drawDisplayPage(
  ctx: CanvasRenderingContext2D,
  page: DisplayPage,
  options: DrawPageOptions = {}
): Promise<void> {
  ctx.clearRect(0, 0, page.width, page.height);
  if (page.background) {
    ctx.fillStyle = page.background;
    ctx.fillRect(0, 0, page.width, page.height);
  }
  for (const border of (page.pageBorders ?? []).filter((p) => p.zOrder === 'back')) {
    drawPageBorder(ctx, border);
  }
  for (const primitive of page.primitives) {
    await drawPrimitive(ctx, primitive, options);
  }
  for (const area of page.noteAreas ?? []) {
    for (const primitive of area.separatorPrimitives ?? []) {
      await drawPrimitive(ctx, primitive, options);
    }
    for (const primitive of area.primitives ?? []) {
      await drawPrimitive(ctx, primitive, options);
    }
  }
  for (const region of [page.header, page.footer]) {
    if (!region) continue;
    for (const primitive of region.primitives) {
      await drawPrimitive(ctx, primitive, options);
    }
  }
  for (const border of (page.pageBorders ?? []).filter((p) => p.zOrder !== 'back')) {
    drawPageBorder(ctx, border);
  }
}

/** Replays a single primitive; exported for hosts that paint incrementally. */
export async function drawPrimitive(
  ctx: CanvasRenderingContext2D,
  primitive: DisplayPrimitive,
  options: DrawPageOptions = {}
): Promise<void> {
  const clip = primitive.clipGroup?.clip;
  if (clip) {
    ctx.save();
    const x = finiteOr(clip.x, 0);
    const y = finiteOr(clip.y, 0);
    const w = Math.max(0, finiteOr(clip.w, 0));
    const h = Math.max(0, finiteOr(clip.h, 0));
    ctx.beginPath();
    ctx.rect(x, y, w, h);
    ctx.clip();
    multiplyGlobalAlpha(ctx, primitive.clipGroup?.opacity);
  }
  try {
    await drawPrimitiveCore(ctx, primitive, options);
  } finally {
    if (clip) ctx.restore();
  }
}

async function drawPrimitiveCore(
  ctx: CanvasRenderingContext2D,
  primitive: DisplayPrimitive,
  options: DrawPageOptions
): Promise<void> {
  switch (primitive.kind) {
    case 'text':
      drawTextRun(ctx, primitive);
      break;
    case 'glyphRun':
      drawGlyphRun(ctx, primitive, options.glyphCache);
      break;
    case 'rect':
      drawRect(ctx, primitive);
      break;
    case 'line':
      drawLine(ctx, primitive);
      break;
    case 'shape':
      await drawShape(ctx, primitive, options.resolveImage);
      break;
    case 'decoration':
      if (primitive.style && primitive.style !== 'solid') {
        drawBorderRecipe(
          ctx,
          primitive.x,
          primitive.y + primitive.h / 2,
          primitive.x + primitive.w,
          primitive.y + primitive.h / 2,
          Math.max(primitive.h, 1),
          primitive.color,
          primitive.style
        );
        break;
      }
      if (primitive.dashed || primitive.dotted) {
        // dashed/dotted rules: stroke a segmented line along the decoration
        // rect's vertical center instead of filling it. Tracked insertions use
        // dashed; hidden text uses dotted.
        const thickness = Math.max(primitive.h, 1);
        const dash = primitive.dotted ? thickness : Math.max(Math.round(thickness * 2), 2);
        const gap = primitive.dotted ? Math.max(Math.round(thickness * 2), 2) : dash;
        ctx.strokeStyle = primitive.color;
        ctx.lineWidth = thickness;
        ctx.setLineDash([dash, gap]);
        const midY = primitive.y + primitive.h / 2;
        ctx.beginPath();
        ctx.moveTo(primitive.x, midY);
        ctx.lineTo(primitive.x + primitive.w, midY);
        ctx.stroke();
        ctx.setLineDash([]);
        break;
      }
      // solid decorations are filled rects per the contract; the display list
      // already sized underline/strike thickness and highlight extents
      ctx.fillStyle = primitive.color;
      ctx.fillRect(primitive.x, primitive.y, primitive.w, primitive.h);
      break;
    case 'image':
      await drawImagePrimitive(ctx, primitive, options.resolveImage);
      break;
    default:
      throw unknownPrimitiveError(primitive);
  }
}

// text runs anchor at the left pen origin with an alphabetic baseline; rtl
// only flips the bidi resolution, never the anchor, so x stays the run's
// left edge exactly as measured upstream
function drawTextRun(ctx: CanvasRenderingContext2D, run: TextRunPrimitive): void {
  let wrapped = beginVisualState(
    ctx,
    textRunRect(run),
    run.hidden && run.opacity === undefined ? 0.4 : run.opacity,
    run.rotationDeg,
    run.horizontalScale
  );
  if (!wrapped && textEffectsNeedIsolation(run)) {
    ctx.save();
    wrapped = true;
  }
  ctx.font = fontWithVariant(run.font, run.smallCaps);
  ctx.fillStyle = run.color;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  ctx.direction = run.rtl ? 'rtl' : 'ltr';
  if ('letterSpacing' in ctx) {
    // reset to 0 when absent so spacing never leaks between runs
    (ctx as CanvasRenderingContext2D & { letterSpacing: string }).letterSpacing =
      `${run.letterSpacing ?? 0}px`;
  }
  if ('wordSpacing' in ctx) {
    // justified lines carry a per-space stretch (Word jc=both/distribute); the
    // builder already baked the same stretch into the primitive width, so the
    // paint and the mirror geometry agree. Reset to 0 when absent so the
    // spacing never leaks into the next run.
    (ctx as CanvasRenderingContext2D & { wordSpacing: string }).wordSpacing =
      `${run.wordSpacing ?? 0}px`;
  }
  const text = run.allCaps ? run.text.toUpperCase() : run.text;
  if (run.leaderGlyphs?.glyph && (run.leaderGlyphs.count ?? 0) > 0) {
    drawLeaderText(ctx, run);
  } else if (!drawModernRunText(ctx, text, run.x, run.baselineY, run, textRunRect(run))) {
    drawCanvasTextWithEffects(ctx, text, run.x, run.baselineY, run);
  }
  drawTextEmphasisMarks(ctx, text, run.x, run.baselineY, run.width, run);
  if (wrapped) ctx.restore();
}

function drawRect(ctx: CanvasRenderingContext2D, rect: RectPrimitive): void {
  // Structural revision bars are regular rect primitives in the display list;
  // replaying them here keeps canvas paint order identical to the Rust output.
  const isolated = rect.opacity !== undefined;
  if (isolated) {
    ctx.save();
    multiplyGlobalAlpha(ctx, rect.opacity);
  }
  ctx.fillStyle = rect.fill;
  ctx.fillRect(rect.x, rect.y, rect.w, rect.h);
  if (isolated) ctx.restore();
}

// warn once per session when glyph outlines are unavailable and a run falls
// back to fillText — a missing wasm export or a bad glyph would otherwise log
// on every run of every page
let glyphFallbackWarned = false;
function warnGlyphFallbackOnce(error: unknown): void {
  if (glyphFallbackWarned) return;
  glyphFallbackWarned = true;
  console.warn(
    '[canvasBackend] glyph outlines unavailable; painting glyph runs with fillText',
    error
  );
}

/**
 * Paints a glyph run as Path2D outlines. Each glyph: translate to its pen
 * origin + baseline, scale by `size`/upem with a NEGATIVE y so the font's
 * y-up outline flips into canvas y-down, then fill the cached path. Empty
 * (whitespace) glyphs carry a null path and are skipped. `fillStyle` is set
 * once per run.
 *
 * Safety net: with no cache, or if any glyph outline fails to resolve (missing
 * wasm export pre-integration, or a bad glyph), the whole run is repainted via
 * `fillText` so text is never invisible. Outlines are resolved up front so a
 * mid-run failure never leaves a half-painted run.
 */
function drawGlyphRun(
  ctx: CanvasRenderingContext2D,
  run: GlyphRunPrimitive,
  cache: GlyphCache | undefined
): void {
  let wrapped = beginVisualState(
    ctx,
    glyphRunRect(run),
    run.hidden && run.opacity === undefined ? 0.4 : run.opacity,
    run.rotationDeg,
    run.horizontalScale
  );
  if (!wrapped && textEffectsNeedIsolation(run)) {
    ctx.save();
    wrapped = true;
  }
  if (cache) {
    try {
      const outlines = run.glyphs.map((g) => cache.get(run.fontId, g.id));
      const effects = run.modernEffects;
      const rect = glyphRunRect(run);
      const paint = modernTextFillStyle(ctx, effects, rect, run.color);
      const paintOutlines = (): void => {
        for (let i = 0; i < run.glyphs.length; i++) {
          const { path, upem } = outlines[i];
          if (!path) continue; // whitespace / empty-contour glyph
          const g = run.glyphs[i];
          const scale = run.size / upem;
          ctx.save();
          ctx.translate(g.x, g.y);
          ctx.scale(scale, -scale);
          const modernOutline = effects?.textOutline;
          if (run.textOutline || modernOutline) {
            ctx.strokeStyle = modernOutline?.color ?? run.color;
            const strokeWidth = modernOutline?.width ?? Math.max(0.5, run.size / 16);
            ctx.lineWidth = Math.max(0.5, strokeWidth) / scale;
            ctx.stroke(path);
            // modern outline strokes AROUND a still-filled glyph unless the
            // fill is explicitly none; classic w:outline stays hollow
            if (modernOutline && !run.textOutline && effects?.textFill?.kind !== 'none') {
              ctx.fill(path);
            }
          } else if (effects?.textFill?.kind !== 'none') {
            ctx.fill(path);
          }
          ctx.restore();
        }
      };
      // reflection first, under the main paint (approximation: flipped repaint
      // below the run box at the start opacity)
      const reflection = effects?.reflection;
      if (reflection) {
        ctx.save();
        const axis = rect.y + rect.h;
        ctx.translate(0, 2 * axis + finiteOr(reflection.distance, 0));
        ctx.scale(1, -1);
        multiplyGlobalAlpha(ctx, clamp(finiteOr(reflection.startOpacity, 0.35), 0, 1));
        ctx.fillStyle = paint;
        paintOutlines();
        ctx.restore();
      }
      ctx.fillStyle = paint;
      applyCanvasTextShadow(ctx, run.textShadow, run.size);
      applyModernTextShadow(ctx, effects);
      paintOutlines();
      drawGlyphEmphasisMarks(ctx, run);
      if (wrapped) ctx.restore();
      return;
    } catch (error) {
      warnGlyphFallbackOnce(error);
      // fall through to the fillText safety net below
    }
  }
  drawGlyphRunFallback(ctx, run);
  if (wrapped) ctx.restore();
}

// browser-text safety net for a glyph run: paints the source text at the first
// glyph's pen origin. `fallbackFont` carries the resolved CSS shorthand of the
// face the run was shaped with, so the fallback keeps the measured
// family/weight/style; pre-contract emissions without it degrade to a generic
// family at the run's size.
function drawGlyphRunFallback(ctx: CanvasRenderingContext2D, run: GlyphRunPrimitive): void {
  const first = run.glyphs[0];
  if (!first) return;
  ctx.font = fontWithVariant(run.fallbackFont ?? `${run.size}px sans-serif`, run.smallCaps);
  ctx.fillStyle = run.color;
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  ctx.direction = run.rtl ? 'rtl' : 'ltr';
  const text = run.allCaps ? run.text.toUpperCase() : run.text;
  if (!drawModernRunText(ctx, text, first.x, first.y, run, glyphRunRect(run))) {
    drawCanvasTextWithEffects(ctx, text, first.x, first.y, run);
  }
  drawGlyphEmphasisMarks(ctx, run);
}

function drawLine(ctx: CanvasRenderingContext2D, line: LinePrimitive): void {
  const isolated = line.opacity !== undefined;
  if (isolated) {
    ctx.save();
    multiplyGlobalAlpha(ctx, line.opacity);
  }
  drawBorderRecipe(
    ctx,
    line.x1,
    line.y1,
    line.x2,
    line.y2,
    line.strokeWidth,
    line.color,
    line.borderStyle ?? 'solid',
    line.secondaryColor,
    line.dash
  );
  if (isolated) ctx.restore();
}

async function drawShape(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive,
  resolveImage: ImageResolver | undefined
): Promise<void> {
  const fillPaint = shape.fillPaint;
  const strokePaint = shape.strokePaint;
  if (!shape.fill && !shape.stroke && !fillPaint && !strokePaint) return;
  ctx.save();
  multiplyGlobalAlpha(ctx, shape.opacity);
  applyShapeTransform(ctx, shape);
  const path = createShapePath(shape);
  if (!path) {
    ctx.beginPath();
    appendShapePath(ctx, shape.geometryPath);
  }
  applyShapeEffects(ctx, shape);
  const fillKind = fillPaint?.kind;
  // pictureSrc is the parser-resolved SAFE embedded source (data:/blob:); the
  // raw relId is kept as a legacy key for hosts whose resolver maps rIds to
  // already-decoded sources. The default resolver only decodes data:/blob:.
  const pictureSource =
    fillKind === 'picture' ? (fillPaint?.pictureSrc ?? fillPaint?.pictureRelId) : undefined;
  if (fillKind === 'picture' && pictureSource && resolveImage) {
    const source = await resolveImage(pictureSource);
    if (source) {
      drawPictureShapeFill(ctx, shape, path, source);
    } else {
      // unresolvable picture: fall back to the legacy solid fill so the shape
      // is not invisible (matches the pre-picture-payload behavior baseline)
      const paint = shapeFillStyle(ctx, shape);
      if (paint) fillShapePath(ctx, path, paint);
    }
  } else if (fillKind === 'pattern') {
    drawPatternShapeFill(ctx, shape, path);
  } else if (fillKind !== 'none') {
    const paint = shapeFillStyle(ctx, shape);
    if (paint) fillShapePath(ctx, path, paint);
  }
  if (shape.stroke || strokePaint) {
    drawShapeStroke(ctx, shape, path);
  }
  ctx.restore();
}

function applyShapeTransform(ctx: CanvasRenderingContext2D, shape: ShapePrimitive): void {
  const transform = shape.transform;
  if (!transform?.rotation && !transform?.flipH && !transform?.flipV) return;
  const cx = shape.x + shape.w / 2;
  const cy = shape.y + shape.h / 2;
  ctx.translate(cx, cy);
  if (transform.rotation) {
    ctx.rotate((transform.rotation * Math.PI) / 180);
  }
  if (transform.flipH || transform.flipV) {
    ctx.scale(transform.flipH ? -1 : 1, transform.flipV ? -1 : 1);
  }
  ctx.translate(-cx, -cy);
}

type ShapePathTarget = Pick<
  Path2D,
  'moveTo' | 'lineTo' | 'quadraticCurveTo' | 'bezierCurveTo' | 'closePath'
>;

function createShapePath(shape: ShapePrimitive): Path2D | null {
  const PathCtor = typeof Path2D === 'undefined' ? undefined : Path2D;
  if (!PathCtor) return null;
  const path = new PathCtor();
  appendShapePath(path, shape.geometryPath);
  return path;
}

function appendShapePath(target: ShapePathTarget, commands: ShapePrimitive['geometryPath']): void {
  for (const cmd of commands) {
    switch (cmd.type) {
      case 'move':
        target.moveTo(cmd.x, cmd.y);
        break;
      case 'line':
        target.lineTo(cmd.x, cmd.y);
        break;
      case 'quad':
        target.quadraticCurveTo(cmd.cpx, cmd.cpy, cmd.x, cmd.y);
        break;
      case 'cubic':
        target.bezierCurveTo(cmd.cp1x, cmd.cp1y, cmd.cp2x, cmd.cp2y, cmd.x, cmd.y);
        break;
      case 'close':
        target.closePath();
        break;
    }
  }
}

function fillShapePath(
  ctx: CanvasRenderingContext2D,
  path: Path2D | null,
  paint: string | CanvasGradient | CanvasPattern
): void {
  ctx.fillStyle = paint;
  if (path) ctx.fill(path);
  else ctx.fill();
}

function shapeFillStyle(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive
): string | CanvasGradient | CanvasPattern | null {
  const paint = shape.fillPaint;
  if (!paint || paint.kind === 'solid' || paint.kind === 'theme') {
    return paint?.color ?? shape.fill ?? null;
  }
  if (paint.kind !== 'gradient') return paint.color ?? shape.fill ?? null;

  const stops = normalizedGradientStops(paint.stops, paint.color ?? shape.fill ?? '#000000');
  let gradient: CanvasGradient;
  if (
    paint.gradientType === 'radial' ||
    paint.gradientType === 'rectangular' ||
    paint.gradientType === 'path'
  ) {
    const cx = shape.x + shape.w / 2;
    const cy = shape.y + shape.h / 2;
    gradient = ctx.createRadialGradient(cx, cy, 0, cx, cy, Math.max(shape.w, shape.h) / 2);
  } else {
    const { x1, y1, x2, y2 } = linearGradientEndpoints(shape, paint.angle ?? 0);
    gradient = ctx.createLinearGradient(x1, y1, x2, y2);
  }
  for (const stop of stops) gradient.addColorStop(stop.position, stop.color);
  return gradient;
}

function normalizedGradientStops(
  raw: NonNullable<ShapePrimitive['fillPaint']>['stops'],
  fallback: string
): Array<{ position: number; color: string }> {
  const stops = (raw ?? [])
    .filter((stop): stop is { position?: number; color?: string } => stop != null)
    .map((stop) => {
      const value = finiteOr(stop.position, 0);
      return {
        position: clamp(value > 1 ? value / 100000 : value, 0, 1),
        color: stop.color ?? fallback,
      };
    })
    .sort((a, b) => a.position - b.position);
  if (stops.length === 0)
    return [
      { position: 0, color: fallback },
      { position: 1, color: fallback },
    ];
  if (stops.length === 1)
    return [
      { position: 0, color: stops[0].color },
      { position: 1, color: stops[0].color },
    ];
  return stops;
}

function linearGradientEndpoints(
  shape: Pick<ShapePrimitive, 'x' | 'y' | 'w' | 'h'>,
  angleDeg: number
): { x1: number; y1: number; x2: number; y2: number } {
  const radians = (angleDeg * Math.PI) / 180;
  const dx = Math.cos(radians);
  const dy = Math.sin(radians);
  const cx = shape.x + shape.w / 2;
  const cy = shape.y + shape.h / 2;
  const reach = (Math.abs(dx) * shape.w) / 2 + (Math.abs(dy) * shape.h) / 2;
  return {
    x1: cx - dx * reach,
    y1: cy - dy * reach,
    x2: cx + dx * reach,
    y2: cy + dy * reach,
  };
}

function clipShapePath(ctx: CanvasRenderingContext2D, path: Path2D | null): void {
  if (path) ctx.clip(path);
  else ctx.clip();
}

// resource cap for tiled picture fills: beyond this many tiles the fill
// degrades to one stretched draw instead of feeding a file-supplied tile size
// into an unbounded loop
const MAX_PICTURE_FILL_TILES = 4096;

function drawPictureShapeFill(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive,
  path: Path2D | null,
  source: CanvasImageSource
): void {
  const fillPaint = shape.fillPaint;
  ctx.save();
  clipShapePath(ctx, path);
  multiplyGlobalAlpha(ctx, fillPaint?.pictureOpacity);
  if (fillPaint?.pictureFillMode === 'tile' && drawTiledPictureFill(ctx, shape, source)) {
    ctx.restore();
    return;
  }
  // stretch: map the srcRect-cropped source onto the fillRect target band of
  // the shape box. drawCroppedImage owns the negative/outset crop clamping.
  const stretch = fillPaint?.pictureStretchRect;
  const left = clamp(finiteOr(stretch?.left, 0), -10, 10);
  const top = clamp(finiteOr(stretch?.top, 0), -10, 10);
  const right = clamp(finiteOr(stretch?.right, 0), -10, 10);
  const bottom = clamp(finiteOr(stretch?.bottom, 0), -10, 10);
  const frame: GeoRect = {
    x: shape.x + left * shape.w,
    y: shape.y + top * shape.h,
    w: shape.w * (1 - left - right),
    h: shape.h * (1 - top - bottom),
  };
  if (frame.w > 0 && frame.h > 0) {
    drawCroppedImage(ctx, source, frame, pictureFillCrop(fillPaint));
  }
  ctx.restore();
}

function pictureFillCrop(
  fillPaint: ShapePrimitive['fillPaint']
): { top: number; right: number; bottom: number; left: number } | undefined {
  const srcRect = fillPaint?.pictureSrcRect;
  if (!srcRect) return undefined;
  return {
    top: finiteOr(srcRect.top, 0),
    right: finiteOr(srcRect.right, 0),
    bottom: finiteOr(srcRect.bottom, 0),
    left: finiteOr(srcRect.left, 0),
  };
}

/**
 * Tiled picture fill (`a:tile`): repeats the source at the authored scale from
 * the alignment-anchored origin, honoring tile offsets and alternate-tile
 * mirroring. Returns false (nothing painted) when the tile grid would exceed
 * the resource cap so the caller falls back to a single stretched draw.
 */
function drawTiledPictureFill(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive,
  source: CanvasImageSource
): boolean {
  const { width: sw, height: sh } = intrinsicSize(source);
  if (sw <= 0 || sh <= 0 || shape.w <= 0 || shape.h <= 0) return false;
  const tile = shape.fillPaint?.pictureTile;
  const tileW = Math.max(1, sw * clamp(finiteOr(tile?.scaleX, 1), 0.01, 100));
  const tileH = Math.max(1, sh * clamp(finiteOr(tile?.scaleY, 1), 0.01, 100));
  const columns = Math.ceil(shape.w / tileW) + 1;
  const rows = Math.ceil(shape.h / tileH) + 1;
  if (columns * rows > MAX_PICTURE_FILL_TILES) return false;

  const alignment = tile?.alignment ?? 'tl';
  let anchorX = shape.x;
  if (alignment === 't' || alignment === 'ctr' || alignment === 'b') {
    anchorX = shape.x + (shape.w - tileW) / 2;
  } else if (alignment === 'tr' || alignment === 'r' || alignment === 'br') {
    anchorX = shape.x + shape.w - tileW;
  }
  let anchorY = shape.y;
  if (alignment === 'l' || alignment === 'ctr' || alignment === 'r') {
    anchorY = shape.y + (shape.h - tileH) / 2;
  } else if (alignment === 'bl' || alignment === 'b' || alignment === 'br') {
    anchorY = shape.y + shape.h - tileH;
  }
  // clamp file-supplied offsets so the tile-origin math stays well-conditioned
  anchorX += clamp(finiteOr(tile?.offsetX, 0), -100000, 100000);
  anchorY += clamp(finiteOr(tile?.offsetY, 0), -100000, 100000);

  // walk the anchor back so the grid covers the whole shape box
  const startCol = Math.ceil((anchorX - shape.x) / tileW);
  const startRow = Math.ceil((anchorY - shape.y) / tileH);
  const flip = tile?.flip ?? 'none';
  const flipX = flip === 'x' || flip === 'xy';
  const flipY = flip === 'y' || flip === 'xy';
  for (let row = -startRow; ; row++) {
    const y = anchorY + row * tileH;
    if (y >= shape.y + shape.h) break;
    for (let col = -startCol; ; col++) {
      const x = anchorX + col * tileW;
      if (x >= shape.x + shape.w) break;
      const mirrorX = flipX && Math.abs(col) % 2 === 1;
      const mirrorY = flipY && Math.abs(row) % 2 === 1;
      if (mirrorX || mirrorY) {
        ctx.save();
        ctx.translate(mirrorX ? 2 * x + tileW : 0, mirrorY ? 2 * y + tileH : 0);
        ctx.scale(mirrorX ? -1 : 1, mirrorY ? -1 : 1);
        ctx.drawImage(source, x, y, tileW, tileH);
        ctx.restore();
      } else {
        ctx.drawImage(source, x, y, tileW, tileH);
      }
    }
  }
  return true;
}

function drawPatternShapeFill(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive,
  path: Path2D | null
): void {
  const paint = shape.fillPaint;
  const foreground = paint?.foregroundColor ?? paint?.color ?? shape.fill ?? '#000000';
  const background = paint?.backgroundColor ?? '#ffffff';
  const preset = paint?.patternPreset ?? 'pct50';
  ctx.save();
  clipShapePath(ctx, path);
  ctx.fillStyle = background;
  ctx.fillRect(shape.x, shape.y, shape.w, shape.h);
  ctx.strokeStyle = foreground;
  ctx.fillStyle = foreground;
  ctx.lineWidth = 1;
  drawPatternMarks(ctx, shape, preset);
  ctx.restore();
}

function drawPatternMarks(
  ctx: CanvasRenderingContext2D,
  shape: Pick<ShapePrimitive, 'x' | 'y' | 'w' | 'h'>,
  preset: string
): void {
  const step = patternStep(preset);
  const xEnd = shape.x + shape.w;
  const yEnd = shape.y + shape.h;
  const horizontal = /horz|horizontal|cross|grid|plaid|weave|trellis/i.test(preset);
  const vertical = /vert|vertical|cross|grid|plaid|weave|trellis/i.test(preset);
  const downDiag = /dnDiag|downDiag|diagCross|trellis|weave/i.test(preset);
  const upDiag = /upDiag|diagCross|trellis|weave|zigZag/i.test(preset);
  const dotted = /pct|dot|sphere|confetti/i.test(preset);

  if (dotted || (!horizontal && !vertical && !downDiag && !upDiag)) {
    const radius = /pct5|pct10|sm/i.test(preset) ? 0.5 : 1;
    for (let y = shape.y; y <= yEnd; y += step) {
      for (let x = shape.x; x <= xEnd; x += step) {
        ctx.beginPath();
        ctx.arc(x, y, radius, 0, Math.PI * 2);
        ctx.fill();
      }
    }
  }
  if (horizontal) {
    for (let y = shape.y; y <= yEnd; y += step) strokeSegment(ctx, shape.x, y, xEnd, y);
  }
  if (vertical) {
    for (let x = shape.x; x <= xEnd; x += step) strokeSegment(ctx, x, shape.y, x, yEnd);
  }
  if (downDiag) {
    for (let offset = -shape.h; offset <= shape.w; offset += step) {
      strokeSegment(ctx, shape.x + offset, shape.y, shape.x + offset + shape.h, yEnd);
    }
  }
  if (upDiag) {
    for (let offset = 0; offset <= shape.w + shape.h; offset += step) {
      strokeSegment(ctx, shape.x + offset, yEnd, shape.x + offset - shape.h, shape.y);
    }
  }
}

function patternStep(preset: string): number {
  if (/pct5|pct10|sm|lt/i.test(preset)) return 8;
  if (/pct75|pct80|pct90|dk/i.test(preset)) return 3;
  return 5;
}

function applyShapeEffects(ctx: CanvasRenderingContext2D, shape: ShapePrimitive): void {
  const effect = shape.effects?.find(
    (item) => item.kind === 'shadow' || item.kind === 'outerShadow' || item.kind === 'glow'
  );
  if (!effect) return;
  ctx.shadowColor = effect.color ?? '#000000';
  ctx.shadowBlur = Math.max(0, effect.blurRadius ?? effect.size ?? 0);
  if (effect.kind === 'glow') return;
  const radians = ((effect.direction ?? 0) * Math.PI) / 180;
  const distance = effect.distance ?? 0;
  ctx.shadowOffsetX = Math.cos(radians) * distance;
  ctx.shadowOffsetY = Math.sin(radians) * distance;
}

function drawShapeStroke(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive,
  path: Path2D | null
): void {
  const paint = shape.strokePaint;
  const stroke = shape.stroke;
  const color = paint?.color ?? stroke?.color ?? '#000000';
  const width = Math.max(0, paint?.width ?? stroke?.width ?? 1);
  if (width === 0) return;
  ctx.strokeStyle = color;
  if (paint?.cap) ctx.lineCap = paint.cap === 'flat' ? 'butt' : paint.cap;
  if (paint?.join) ctx.lineJoin = paint.join;
  if (paint?.miterLimit !== undefined) ctx.miterLimit = Math.max(1, paint.miterLimit);
  const dash = paint?.customDash?.length
    ? paint.customDash.map((value) => Math.max(0, value * width))
    : shapeDashPattern(paint?.dash ?? stroke?.dash, width);
  ctx.lineWidth = width;
  ctx.setLineDash(dash);

  const compound = paint?.compound ?? 'single';
  if (compound === 'single') {
    if (path) ctx.stroke(path);
    else ctx.stroke();
    drawShapeLineEnds(ctx, shape, paint?.headEnd, paint?.tailEnd, color, width);
    return;
  }
  const widths = compoundStrokeWidths(compound, width);
  for (const strokeWidth of widths) {
    ctx.lineWidth = strokeWidth;
    if (path) ctx.stroke(path);
    else ctx.stroke();
  }
  drawShapeLineEnds(ctx, shape, paint?.headEnd, paint?.tailEnd, color, width);
}

function compoundStrokeWidths(
  compound: NonNullable<NonNullable<ShapePrimitive['strokePaint']>['compound']>,
  width: number
): number[] {
  switch (compound) {
    case 'double':
      return [width, Math.max(0.5, width / 3)];
    case 'triple':
      return [width, Math.max(0.5, width * 0.6), Math.max(0.5, width * 0.2)];
    case 'thinThick':
    case 'thickThin':
      return [width, Math.max(0.5, width * 0.4)];
    default:
      return [width];
  }
}

function drawShapeLineEnds(
  ctx: CanvasRenderingContext2D,
  shape: ShapePrimitive,
  head: NonNullable<ShapePrimitive['strokePaint']>['headEnd'] | undefined,
  tail: NonNullable<ShapePrimitive['strokePaint']>['tailEnd'] | undefined,
  color: string,
  width: number
): void {
  const endpoints = shapePathEndpoints(shape.geometryPath);
  if (!endpoints) return;
  if (head?.type && head.type !== 'none') {
    drawLineEnd(ctx, endpoints.start, endpoints.next, head, color, width);
  }
  if (tail?.type && tail.type !== 'none') {
    drawLineEnd(ctx, endpoints.end, endpoints.prev, tail, color, width);
  }
}

interface Point {
  x: number;
  y: number;
}

function shapePathEndpoints(
  commands: ShapePrimitive['geometryPath']
): { start: Point; next: Point; prev: Point; end: Point } | null {
  const points: Point[] = [];
  for (const command of commands) {
    if (command.type === 'move' || command.type === 'line')
      points.push({ x: command.x, y: command.y });
    else if (command.type === 'quad') points.push({ x: command.x, y: command.y });
    else if (command.type === 'cubic') points.push({ x: command.x, y: command.y });
  }
  if (points.length < 2) return null;
  return {
    start: points[0],
    next: points[1],
    prev: points[points.length - 2],
    end: points[points.length - 1],
  };
}

function drawLineEnd(
  ctx: CanvasRenderingContext2D,
  tip: Point,
  toward: Point,
  end: { type?: string; width?: string; length?: string },
  color: string,
  strokeWidth: number
): void {
  const angle = Math.atan2(toward.y - tip.y, toward.x - tip.x);
  const length = strokeWidth * lineEndScale(end.length, 4);
  const halfWidth = strokeWidth * lineEndScale(end.width, 2);
  const baseX = tip.x + Math.cos(angle) * length;
  const baseY = tip.y + Math.sin(angle) * length;
  const normalX = -Math.sin(angle) * halfWidth;
  const normalY = Math.cos(angle) * halfWidth;
  ctx.save();
  ctx.fillStyle = color;
  ctx.strokeStyle = color;
  if (end.type === 'oval') {
    ctx.beginPath();
    ctx.arc(tip.x, tip.y, halfWidth, 0, Math.PI * 2);
    ctx.fill();
  } else if (end.type === 'diamond') {
    ctx.beginPath();
    ctx.moveTo(tip.x, tip.y);
    ctx.lineTo(baseX + normalX, baseY + normalY);
    ctx.lineTo(tip.x + Math.cos(angle) * length * 2, tip.y + Math.sin(angle) * length * 2);
    ctx.lineTo(baseX - normalX, baseY - normalY);
    ctx.closePath();
    ctx.fill();
  } else {
    ctx.beginPath();
    ctx.moveTo(tip.x, tip.y);
    ctx.lineTo(baseX + normalX, baseY + normalY);
    ctx.lineTo(baseX - normalX, baseY - normalY);
    ctx.closePath();
    if (end.type === 'arrow') ctx.stroke();
    else ctx.fill();
  }
  ctx.restore();
}

function lineEndScale(value: string | undefined, base: number): number {
  if (value === 'sm') return base * 0.75;
  if (value === 'lg') return base * 1.5;
  return base;
}

function shapeDashPattern(dash: string | undefined, width: number): number[] {
  switch (dash) {
    case 'dot':
    case 'dotted':
    case 'sysDot':
      return [Math.max(1, width), Math.max(2, width * 2)];
    case 'dash':
    case 'dashed':
    case 'dashSmallGap':
    case 'sysDash':
      return [Math.max(2, width * 3), Math.max(2, width * 2)];
    case 'lgDash':
    case 'dashLong':
    case 'dashLongHeavy':
      return [Math.max(4, width * 6), Math.max(2, width * 2)];
    case 'dashDot':
    case 'lgDashDot':
    case 'sysDashDot':
    case 'dashDotHeavy':
      return [Math.max(2, width * 3), Math.max(2, width * 2), width, Math.max(2, width * 2)];
    case 'dashDotDot':
    case 'lgDashDotDot':
    case 'sysDashDotDot':
    case 'dashDotDotHeavy':
      return [
        Math.max(2, width * 3),
        Math.max(2, width * 2),
        width,
        Math.max(2, width * 2),
        width,
        Math.max(2, width * 2),
      ];
    default:
      return [];
  }
}

function drawBorderRecipe(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  width: number,
  color: string,
  style: DisplayBorderStyle,
  secondaryColor?: string,
  dashOverride?: number[]
): void {
  const safeWidth = Math.max(width, 0.5);
  if (style === 'wave' || style === 'doubleWave') {
    drawWaveBorder(ctx, x1, y1, x2, y2, safeWidth, color, style === 'doubleWave');
    return;
  }
  if (style === 'double' || style === 'triple' || style === 'thinThick' || style === 'thickThin') {
    drawCompoundBorder(ctx, x1, y1, x2, y2, safeWidth, color, style);
    return;
  }
  if (style === 'groove' || style === 'ridge' || style === 'inset' || style === 'outset') {
    drawThreeDBorder(ctx, x1, y1, x2, y2, safeWidth, color, secondaryColor, style);
    return;
  }
  ctx.strokeStyle = color;
  ctx.lineWidth = safeWidth;
  ctx.setLineDash(dashOverride ?? borderStyleDash(style, safeWidth));
  strokeSegment(ctx, x1, y1, x2, y2);
}

function borderStyleDash(style: DisplayBorderStyle, width: number): number[] {
  switch (style) {
    case 'dotted':
      return [Math.max(1, width), Math.max(1.5, width * 1.5)];
    case 'dashed':
      return [Math.max(2, width * 4), Math.max(2, width * 2)];
    case 'dashDot':
      return [width * 4, width * 1.5, width, width * 1.5];
    case 'dashDotDot':
      return [width * 4, width * 1.5, width, width * 1.5, width, width * 1.5];
    default:
      return [];
  }
}

function segmentNormal(x1: number, y1: number, x2: number, y2: number): Point {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const length = Math.hypot(dx, dy) || 1;
  return { x: -dy / length, y: dx / length };
}

function drawCompoundBorder(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  width: number,
  color: string,
  style: 'double' | 'triple' | 'thinThick' | 'thickThin'
): void {
  const normal = segmentNormal(x1, y1, x2, y2);
  const strokes =
    style === 'triple'
      ? [
          { offset: -width, width: Math.max(0.5, width / 3) },
          { offset: 0, width: Math.max(0.5, width / 3) },
          { offset: width, width: Math.max(0.5, width / 3) },
        ]
      : style === 'thinThick'
        ? [
            { offset: -width * 0.45, width: Math.max(0.5, width * 0.25) },
            { offset: width * 0.3, width: Math.max(0.75, width * 0.55) },
          ]
        : style === 'thickThin'
          ? [
              { offset: -width * 0.3, width: Math.max(0.75, width * 0.55) },
              { offset: width * 0.45, width: Math.max(0.5, width * 0.25) },
            ]
          : [
              { offset: -width / 2, width: Math.max(0.5, width / 3) },
              { offset: width / 2, width: Math.max(0.5, width / 3) },
            ];
  ctx.strokeStyle = color;
  ctx.setLineDash([]);
  for (const stroke of strokes) {
    ctx.lineWidth = stroke.width;
    const ox = normal.x * stroke.offset;
    const oy = normal.y * stroke.offset;
    strokeSegment(ctx, x1 + ox, y1 + oy, x2 + ox, y2 + oy);
  }
}

function drawWaveBorder(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  width: number,
  color: string,
  double: boolean
): void {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const length = Math.hypot(dx, dy);
  if (length === 0) return;
  const tangent = { x: dx / length, y: dy / length };
  const normal = { x: -tangent.y, y: tangent.x };
  const wavelength = Math.max(4, width * 4);
  const amplitude = Math.max(1, width);
  const offsets = double ? [-amplitude, amplitude] : [0];
  ctx.strokeStyle = color;
  ctx.lineWidth = Math.max(0.5, width / (double ? 2 : 1));
  ctx.setLineDash([]);
  for (const lane of offsets) {
    ctx.beginPath();
    ctx.moveTo(x1 + normal.x * lane, y1 + normal.y * lane);
    let distance = 0;
    let sign = 1;
    while (distance < length) {
      const next = Math.min(length, distance + wavelength / 2);
      const mid = (distance + next) / 2;
      ctx.quadraticCurveTo(
        x1 + tangent.x * mid + normal.x * (lane + amplitude * sign),
        y1 + tangent.y * mid + normal.y * (lane + amplitude * sign),
        x1 + tangent.x * next + normal.x * lane,
        y1 + tangent.y * next + normal.y * lane
      );
      distance = next;
      sign *= -1;
    }
    ctx.stroke();
  }
}

function drawThreeDBorder(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  width: number,
  color: string,
  secondaryColor: string | undefined,
  style: 'groove' | 'ridge' | 'inset' | 'outset'
): void {
  const normal = segmentNormal(x1, y1, x2, y2);
  const light = secondaryColor ?? mixHexColor(color, '#ffffff', 0.55);
  const dark = mixHexColor(color, '#000000', 0.45);
  const reverse = style === 'ridge' || style === 'outset';
  const first = reverse ? light : dark;
  const second = reverse ? dark : light;
  const offset = width / 4;
  ctx.setLineDash([]);
  ctx.lineWidth = Math.max(0.5, width / 2);
  ctx.strokeStyle = first;
  strokeSegment(
    ctx,
    x1 - normal.x * offset,
    y1 - normal.y * offset,
    x2 - normal.x * offset,
    y2 - normal.y * offset
  );
  ctx.strokeStyle = second;
  strokeSegment(
    ctx,
    x1 + normal.x * offset,
    y1 + normal.y * offset,
    x2 + normal.x * offset,
    y2 + normal.y * offset
  );
}

function drawPageBorder(ctx: CanvasRenderingContext2D, border: PageBorderPrimitive): void {
  const left = border.x;
  const top = border.y;
  const right = border.x + border.w;
  const bottom = border.y + border.h;
  if (border.top) drawBorderSegment(ctx, left, top, right, top, border.top, 'horizontal');
  if (border.right) drawBorderSegment(ctx, right, top, right, bottom, border.right, 'vertical');
  if (border.bottom)
    drawBorderSegment(ctx, left, bottom, right, bottom, border.bottom, 'horizontal');
  if (border.left) drawBorderSegment(ctx, left, top, left, bottom, border.left, 'vertical');
}

function drawBorderSegment(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  side: NonNullable<PageBorderPrimitive['top']>,
  _axis: 'horizontal' | 'vertical'
): void {
  drawBorderRecipe(ctx, x1, y1, x2, y2, side.width, side.color, side.style);
  ctx.setLineDash([]);
}

function strokeSegment(
  ctx: CanvasRenderingContext2D,
  x1: number,
  y1: number,
  x2: number,
  y2: number
): void {
  ctx.beginPath();
  ctx.moveTo(x1, y1);
  ctx.lineTo(x2, y2);
  ctx.stroke();
}

async function drawImagePrimitive(
  ctx: CanvasRenderingContext2D,
  image: ImagePrimitive,
  resolveImage: ImageResolver | undefined
): Promise<void> {
  if (!resolveImage) return;
  const source = await resolveImage(image.relId);
  if (!source) return;
  const frame = imageContentFrame(image);
  ctx.save();
  multiplyGlobalAlpha(ctx, image.opacity);
  const filter = imageFilter(image);
  if (filter && 'filter' in ctx) {
    (ctx as CanvasRenderingContext2D & { filter: string }).filter = filter;
  }
  if (image.rotationDeg || image.flipH || image.flipV) {
    const cx = frame.x + frame.w / 2;
    const cy = frame.y + frame.h / 2;
    ctx.translate(cx, cy);
    if (image.rotationDeg) ctx.rotate((image.rotationDeg * Math.PI) / 180);
    if (image.flipH || image.flipV) ctx.scale(image.flipH ? -1 : 1, image.flipV ? -1 : 1);
    ctx.translate(-cx, -cy);
  }
  drawCroppedImage(ctx, source, frame, image.crop);
  drawImageBorder(ctx, image, frame);
  if (image.revision) {
    ctx.strokeStyle = image.revision.kind === 'ins' ? 'rgb(46, 125, 50)' : 'rgb(198, 40, 40)';
    ctx.lineWidth = 2;
    ctx.setLineDash([]);
    ctx.strokeRect(frame.x, frame.y, frame.w, frame.h);
    if (image.revision.kind === 'del') {
      ctx.beginPath();
      ctx.moveTo(frame.x, frame.y + frame.h);
      ctx.lineTo(frame.x + frame.w, frame.y);
      ctx.stroke();
    }
  }
  ctx.restore();
}

function imageContentFrame(image: ImagePrimitive): GeoRect {
  return {
    x: finiteOr(image.contentFrame?.x, image.x),
    y: finiteOr(image.contentFrame?.y, image.y),
    w: Math.max(0, finiteOr(image.contentFrame?.w, image.w)),
    h: Math.max(0, finiteOr(image.contentFrame?.h, image.h)),
  };
}

function drawCroppedImage(
  ctx: CanvasRenderingContext2D,
  source: CanvasImageSource,
  frame: GeoRect,
  crop: ImagePrimitive['crop']
): void {
  if (!crop) {
    ctx.drawImage(source, frame.x, frame.y, frame.w, frame.h);
    return;
  }
  const { width: sw, height: sh } = intrinsicSize(source);
  if (sw <= 0 || sh <= 0 || frame.w <= 0 || frame.h <= 0) return;
  const requested = {
    x: crop.left * sw,
    y: crop.top * sh,
    w: sw * (1 - crop.left - crop.right),
    h: sh * (1 - crop.top - crop.bottom),
  };
  if (requested.w <= 0 || requested.h <= 0) return;
  if (
    requested.x >= 0 &&
    requested.y >= 0 &&
    requested.x + requested.w <= sw &&
    requested.y + requested.h <= sh
  ) {
    ctx.drawImage(
      source,
      requested.x,
      requested.y,
      requested.w,
      requested.h,
      frame.x,
      frame.y,
      frame.w,
      frame.h
    );
    return;
  }
  const clipped = {
    x: clamp(requested.x, 0, sw),
    y: clamp(requested.y, 0, sh),
    right: clamp(requested.x + requested.w, 0, sw),
    bottom: clamp(requested.y + requested.h, 0, sh),
  };
  const clippedW = clipped.right - clipped.x;
  const clippedH = clipped.bottom - clipped.y;
  if (clippedW <= 0 || clippedH <= 0) return;
  const destination = {
    x: frame.x + ((clipped.x - requested.x) / requested.w) * frame.w,
    y: frame.y + ((clipped.y - requested.y) / requested.h) * frame.h,
    w: (clippedW / requested.w) * frame.w,
    h: (clippedH / requested.h) * frame.h,
  };
  ctx.drawImage(
    source,
    clipped.x,
    clipped.y,
    clippedW,
    clippedH,
    destination.x,
    destination.y,
    destination.w,
    destination.h
  );
}

function imageFilter(image: ImagePrimitive): string | null {
  const pieces: string[] = [];
  const legacy = sanitizeCanvasFilter(image.filter);
  if (legacy) pieces.push(legacy);
  for (const effect of image.effects ?? []) {
    const amount = finiteOr(effect.amount, 1);
    switch (effect.kind) {
      case 'brightness':
      case 'contrast':
      case 'saturate':
        pieces.push(`${effect.kind}(${Math.max(0, amount)})`);
        break;
      case 'saturation':
        pieces.push(`saturate(${Math.max(0, amount)})`);
        break;
      case 'grayscale':
        pieces.push(`grayscale(${clamp(amount, 0, 1)})`);
        break;
      case 'biLevel':
        pieces.push(`grayscale(1) contrast(${Math.max(1, amount || 4)})`);
        break;
      case 'blur':
        pieces.push(`blur(${clamp(amount, 0, 128)}px)`);
        break;
    }
  }
  return pieces.length > 0 ? pieces.join(' ') : null;
}

function sanitizeCanvasFilter(filter: string | undefined): string | null {
  if (!filter) return null;
  const allowed =
    /^(?:\s*(?:blur\(\d+(?:\.\d+)?px\)|brightness\(\d+(?:\.\d+)?\)|contrast\(\d+(?:\.\d+)?\)|grayscale\(\d+(?:\.\d+)?%?\)|hue-rotate\(-?\d+(?:\.\d+)?deg\)|invert\(\d+(?:\.\d+)?%?\)|opacity\(\d+(?:\.\d+)?%?\)|saturate\(\d+(?:\.\d+)?\)|sepia\(\d+(?:\.\d+)?%?\))\s*)+$/;
  return filter.length <= 512 && allowed.test(filter) ? filter.trim() : null;
}

function drawImageBorder(
  ctx: CanvasRenderingContext2D,
  image: ImagePrimitive,
  frame: GeoRect
): void {
  const border = image.border;
  const width = finiteOr(border?.width, 0);
  if (!border || width <= 0) return;
  ctx.strokeStyle = border.color ?? '#000000';
  ctx.lineWidth = width;
  ctx.setLineDash(border.dash ?? shapeDashPattern(border.style, width));
  ctx.strokeRect(frame.x, frame.y, frame.w, frame.h);
}

function beginVisualState(
  ctx: CanvasRenderingContext2D,
  rect: GeoRect,
  opacity?: number,
  rotationDeg?: number,
  horizontalScale?: number
): boolean {
  if (opacity === undefined && !rotationDeg && !hasHorizontalScale(horizontalScale)) return false;
  ctx.save();
  beginPrimitiveVisualTransform(ctx, rect, opacity, rotationDeg, horizontalScale);
  return true;
}

function beginPrimitiveVisualTransform(
  ctx: CanvasRenderingContext2D,
  rect: GeoRect,
  opacity?: number,
  rotationDeg?: number,
  horizontalScale?: number
): void {
  multiplyGlobalAlpha(ctx, opacity);
  if (rotationDeg) {
    const cx = rect.x + rect.w / 2;
    const cy = rect.y + rect.h / 2;
    ctx.translate(cx, cy);
    ctx.rotate((rotationDeg * Math.PI) / 180);
    ctx.translate(-cx, -cy);
  }
  if (hasHorizontalScale(horizontalScale)) {
    const scaleX = horizontalScale / 100;
    const cy = rect.y + rect.h / 2;
    ctx.translate(rect.x, cy);
    ctx.scale(scaleX, 1);
    ctx.translate(-rect.x, -cy);
  }
}

function drawLeaderText(ctx: CanvasRenderingContext2D, run: TextRunPrimitive): void {
  const leader = run.leaderGlyphs;
  if (!leader?.glyph) return;
  const count = Math.max(0, Math.floor(leader.count ?? 0));
  const advance = finiteOr(leader.advance, 0);
  if (count === 0 || advance <= 0) return;
  const x = finiteOr(leader.x, run.x);
  const baselineY = finiteOr(leader.baselineY, run.baselineY);
  ctx.font = leader.font ?? run.font;
  ctx.fillStyle = leader.color ?? run.color;
  for (let i = 0; i < count; i++) {
    const visualIndex = leader.rtl ? count - 1 - i : i;
    ctx.fillText(leader.glyph, x + visualIndex * advance, baselineY);
  }
}

function hasHorizontalScale(horizontalScale: number | undefined): horizontalScale is number {
  return horizontalScale !== undefined && horizontalScale !== 100;
}

function fontWithVariant(font: string, smallCaps: boolean | undefined): string {
  if (!smallCaps || /\bsmall-caps\b/.test(font)) return font;
  return font
    .replace(/^(italic|oblique)\s+/, '$1 small-caps ')
    .replace(/^((?!italic|oblique).*)$/, 'small-caps $1');
}

type RunModernEffects = NonNullable<TextRunPrimitive['modernEffects']>;

/**
 * Resolve the modern `textFill` paint for a run: solid color, linear gradient
 * across the run's geometry, or the plain run color when no override applies.
 */
function modernTextFillStyle(
  ctx: CanvasRenderingContext2D,
  effects: RunModernEffects | undefined,
  rect: GeoRect,
  fallback: string
): string | CanvasGradient {
  const fill = effects?.textFill;
  if (!fill) return fallback;
  if (fill.kind === 'gradient') {
    const stops = normalizedGradientStops(fill.stops, fallback);
    // DrawingML angle 0 points right and w14 defaults read top-to-bottom, so
    // fall back to 90deg when no angle is authored
    const { x1, y1, x2, y2 } = linearGradientEndpoints(
      { x: rect.x, y: rect.y, w: Math.max(rect.w, 1), h: Math.max(rect.h, 1) },
      finiteOr(fill.angle, 90)
    );
    const gradient = ctx.createLinearGradient(x1, y1, x2, y2);
    for (const stop of stops) gradient.addColorStop(stop.position, stop.color);
    return gradient;
  }
  if (fill.kind === 'solid' && fill.color) return fill.color;
  return fallback;
}

/** Replay modern glow/shadow as canvas shadow state (glow wins when both). */
function applyModernTextShadow(
  ctx: CanvasRenderingContext2D,
  effects: RunModernEffects | undefined
): void {
  if (!effects) return;
  if (effects.glow) {
    ctx.shadowColor = effects.glow.color ?? 'rgba(255, 192, 0, 0.75)';
    ctx.shadowBlur = Math.max(0, finiteOr(effects.glow.radius, 4));
    ctx.shadowOffsetX = 0;
    ctx.shadowOffsetY = 0;
    return;
  }
  const shadow = effects.shadow;
  if (!shadow) return;
  ctx.shadowColor = shadow.color ?? 'rgba(0, 0, 0, 0.4)';
  ctx.shadowBlur = Math.max(0, finiteOr(shadow.blurRadius, 0));
  const radians = (finiteOr(shadow.direction, 0) * Math.PI) / 180;
  const distance = finiteOr(shadow.distance, 0);
  ctx.shadowOffsetX = Math.cos(radians) * distance;
  ctx.shadowOffsetY = Math.sin(radians) * distance;
}

/**
 * Paint a run that carries modern w14 effects via fillText: reflection
 * under-pass, glow/shadow state, textFill paint override, and the modern
 * compound outline. Returns false when the run has no modern payload so the
 * caller keeps the classic effect path (byte-identical to before).
 */
function drawModernRunText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  baselineY: number,
  run: Pick<TextRunPrimitive | GlyphRunPrimitive, 'color' | 'modernEffects'>,
  rect: GeoRect
): boolean {
  const effects = run.modernEffects;
  if (!effects) return false;
  const paint = modernTextFillStyle(ctx, effects, rect, run.color);
  const reflection = effects.reflection;
  if (reflection) {
    // approximation: flipped repaint below the run box at the start opacity
    ctx.save();
    const axis = rect.y + rect.h;
    ctx.translate(0, 2 * axis + finiteOr(reflection.distance, 0));
    ctx.scale(1, -1);
    multiplyGlobalAlpha(ctx, clamp(finiteOr(reflection.startOpacity, 0.35), 0, 1));
    ctx.fillStyle = paint;
    ctx.fillText(text, x, baselineY);
    ctx.restore();
  }
  applyModernTextShadow(ctx, effects);
  if (effects.textFill?.kind !== 'none') {
    ctx.fillStyle = paint;
    ctx.fillText(text, x, baselineY);
  }
  const outline = effects.textOutline;
  if (outline && !outline.noFill) {
    const width = Math.max(0.5, finiteOr(outline.width, 1));
    ctx.strokeStyle = outline.color ?? run.color;
    ctx.lineWidth = width;
    ctx.setLineDash(shapeDashPattern(outline.dash, width));
    ctx.strokeText(text, x, baselineY);
    ctx.setLineDash([]);
  }
  return true;
}

function drawCanvasTextWithEffects(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  run: Pick<TextRunPrimitive | GlyphRunPrimitive, 'color' | 'textShadow' | 'textOutline'> & {
    font?: string;
    size?: number;
  }
): void {
  if (!run.textShadow && !run.textOutline) {
    ctx.fillText(text, x, y);
    return;
  }
  const fontSize = run.size ?? fontSizeFromCssFont(run.font);
  const unit = Math.max(0.5, fontSize / 16);
  if (run.textShadow === 'emboss') {
    ctx.fillStyle = 'rgba(255,255,255,0.5)';
    ctx.fillText(text, x + unit, y + unit);
    ctx.fillStyle = 'rgba(0,0,0,0.3)';
    ctx.fillText(text, x - unit, y - unit);
  } else if (run.textShadow === 'imprint') {
    ctx.fillStyle = 'rgba(255,255,255,0.5)';
    ctx.fillText(text, x - unit, y - unit);
    ctx.fillStyle = 'rgba(0,0,0,0.3)';
    ctx.fillText(text, x + unit, y + unit);
  } else {
    applyCanvasTextShadow(ctx, run.textShadow, fontSize);
  }

  ctx.fillStyle = run.color;
  if (run.textOutline) {
    ctx.strokeStyle = run.color;
    ctx.lineWidth = unit;
    ctx.strokeText(text, x, y);
    return;
  }
  ctx.fillText(text, x, y);
}

function textEffectsNeedIsolation(
  run: Pick<TextRunPrimitive | GlyphRunPrimitive, 'textShadow' | 'textOutline' | 'modernEffects'>
): boolean {
  return run.textShadow !== undefined || run.textOutline === true || run.modernEffects != null;
}

function applyCanvasTextShadow(
  ctx: CanvasRenderingContext2D,
  shadow: TextRunPrimitive['textShadow'] | GlyphRunPrimitive['textShadow'],
  fontSize: number = 16
): void {
  if (shadow !== 'shadow') return;
  ctx.shadowColor = 'rgba(0,0,0,0.3)';
  ctx.shadowBlur = Math.max(1, fontSize / 8);
  ctx.shadowOffsetX = Math.max(0.5, fontSize / 16);
  ctx.shadowOffsetY = Math.max(0.5, fontSize / 16);
}

function drawTextEmphasisMarks(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  baselineY: number,
  width: number,
  run: Pick<TextRunPrimitive | GlyphRunPrimitive, 'color' | 'emphasisMark'> & {
    font?: string;
    size?: number;
  }
): void {
  if (!run.emphasisMark) return;
  const chars = [...text];
  if (chars.length === 0 || width <= 0) return;
  const mark = emphasisMarkChar(run.emphasisMark);
  if (!mark) return;
  const size = run.size ?? fontSizeFromCssFont(run.font);
  const markSize = Math.max(4, size * 0.5);
  const y = run.emphasisMark === 'underDot' ? baselineY + size * 0.45 : baselineY - size * 0.85;
  const step = width / chars.length;
  ctx.save();
  ctx.font = `${markSize}px sans-serif`;
  ctx.fillStyle = run.color;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  for (let i = 0; i < chars.length; i++) {
    if (/\s/.test(chars[i])) continue;
    ctx.fillText(mark, x + step * i + step / 2, y);
  }
  ctx.restore();
}

function drawGlyphEmphasisMarks(ctx: CanvasRenderingContext2D, run: GlyphRunPrimitive): void {
  if (!run.emphasisMark || run.glyphs.length === 0) return;
  const mark = emphasisMarkChar(run.emphasisMark);
  if (!mark) return;
  const clusters = new Map<number, { minX: number; maxX: number; baselineY: number }>();
  for (const glyph of run.glyphs) {
    const existing = clusters.get(glyph.cluster);
    const right = glyph.x + Math.max(0, glyph.advance ?? 0);
    if (existing) {
      existing.minX = Math.min(existing.minX, glyph.x);
      existing.maxX = Math.max(existing.maxX, right);
      existing.baselineY = Math.max(existing.baselineY, glyph.y);
    } else {
      clusters.set(glyph.cluster, { minX: glyph.x, maxX: right, baselineY: glyph.y });
    }
  }
  const offsets = [...clusters.keys()].sort((a, b) => a - b);
  const bytes = new TextEncoder().encode(run.text);
  const markSize = Math.max(4, run.size * 0.5);
  ctx.save();
  ctx.font = `${markSize}px sans-serif`;
  ctx.fillStyle = run.color;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  for (let i = 0; i < offsets.length; i++) {
    const start = offsets[i];
    const end = offsets[i + 1] ?? bytes.length;
    const source = new TextDecoder().decode(bytes.slice(start, end));
    if (!source || /^\s+$/u.test(source)) continue;
    const cluster = clusters.get(start);
    if (!cluster) continue;
    const fallbackAdvance = run.size * 0.5;
    const maxX = cluster.maxX > cluster.minX ? cluster.maxX : cluster.minX + fallbackAdvance;
    const y =
      run.emphasisMark === 'underDot'
        ? cluster.baselineY + run.size * 0.45
        : cluster.baselineY - run.size * 0.85;
    ctx.fillText(mark, (cluster.minX + maxX) / 2, y);
  }
  ctx.restore();
}

function emphasisMarkChar(mark: NonNullable<TextRunPrimitive['emphasisMark']>): string | null {
  switch (mark) {
    case 'dot':
    case 'underDot':
      return '\u2022';
    case 'comma':
      return '\ufe45';
    case 'circle':
      return '\u25cb';
    default:
      return null;
  }
}

function fontSizeFromCssFont(font: string | undefined): number {
  const match = font?.match(/(\d+(?:\.\d+)?)px\b/);
  return match ? Number(match[1]) : 16;
}

// intrinsic pixel size of a CanvasImageSource, needed only to convert the
// contract's fractional crop into a source rect
function intrinsicSize(source: CanvasImageSource): { width: number; height: number } {
  const s = source as unknown as {
    naturalWidth?: number;
    naturalHeight?: number;
    videoWidth?: number;
    videoHeight?: number;
    width?: number | { baseVal: { value: number } };
    height?: number | { baseVal: { value: number } };
  };
  if (typeof s.naturalWidth === 'number' && typeof s.naturalHeight === 'number') {
    return { width: s.naturalWidth, height: s.naturalHeight };
  }
  if (typeof s.videoWidth === 'number' && typeof s.videoHeight === 'number') {
    return { width: s.videoWidth, height: s.videoHeight };
  }
  const width = typeof s.width === 'number' ? s.width : (s.width?.baseVal.value ?? 0);
  const height = typeof s.height === 'number' ? s.height : (s.height?.baseVal.value ?? 0);
  return { width, height };
}

function multiplyGlobalAlpha(ctx: CanvasRenderingContext2D, opacity: number | undefined): void {
  if (opacity === undefined) return;
  const current = Number.isFinite(ctx.globalAlpha) ? ctx.globalAlpha : 1;
  ctx.globalAlpha = current * clamp(opacity, 0, 1);
}

function finiteOr(value: number | undefined, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function mixHexColor(color: string, mix: string, amount: number): string {
  const a = parseHexColor(color);
  const b = parseHexColor(mix);
  if (!a || !b) return color;
  const channel = (left: number, right: number) =>
    Math.round(left + (right - left) * clamp(amount, 0, 1))
      .toString(16)
      .padStart(2, '0');
  return `#${channel(a.r, b.r)}${channel(a.g, b.g)}${channel(a.b, b.b)}`;
}

function parseHexColor(color: string): { r: number; g: number; b: number } | null {
  const match = color.match(/^#([0-9a-f]{3}|[0-9a-f]{6})$/i);
  if (!match) return null;
  const hex =
    match[1].length === 3 ? [...match[1]].map((value) => value + value).join('') : match[1];
  return {
    r: Number.parseInt(hex.slice(0, 2), 16),
    g: Number.parseInt(hex.slice(2, 4), 16),
    b: Number.parseInt(hex.slice(4, 6), 16),
  };
}

function unknownPrimitiveError(value: never): Error {
  const kind = (value as { kind?: unknown }).kind;
  return new Error(`[canvasBackend] unsupported display primitive kind: ${String(kind)}`);
}
