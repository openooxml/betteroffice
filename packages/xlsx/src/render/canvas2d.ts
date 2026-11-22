/**
 * Canvas2D backend for the display list — the browser paint target.
 *
 * DOM canvas types are used deliberately here; this file is the browser backend
 * and is excluded from the pure seam (viewport/ + display-list/). The function
 * is otherwise pure: it reads a plain display list and issues draw calls, with
 * no allocation of app state and no reads back from the canvas.
 */

import type {
  DisplayList,
  DrawCmd,
  ImageCmd,
  LineCmd,
  PathCmd,
  PatternFillCmd,
  Rect,
  TextAlign,
  TextCmd,
  TextWrap,
} from '../display-list/types';
import { fontSizePx, fontString } from './fontString';
import { PATTERN_MASKS } from './patterns';
import { paintNumberFormat } from './number-format';
import { runPieces, wrapRunLines, type TextPiece } from './textRuns';
import { wrapLines, wrappedLineY } from './wrap';

const STACKED_LINE_HEIGHT_RATIO = 1.2;
const STACKED_ASCENT_RATIO = 0.8;
const PX_PER_PT = 96 / 72;
const FIT_TOLERANCE = 1e-4;
const WRAPPED_ASCENT_PER_PT = PX_PER_PT * 0.7;
const WRAPPED_DESCENT_PER_PT = PX_PER_PT * 0.2;

const ALIGN_TO_TEXT_ALIGN: Record<NonNullable<TextCmd['align']>, CanvasTextAlign> = {
  left: 'left',
  center: 'center',
  right: 'right',
};

const LINE_DASH: Record<Exclude<NonNullable<LineCmd['style']>, 'double'>, number[]> = {
  dashed: [4, 2],
  dotted: [1, 2],
  mediumDashed: [8, 4],
  dashDot: [4, 2, 1, 2],
  mediumDashDot: [8, 4, 2, 4],
  dashDotDot: [4, 2, 1, 2, 1, 2],
  mediumDashDotDot: [8, 4, 2, 4, 2, 4],
  slantDashDot: [4, 2, 1, 2],
};

const patternCache = new WeakMap<CanvasRenderingContext2D, Map<string, CanvasPattern>>();

export type PictureSource = ImageBitmap | HTMLImageElement | HTMLCanvasElement | OffscreenCanvas;

export interface PaintOptions {
  /**
   * The decoded image for an `image` command's `assetId`. Without one, or when
   * it returns nothing, the picture is skipped.
   */
  resolveImage?: (assetId: string) => PictureSource | null | undefined;
}

/**
 * Paint a display list into a 2D context. `dpr` maps device-independent list
 * coordinates onto the backing store, so callers size the canvas at
 * `width * dpr` × `height * dpr`. `origin` is in backing-store pixels.
 */
export function paintDisplayList(
  ctx: CanvasRenderingContext2D,
  dl: DisplayList,
  dpr: number,
  origin: { x: number; y: number } = { x: 0, y: 0 },
  options: PaintOptions = {}
): void {
  ctx.save();
  ctx.setTransform(dpr, 0, 0, dpr, origin.x, origin.y);
  ctx.beginPath();
  ctx.rect(0, 0, dl.width, dl.height);
  ctx.clip();
  ctx.clearRect(0, 0, dl.width, dl.height);
  const left = Math.max(0, -origin.x / dpr);
  const top = Math.max(0, -origin.y / dpr);
  const right = Math.min(dl.width, ((ctx.canvas?.width ?? Infinity) - origin.x) / dpr);
  const bottom = Math.min(dl.height, ((ctx.canvas?.height ?? Infinity) - origin.y) / dpr);
  const bounds = { x: left, y: top, w: right - left, h: bottom - top };
  for (const cmd of dl.commands) paintCommand(ctx, cmd, dpr, origin, options, bounds);
  ctx.restore();
}

function paintCommand(
  ctx: CanvasRenderingContext2D,
  cmd: DrawCmd,
  dpr: number,
  origin: { x: number; y: number },
  options: PaintOptions,
  bounds: Rect,
): void {
  switch (cmd.op) {
    case 'image':
      paintClipped(ctx, cmd.clip, () => paintImage(ctx, cmd, options));
      return;
    case 'fillRect':
      paintClipped(ctx, cmd.clip, () => {
        ctx.fillStyle = cmd.color;
        ctx.fillRect(cmd.x, cmd.y, cmd.w, cmd.h);
      });
      return;
    case 'patternFill':
      paintClipped(ctx, cmd.clip, () => paintPattern(ctx, cmd, dpr, origin));
      return;
    case 'line':
      paintClipped(ctx, cmd.clip, () => paintLine(ctx, cmd, dpr, origin));
      return;
    case 'path':
      paintClipped(ctx, cmd.clip, () => paintPath(ctx, cmd));
      return;
    case 'text':
      paintText(ctx, cmd, bounds);
      return;
  }
}

function paintPattern(ctx: CanvasRenderingContext2D, cmd: PatternFillCmd, dpr: number, origin: { x: number; y: number }): void {
  let cache = patternCache.get(ctx);
  if (!cache) {
    cache = new Map();
    patternCache.set(ctx, cache);
  }
  const key = `${cmd.pattern}/${cmd.fg}/${cmd.bg}/${dpr}`;
  let pattern = cache.get(key);
  if (!pattern) {
    const tile = typeof OffscreenCanvas === 'function' ? new OffscreenCanvas(8, 8)
      : typeof document !== 'undefined' ? document.createElement('canvas')
      : new (ctx.canvas.constructor as new (width: number, height: number) => HTMLCanvasElement)(8, 8);
    tile.width = 8;
    tile.height = 8;
    const context = tile.getContext('2d') as CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;
    const image = context.createImageData(8, 8);
    const fg = Number.parseInt(cmd.fg.slice(1), 16);
    const bg = Number.parseInt(cmd.bg.slice(1), 16);
    const mask = PATTERN_MASKS[cmd.pattern];
    for (let y = 0; y < 8; y += 1) for (let x = 0; x < 8; x += 1) {
      const color = mask[y]! & (1 << x) ? fg : bg;
      const index = (y * 8 + x) * 4;
      image.data[index] = (color >> 16) & 255;
      image.data[index + 1] = (color >> 8) & 255;
      image.data[index + 2] = color & 255;
      image.data[index + 3] = 255;
    }
    context.putImageData(image, 0, 0);
    pattern = ctx.createPattern(tile, 'repeat') ?? undefined;
    if (!pattern) return;
    if (cache.size >= 256) cache.clear();
    cache.set(key, pattern);
  }
  pattern.setTransform({ a: 1 / dpr, b: 0, c: 0, d: 1 / dpr,
    e: (Math.round(cmd.originX * dpr + origin.x) - origin.x) / dpr,
    f: (Math.round(cmd.originY * dpr + origin.y) - origin.y) / dpr });
  ctx.save();
  ctx.imageSmoothingEnabled = false;
  ctx.fillStyle = pattern;
  const x = (Math.round(cmd.x * dpr + origin.x) - origin.x) / dpr;
  const y = (Math.round(cmd.y * dpr + origin.y) - origin.y) / dpr;
  ctx.fillRect(x, y, (Math.round((cmd.x + cmd.w) * dpr + origin.x) - origin.x) / dpr - x,
    (Math.round((cmd.y + cmd.h) * dpr + origin.y) - origin.y) / dpr - y);
  ctx.restore();
}

function paintImage(ctx: CanvasRenderingContext2D, cmd: ImageCmd, options: PaintOptions): void {
  const image = options.resolveImage?.(cmd.assetId);
  if (!image) return;
  const [width, height] = 'naturalWidth' in image ? [image.naturalWidth, image.naturalHeight] : [image.width, image.height];
  const { left = 0, top = 0, right = 0, bottom = 0 } = cmd.crop ?? {};
  const sw = width * (1 - left - right);
  const sh = height * (1 - top - bottom);
  if (!(sw > 0 && sh > 0 && cmd.w > 0 && cmd.h > 0)) return;
  ctx.drawImage(image, width * left, height * top, sw, sh, cmd.x, cmd.y, cmd.w, cmd.h);
}

function paintClipped(
  ctx: CanvasRenderingContext2D,
  clip: Rect | undefined,
  paint: () => void
): void {
  if (!clip) {
    paint();
    return;
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(clip.x, clip.y, clip.w, clip.h);
  ctx.clip();
  paint();
  ctx.restore();
}

function paintPath(ctx: CanvasRenderingContext2D, cmd: PathCmd): void {
  ctx.beginPath();
  for (const command of cmd.commands) {
    switch (command.type) {
      case 'move':
        ctx.moveTo(command.x, command.y);
        break;
      case 'line':
        ctx.lineTo(command.x, command.y);
        break;
      case 'quad':
        ctx.quadraticCurveTo(command.cpx, command.cpy, command.x, command.y);
        break;
      case 'cubic':
        ctx.bezierCurveTo(
          command.cp1x,
          command.cp1y,
          command.cp2x,
          command.cp2y,
          command.x,
          command.y
        );
        break;
      case 'close':
        ctx.closePath();
        break;
    }
  }
  ctx.fillStyle = cmd.fill;
  ctx.fill();
  if (cmd.stroke) {
    ctx.strokeStyle = cmd.stroke.color;
    ctx.lineWidth = cmd.stroke.width;
    ctx.setLineDash([]);
    ctx.stroke();
  }
}

function paintLine(ctx: CanvasRenderingContext2D, cmd: LineCmd, dpr: number, origin: { x: number; y: number }): void {
  ctx.strokeStyle = cmd.color;
  ctx.save();
  ctx.lineCap = 'butt';
  let { x1, y1, x2, y2 } = cmd;
  const width = Math.max(1, Math.round(cmd.width * dpr));
  // a double without a rail width keeps the legacy passes around the unsnapped line
  const legacyDouble = cmd.style === 'double' && cmd.rail === undefined;
  if (cmd.style && !legacyDouble) {
    if (y1 === y2) {
      y1 = y2 = (Math.floor(y1 * dpr + origin.y) + (width % 2) / 2 - origin.y) / dpr;
      x1 = (Math.round(x1 * dpr + origin.x) - origin.x) / dpr;
      x2 = (Math.round(x2 * dpr + origin.x) - origin.x) / dpr;
    } else if (x1 === x2) {
      x1 = x2 = (Math.floor(x1 * dpr + origin.x) + (width % 2) / 2 - origin.x) / dpr;
      y1 = (Math.round(y1 * dpr + origin.y) - origin.y) / dpr;
      y2 = (Math.round(y2 * dpr + origin.y) - origin.y) / dpr;
    }
  }
  if (cmd.style === 'double') {
    const off = cmd.rail === undefined ? Math.max(cmd.width * 0.8, 0.8) : (cmd.width - cmd.rail) / 2;
    const horizontal = Math.abs(y1 - y2) <= Math.abs(x1 - x2);
    const [dx, dy] = horizontal ? [0, off] : [off, 0];
    ctx.lineWidth = cmd.rail ?? Math.max(cmd.width * 0.6, 0.5);
    ctx.setLineDash([]);
    strokeSegment(ctx, x1 - dx, y1 - dy, x2 - dx, y2 - dy);
    strokeSegment(ctx, x1 + dx, y1 + dy, x2 + dx, y2 + dy);
    ctx.restore();
    return;
  }
  ctx.lineWidth = cmd.style ? width / dpr : cmd.width;
  const scale = cmd.style?.startsWith('medium') ? cmd.width / 2 : cmd.width;
  ctx.setLineDash(cmd.style ? LINE_DASH[cmd.style].map((run) => run * scale) : []);
  strokeSegment(ctx, x1, y1, x2, y2);
  ctx.restore();
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

// text is clipped to its cell box; save/clip/restore keeps the clip local to
// this run so it never bleeds into later commands. underline/strike are drawn
// as lines positioned off the measured run box, since canvas has no native
// text-decoration.
function paintText(ctx: CanvasRenderingContext2D, cmd: TextCmd, bounds: Rect): void {
  ctx.save();
  if (cmd.clip) {
    ctx.beginPath();
    ctx.rect(cmd.clip.x, cmd.clip.y, cmd.clip.w, cmd.clip.h);
    ctx.clip();
  }
  ctx.fillStyle = cmd.color;
  const font = fontString(cmd);
  ctx.font = font;
  (ctx as CanvasRenderingContext2D & { textRendering?: string }).textRendering = 'geometricPrecision';
  ctx.textBaseline = 'alphabetic';
  const pieces = cmd.runs?.length ? runPieces(cmd) : undefined;
  if (pieces && cmd.wrap) {
    paintWrappedPieces(ctx, cmd, cmd.wrap, pieces, newPiecePaint(font), bounds);
  } else if (pieces) {
    paintPieces(ctx, cmd, pieces, cmd.y, newPiecePaint(font));
  } else if (cmd.stacked) {
    paintStackedText(ctx, cmd);
  } else if (cmd.rotationDeg !== undefined && cmd.rotationDeg !== null) {
    paintRotatedText(ctx, cmd, cmd.rotationDeg);
  } else if (cmd.numberFormat) {
    paintNumberFormat(ctx, cmd, cmd.numberFormat, paintRun);
  } else if (cmd.wrap) {
    const wrap = cmd.wrap;
    const align = cmd.align ?? 'left';
    ctx.textAlign = ALIGN_TO_TEXT_ALIGN[align];
    const lines = wrapLines(cmd.text, wrap.width, (run) => ctx.measureText(run).width);
    const count = Math.min(lines.length, wrap.maxLines ?? lines.length);
    lines.slice(0, count).forEach((line, index) => {
      paintRun(ctx, { ...cmd, y: wrappedLineY(cmd.y, index, count, wrap) }, line, cmd.x, align);
    });
  } else {
    const align = cmd.align ?? 'left';
    ctx.textAlign = ALIGN_TO_TEXT_ALIGN[align];
    paintRun(ctx, cmd, cmd.text, cmd.x, align);
  }
  ctx.restore();
}

/**
 * What painting one command's pieces has done to the canvas so far: the font string it last assigned, each piece
 * style's font string, and the width of every text it measured in each font. Lives for one command's paint.
 */
interface PiecePaint {
  font: string;
  fonts: Map<TextPiece['cmd'], string>;
  widths: Map<string, Map<string, number>>;
}

function newPiecePaint(font: string): PiecePaint {
  return { font, fonts: new Map(), widths: new Map() };
}

function pieceFont(paint: PiecePaint, piece: TextPiece): string {
  let font = paint.fonts.get(piece.cmd);
  if (font === undefined) {
    font = fontString(piece.cmd);
    paint.fonts.set(piece.cmd, font);
  }
  return font;
}

function useFont(ctx: CanvasRenderingContext2D, paint: PiecePaint, font: string): void {
  if (paint.font === font) return;
  ctx.font = font;
  paint.font = font;
}

function measurePiece(ctx: CanvasRenderingContext2D, paint: PiecePaint, piece: TextPiece, text: string): number {
  const font = pieceFont(paint, piece);
  let widths = paint.widths.get(font);
  if (widths === undefined) {
    widths = new Map();
    paint.widths.set(font, widths);
  }
  let width = widths.get(text);
  if (width === undefined) {
    useFont(ctx, paint, font);
    width = ctx.measureText(text).width;
    widths.set(text, width);
  }
  return width;
}

function paintPieces(ctx: CanvasRenderingContext2D, cmd: TextCmd, pieces: TextPiece[], y: number, paint: PiecePaint): void {
  const widths = pieces.map((piece) => measurePiece(ctx, paint, piece, piece.text));
  const total = widths.reduce((sum, width) => sum + width, 0);
  const align = cmd.align ?? 'left';
  let x = align === 'right' ? cmd.x - total : align === 'center' ? cmd.x - total / 2 : cmd.x;
  ctx.textAlign = 'left';
  pieces.forEach((piece, index) => {
    useFont(ctx, paint, pieceFont(paint, piece));
    paintPiece(ctx, piece.cmd, piece.text, x, y, widths[index]!);
    x += widths[index]!;
  });
}

function paintWrappedPieces(
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  wrap: TextWrap,
  pieces: TextPiece[],
  paint: PiecePaint,
  bounds: Rect,
): void {
  const wrapped = wrapRunLines(cmd.text, pieces, wrap.width, (piece, part) => measurePiece(ctx, paint, piece, part));
  const lines = wrapped.slice(0, Math.max(0, wrap.maxLines ?? wrapped.length));
  const cellHeight = cmd.clip ? cmd.clip.h + 1 : Infinity;
  const layout = pieces.some((piece) => piece.cmd.fontSize !== pieces[0]!.cmd.fontSize)
    ? wrappedLineMetrics(cmd.y, cmd.fontSize, wrap, lines, cellHeight)
    : undefined;
  if (!layout) {
    lines.forEach((line, index) => {
      paintPieces(ctx, cmd, line, wrappedLineY(cmd.y, index, lines.length, wrap), paint);
    });
    return;
  }

  const clip = cmd.clip ?? bounds;
  const clipTop = Math.max(clip.y, bounds.y);
  const clipBottom = Math.min(clip.y + clip.h, bounds.y + bounds.h);
  const clipLeft = Math.max(clip.x, bounds.x);
  const clipRight = Math.min(clip.x + clip.w, bounds.x + bounds.w);
  if (clipTop >= clipBottom || clipLeft >= clipRight) return;

  const { blockTop, metrics } = layout;
  const baseSized = metrics.every(({ size }) => size === cmd.fontSize);
  for (const [index, { top, size }] of metrics.entries()) {
    const ascent = size * WRAPPED_ASCENT_PER_PT;
    const baseline = baseSized ? wrappedLineY(cmd.y, index, metrics.length, wrap) : blockTop + (top + ascent);
    const lineTop = baseSized ? baseline - ascent : blockTop + top;
    if (lineTop >= clipBottom) break;
    if (baseline + size * WRAPPED_DESCENT_PER_PT > clipTop) paintPieces(ctx, cmd, lines[index]!, baseline, paint);
  }
}

function wrappedLineMetrics(y: number, fontSize: number, wrap: TextWrap, lines: TextPiece[][], cellHeight: number) {
  if (!Number.isFinite(fontSize) || fontSize <= 0 || !Number.isFinite(wrap.lineHeight) || wrap.lineHeight <= 0) {
    return undefined;
  }

  const metrics: { top: number; size: number }[] = [];
  let totalHeight = 0;
  let baseSized = true;
  for (const [index, line] of lines.entries()) {
    const size = line.reduce((largest, piece) => Math.max(largest, piece.cmd.fontSize), line[0]?.cmd.fontSize ?? fontSize);
    const height = wrap.lineHeight * (size / fontSize);
    if (!Number.isFinite(height) || height <= 0) return undefined;
    baseSized &&= size === fontSize;
    const topInLines = baseSized ? index : totalHeight / height;
    const room = (cellHeight - size * PX_PER_PT / 2) / height + FIT_TOLERANCE;
    if (index > 0 && Number.isFinite(cellHeight) && topInLines > room) break;
    metrics.push({ top: totalHeight, size });
    totalHeight += height;
  }
  const last = metrics[metrics.length - 1];
  if (!last) return undefined;
  const lastBaseline = last.top + last.size * WRAPPED_ASCENT_PER_PT;
  const lastBottom = lastBaseline + last.size * WRAPPED_DESCENT_PER_PT;
  if (!Number.isFinite(totalHeight) || !Number.isFinite(lastBottom)) return undefined;
  const topAnchor = y - fontSize * WRAPPED_ASCENT_PER_PT;
  const bottomAnchor = y + fontSize * WRAPPED_DESCENT_PER_PT - lastBottom;
  const blockTop = wrap.valign === 'top' ? topAnchor : wrap.valign === 'center' ? topAnchor / 2 + bottomAnchor / 2 : bottomAnchor;
  return Number.isFinite(blockTop) ? { blockTop, metrics } : undefined;
}

/** `paintRun` for a piece painted left-aligned at `y`, `width` wide, in a font without highlight or dashed underline. */
function paintPiece(ctx: CanvasRenderingContext2D, cmd: TextCmd, text: string, x: number, y: number, width: number): void {
  ctx.fillStyle = cmd.color;
  ctx.fillText(text, x, y);
  if (!cmd.underline && !cmd.strike) return;
  const fontSize = fontSizePx(cmd);
  const thickness = Math.max(fontSize * 0.05, 0.5);
  ctx.fillStyle = cmd.color;
  if (cmd.underline) fillUnderline(ctx, cmd, x, y + fontSize * 0.1, width, thickness);
  if (cmd.strike) ctx.fillRect(x, y - fontSize * 0.26, width, thickness);
}

function rotatedAvailableLength(box: Rect, deg: number, fontSize: number): number {
  const line = fontSize * STACKED_LINE_HEIGHT_RATIO;
  const rad = (Math.abs(deg) * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad));
  const sin = Math.abs(Math.sin(rad));
  let available = Number.POSITIVE_INFINITY;
  if (cos > 1e-6) available = Math.min(available, (box.w - line * sin) / cos);
  if (sin > 1e-6) available = Math.min(available, (box.h - line * cos) / sin);
  if (!Number.isFinite(available)) available = Math.max(box.w, box.h);
  return Math.max(0, available);
}

function graphemeClusters(text: string): string[] {
  const SegmenterCtor = (Intl as unknown as {
    Segmenter?: new (
      locales?: string | string[],
      options?: { granularity: 'grapheme' }
    ) => { segment(value: string): Iterable<{ segment: string }> };
  }).Segmenter;
  if (typeof SegmenterCtor === 'function') {
    try {
      return [...new SegmenterCtor(undefined, { granularity: 'grapheme' }).segment(text)].map((part) => part.segment);
    } catch {
      return [...text];
    }
  }
  return [...text];
}

function paintRotatedText(ctx: CanvasRenderingContext2D, cmd: TextCmd, deg: number): void {
  const box = cmd.rotationBox ?? cmd.clip ?? { x: cmd.x, y: cmd.y, w: 0, h: 0 };
  const fontSize = fontSizePx(cmd);
  let segments: { text: string; offset: number }[] = [{ text: cmd.text, offset: 0 }];
  let width = ctx.measureText(cmd.text).width;
  if (cmd.numberFormat) {
    const available = rotatedAvailableLength(box, deg, fontSize);
    const raw: { text: string; x: number }[] = [];
    paintNumberFormat(
      ctx,
      { ...cmd, numberFormat: undefined },
      { ...cmd.numberFormat, left: 0, width: available },
      (_ctx, _cmd, text, x, align) => {
        const w = ctx.measureText(text).width;
        const left = align === 'right' ? x - w : align === 'center' ? x - w / 2 : x;
        raw.push({ text, x: left });
      }
    );
    if (raw.length > 0) {
      const minX = Math.min(...raw.map((run) => run.x));
      const maxX = Math.max(...raw.map((run) => run.x + ctx.measureText(run.text).width));
      segments = raw.map((run) => ({ text: run.text, offset: run.x - minX }));
      width = Math.max(0, maxX - minX);
    }
  }
  const metrics = ctx.measureText(cmd.text);
  const ascent = (deg === 90 ? metrics.fontBoundingBoxAscent : undefined)
    ?? (metrics.actualBoundingBoxAscent || fontSize * 0.8);
  const descent = (deg === 90 ? metrics.fontBoundingBoxDescent : undefined)
    ?? (metrics.actualBoundingBoxDescent || fontSize * 0.2);
  const lineHeight = ascent + descent;
  const angle = (-deg * Math.PI) / 180;
  const cos = Math.cos(angle);
  const sin = Math.sin(angle);
  const bboxW = Math.abs(width * cos) + Math.abs(lineHeight * sin);
  const bboxH = Math.abs(width * sin) + Math.abs(lineHeight * cos);
  const align = cmd.align ?? 'left';
  const bboxLeft = align === 'center' ? box.x + (box.w - bboxW) / 2 : align === 'right' ? box.x + box.w - bboxW : box.x;
  const valign = cmd.rotationValign ?? 'bottom';
  const bboxTop = valign === 'top' ? box.y : valign === 'center' ? box.y + (box.h - bboxH) / 2 : box.y + box.h - bboxH;
  const cx = bboxLeft + bboxW / 2;
  const cy = bboxTop + bboxH / 2;
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(angle);
  if (cmd.highlight) {
    ctx.fillStyle = cmd.highlight;
    ctx.fillRect(-width / 2 - 2, -lineHeight / 2 - 1, width + 4, lineHeight + 2);
    ctx.fillStyle = cmd.color;
  }
  ctx.textBaseline = 'alphabetic';
  const baseline = (ascent - descent) / 2;
  if (cmd.numberFormat) {
    ctx.textAlign = 'left';
    for (const segment of segments) ctx.fillText(segment.text, -width / 2 + segment.offset, baseline);
  } else {
    ctx.textAlign = 'center';
    ctx.fillText(cmd.text, 0, baseline);
  }
  const thickness = Math.max(fontSize * 0.05, 0.5);
  ctx.fillStyle = cmd.color;
  if (cmd.underline) fillUnderline(ctx, cmd, -width / 2, baseline + fontSize * 0.1, width, thickness);
  else if (cmd.dashedUnderline) ctx.fillRect(-width / 2, baseline + fontSize * 0.1, width, thickness);
  if (cmd.strike) ctx.fillRect(-width / 2, baseline - fontSize * 0.26, width, thickness);
  ctx.restore();
}

function paintStackedText(ctx: CanvasRenderingContext2D, cmd: TextCmd): void {
  const box = cmd.rotationBox ?? cmd.clip ?? { x: cmd.x, y: cmd.y, w: 0, h: 0 };
  const fontSize = fontSizePx(cmd);
  const lineHeight = fontSize * STACKED_LINE_HEIGHT_RATIO;
  const ascent = fontSize * STACKED_ASCENT_RATIO;
  const clusters = graphemeClusters(cmd.text);
  const totalHeight = clusters.length * lineHeight;
  const valign = cmd.rotationValign ?? 'bottom';
  const top = valign === 'top' ? box.y : valign === 'center' ? box.y + (box.h - totalHeight) / 2 : box.y + box.h - totalHeight;
  const align = cmd.align ?? 'left';
  const clip = cmd.clip ?? box;
  const first = Math.max(0, Math.floor((clip.y - top) / lineHeight));
  const last = Math.min(clusters.length - 1, Math.ceil((clip.y + clip.h - top) / lineHeight) - 1);
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  for (let i = first; i <= last; i += 1) {
    const cluster = clusters[i]!;
    const w = ctx.measureText(cluster).width;
    const x = align === 'center' ? box.x + (box.w - w) / 2 : align === 'right' ? box.x + box.w - w : box.x;
    if (x + w <= clip.x || x >= clip.x + clip.w) continue;
    const baseline = top + i * lineHeight + ascent;
    if (cmd.highlight) {
      ctx.fillStyle = cmd.highlight;
      ctx.fillRect(x - 2, baseline - ascent - 1, w + 4, lineHeight + 2);
      ctx.fillStyle = cmd.color;
    }
    ctx.fillText(cluster, x, baseline);
    const thickness = Math.max(fontSize * 0.05, 0.5);
    ctx.fillStyle = cmd.color;
    if (cmd.underline) fillUnderline(ctx, cmd, x, baseline + fontSize * 0.1, w, thickness);
    if (cmd.strike) ctx.fillRect(x, baseline - fontSize * 0.26, w, thickness);
    if (cmd.dashedUnderline) {
      ctx.save();
      ctx.strokeStyle = cmd.color;
      ctx.lineWidth = thickness;
      ctx.setLineDash([3, 2]);
      strokeSegment(ctx, x, baseline + fontSize * 0.1, x + w, baseline + fontSize * 0.1);
      ctx.restore();
    }
  }
}

function paintRun(
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  text: string,
  x: number,
  align: TextAlign
): void {
  if (cmd.highlight) paintHighlight(ctx, cmd, text, x, align);
  ctx.fillStyle = cmd.color;
  ctx.fillText(text, x, cmd.y);
  if (cmd.underline || cmd.strike || cmd.dashedUnderline) paintDecorations(ctx, cmd, text, x, align);
}


function paintDecorations(
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  text: string,
  x: number,
  align: TextAlign
): void {
  const width = ctx.measureText(text).width;
  const x0 = align === 'right' ? x - width : align === 'center' ? x - width / 2 : x;
  const fontSize = fontSizePx(cmd);
  const thickness = Math.max(fontSize * 0.05, 0.5);
  ctx.fillStyle = cmd.color;
  if (cmd.underline) fillUnderline(ctx, cmd, x0, cmd.y + fontSize * 0.1, width, thickness);
  if (cmd.strike) ctx.fillRect(x0, cmd.y - fontSize * 0.26, width, thickness);
  if (cmd.dashedUnderline) {
    ctx.save();
    ctx.strokeStyle = cmd.color;
    ctx.lineWidth = thickness;
    ctx.setLineDash([3, 2]);
    strokeSegment(ctx, x0, cmd.y + fontSize * 0.1, x0 + width, cmd.y + fontSize * 0.1);
    ctx.restore();
  }
}

/** A double kind adds a second line one line width and at least a pixel below the first. */
function fillUnderline(
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  x: number,
  y: number,
  width: number,
  thickness: number
): void {
  ctx.fillRect(x, y, width, thickness);
  if (cmd.underlineStyle === 'double' || cmd.underlineStyle === 'doubleAccounting') {
    ctx.fillRect(x, y + thickness + Math.max(thickness, 1), width, thickness);
  }
}

function paintHighlight(
  ctx: CanvasRenderingContext2D,
  cmd: TextCmd,
  text: string,
  x: number,
  align: TextAlign
): void {
  const metrics = ctx.measureText(text);
  const width = metrics.width;
  const x0 = align === 'right' ? x - width : align === 'center' ? x - width / 2 : x;
  const fontSize = fontSizePx(cmd);
  const ascent = metrics.actualBoundingBoxAscent || fontSize * 0.8;
  const descent = metrics.actualBoundingBoxDescent || fontSize * 0.2;
  ctx.fillStyle = cmd.highlight!;
  ctx.fillRect(x0 - 2, cmd.y - ascent - 1, width + 4, ascent + descent + 2);
}
