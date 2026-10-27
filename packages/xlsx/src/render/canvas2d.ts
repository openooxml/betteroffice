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
  LineCmd,
  PathCmd,
  PatternFillCmd,
  Rect,
  TextAlign,
  TextCmd,
} from '../display-list/types';
import { PATTERN_MASKS } from './patterns';
import { paintNumberFormat } from './number-format';
import { wrapLines, wrappedLineY } from './wrap';

const PX_PER_PT = 96 / 72;
const STACKED_LINE_HEIGHT_RATIO = 1.2;
const STACKED_ASCENT_RATIO = 0.8;

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

/**
 * Paint a display list into a 2D context. `dpr` maps device-independent list
 * coordinates onto the backing store, so callers size the canvas at
 * `width * dpr` × `height * dpr`. `origin` is in backing-store pixels.
 */
export function paintDisplayList(
  ctx: CanvasRenderingContext2D,
  dl: DisplayList,
  dpr: number,
  origin: { x: number; y: number } = { x: 0, y: 0 }
): void {
  ctx.save();
  ctx.setTransform(dpr, 0, 0, dpr, origin.x, origin.y);
  ctx.beginPath();
  ctx.rect(0, 0, dl.width, dl.height);
  ctx.clip();
  ctx.clearRect(0, 0, dl.width, dl.height);
  for (const cmd of dl.commands) paintCommand(ctx, cmd, dpr, origin);
  ctx.restore();
}

function paintCommand(ctx: CanvasRenderingContext2D, cmd: DrawCmd, dpr: number, origin: { x: number; y: number }): void {
  switch (cmd.op) {
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
      paintText(ctx, cmd);
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
    const tile = typeof OffscreenCanvas === 'undefined' ? document.createElement('canvas') : new OffscreenCanvas(8, 8);
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
function paintText(ctx: CanvasRenderingContext2D, cmd: TextCmd): void {
  ctx.save();
  if (cmd.clip) {
    ctx.beginPath();
    ctx.rect(cmd.clip.x, cmd.clip.y, cmd.clip.w, cmd.clip.h);
    ctx.clip();
  }
  ctx.fillStyle = cmd.color;
  ctx.font = fontString(cmd);
  ctx.textBaseline = 'alphabetic';
  if (cmd.stacked) {
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
    lines.forEach((line, index) => {
      paintRun(ctx, { ...cmd, y: wrappedLineY(cmd.y, index, lines.length, wrap) }, line, cmd.x, align);
    });
  } else {
    const align = cmd.align ?? 'left';
    ctx.textAlign = ALIGN_TO_TEXT_ALIGN[align];
    paintRun(ctx, cmd, cmd.text, cmd.x, align);
  }
  ctx.restore();
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
  if (cmd.underline || cmd.dashedUnderline) ctx.fillRect(-width / 2, baseline + fontSize * 0.1, width, thickness);
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
    if (cmd.underline) ctx.fillRect(x, baseline + fontSize * 0.1, w, thickness);
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

// css font shorthand: `[italic] [bold] <size>px <family>`. Workbook fonts are
// often not installed on the host (Calibri on macOS); without a fallback the
// browser's unknown-family default is serif.
function fontString(cmd: TextCmd): string {
  const style = cmd.italic ? 'italic ' : '';
  const weight = cmd.bold ? 'bold ' : '';
  const family = cmd.fontFamily
    ? cmd.chart
      ? cmd.fontFamily
      : `"${cmd.fontFamily.replace(/"/g, '\\"')}", sans-serif`
    : 'sans-serif';
  return `${style}${weight}${fontSizePx(cmd)}px ${family}`;
}

function fontSizePx(cmd: TextCmd): number {
  return cmd.fontSize * PX_PER_PT;
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
  if (cmd.underline) ctx.fillRect(x0, cmd.y + fontSize * 0.1, width, thickness);
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
