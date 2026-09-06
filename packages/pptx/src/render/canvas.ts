import type {
  ChartPrimitive,
  GeometryPathCommand,
  ImagePrimitive,
  Paint,
  PlaceholderPrimitive,
  PositionedTextRun,
  ShapePrimitive,
  SlideDisplayList,
  SlidePrimitive,
  Stroke,
  StrokeEnd,
  TextBoxPrimitive,
} from '../types';

export type CanvasImageResolver = (
  assetId: string
) => CanvasImageSource | Promise<CanvasImageSource | null> | null;

export interface PaintSlideOptions {
  resolveImage?: CanvasImageResolver;
}

export interface SlideCanvasLike {
  width: number;
  height: number;
  style: { width: string; height: string };
}

export function sizeCanvasForSlide(
  canvas: SlideCanvasLike,
  list: Pick<SlideDisplayList, 'width' | 'height'>,
  dpr: number,
  scale = 1
): void {
  canvas.width = Math.round(list.width * scale * dpr);
  canvas.height = Math.round(list.height * scale * dpr);
  canvas.style.width = `${list.width * scale}px`;
  canvas.style.height = `${list.height * scale}px`;
}

export async function paintSlide(
  ctx: CanvasRenderingContext2D,
  list: SlideDisplayList,
  dpr = 1,
  scale = 1,
  options: PaintSlideOptions = {}
): Promise<void> {
  ctx.save();
  try {
    ctx.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0);
    ctx.clearRect(0, 0, list.width, list.height);
    if (list.background) {
      ctx.fillStyle = paintStyle(ctx, list.background, 0, 0, list.width, list.height);
      ctx.fillRect(0, 0, list.width, list.height);
    }
    for (const primitive of list.primitives) await paintPrimitive(ctx, primitive, options);
  } finally {
    ctx.restore();
  }
}

async function paintPrimitive(
  ctx: CanvasRenderingContext2D,
  primitive: SlidePrimitive,
  options: PaintSlideOptions
): Promise<void> {
  ctx.save();
  try {
    applyTransform(ctx, primitive);
    switch (primitive.kind) {
      case 'shape':
        paintShape(ctx, primitive);
        break;
      case 'image':
        await paintImage(ctx, primitive, options.resolveImage);
        break;
      case 'textBox':
        paintTextBox(ctx, primitive);
        break;
      case 'placeholder':
        paintPlaceholder(ctx, primitive);
        break;
      case 'chart':
        await paintChart(ctx, primitive, options);
        break;
    }
  } finally {
    ctx.restore();
  }
}

async function paintChart(
  ctx: CanvasRenderingContext2D,
  chart: ChartPrimitive,
  options: PaintSlideOptions
): Promise<void> {
  ctx.beginPath();
  ctx.rect(chart.x, chart.y, chart.w, chart.h);
  ctx.clip();
  for (const primitive of chart.primitives) await paintPrimitive(ctx, primitive, options);
}

function applyTransform(
  ctx: CanvasRenderingContext2D,
  primitive: Pick<SlidePrimitive, 'x' | 'y' | 'w' | 'h' | 'transform'>
): void {
  const transform = primitive.transform;
  if (!transform) return;
  const centerX = primitive.x + primitive.w / 2;
  const centerY = primitive.y + primitive.h / 2;
  ctx.translate(centerX, centerY);
  ctx.rotate(((transform.rotationDeg ?? 0) * Math.PI) / 180);
  ctx.scale(transform.flipH ? -1 : 1, transform.flipV ? -1 : 1);
  ctx.translate(-centerX, -centerY);
}

function paintShape(ctx: CanvasRenderingContext2D, shape: ShapePrimitive): void {
  buildPath(ctx, shape.path, shape.x, shape.y, shape.w, shape.h);
  if (shape.fill) {
    ctx.fillStyle = paintStyle(ctx, shape.fill, shape.x, shape.y, shape.w, shape.h);
    ctx.fill();
  }
  if (shape.stroke) {
    strokeCurrentPath(ctx, shape.stroke);
    paintLineEnds(ctx, shape);
  }
}

function pathPoints(shape: ShapePrimitive): Array<[number, number]> {
  const points: Array<[number, number]> = [];
  for (const command of shape.path) {
    if (command.type === 'close') continue;
    if (command.type === 'quad') {
      points.push([shape.x + command.cpx * shape.w, shape.y + command.cpy * shape.h]);
    } else if (command.type === 'cubic') {
      points.push([shape.x + command.cp1x * shape.w, shape.y + command.cp1y * shape.h]);
      points.push([shape.x + command.cp2x * shape.w, shape.y + command.cp2y * shape.h]);
    }
    points.push([shape.x + command.x * shape.w, shape.y + command.y * shape.h]);
  }
  return points.filter(
    (point, index) => index === 0 || point[0] !== points[index - 1][0] || point[1] !== points[index - 1][1]
  );
}

const LINE_END_KINDS = new Set(['triangle', 'stealth', 'arrow', 'diamond', 'oval']);

function paintLineEnds(ctx: CanvasRenderingContext2D, shape: ShapePrimitive): void {
  const stroke = shape.stroke;
  if (!stroke || (!stroke.headEnd && !stroke.tailEnd)) return;
  const points = pathPoints(shape);
  if (points.length < 2) return;
  const ends: Array<[StrokeEnd | undefined, [number, number], [number, number]]> = [
    [stroke.headEnd, points[1], points[0]],
    [stroke.tailEnd, points[points.length - 2], points[points.length - 1]],
  ];
  ctx.save();
  ctx.setLineDash([]);
  ctx.fillStyle = stroke.color;
  ctx.strokeStyle = stroke.color;
  ctx.lineWidth = stroke.width;
  ctx.lineJoin = 'miter';
  for (const [end, from, tip] of ends) {
    if (!end || !LINE_END_KINDS.has(end.kind)) continue;
    const [dx, dy] = [tip[0] - from[0], tip[1] - from[1]];
    const distance = Math.hypot(dx, dy);
    if (distance < 1e-6) continue;
    paintLineEnd(ctx, end, tip, dx / distance, dy / distance);
  }
  ctx.restore();
}

function paintLineEnd(
  ctx: CanvasRenderingContext2D,
  end: StrokeEnd,
  [tx, ty]: [number, number],
  ux: number,
  uy: number
): void {
  const half = end.width / 2;
  const [nx, ny] = [-uy * half, ux * half];
  const [bx, by] = [tx - ux * end.length, ty - uy * end.length];
  ctx.beginPath();
  switch (end.kind) {
    case 'triangle':
      ctx.moveTo(tx, ty);
      ctx.lineTo(bx + nx, by + ny);
      ctx.lineTo(bx - nx, by - ny);
      ctx.closePath();
      ctx.fill();
      break;
    case 'stealth':
      ctx.moveTo(tx, ty);
      ctx.lineTo(bx + nx, by + ny);
      ctx.lineTo(tx - ux * end.length * 0.6, ty - uy * end.length * 0.6);
      ctx.lineTo(bx - nx, by - ny);
      ctx.closePath();
      ctx.fill();
      break;
    case 'arrow':
      ctx.moveTo(bx + nx, by + ny);
      ctx.lineTo(tx, ty);
      ctx.lineTo(bx - nx, by - ny);
      ctx.stroke();
      break;
    case 'diamond': {
      const [ax, ay] = [(ux * end.length) / 2, (uy * end.length) / 2];
      ctx.moveTo(tx + ax, ty + ay);
      ctx.lineTo(tx + nx, ty + ny);
      ctx.lineTo(tx - ax, ty - ay);
      ctx.lineTo(tx - nx, ty - ny);
      ctx.closePath();
      ctx.fill();
      break;
    }
    case 'oval':
      ctx.ellipse(tx, ty, end.length / 2, half, Math.atan2(uy, ux), 0, Math.PI * 2);
      ctx.fill();
      break;
  }
}

/** Draws the cropped source through the picture's outline. */
function drawCropped(
  ctx: CanvasRenderingContext2D,
  source: CanvasImageSource,
  image: ImagePrimitive
): void {
  const crop = image.crop;
  const left = clampCrop(crop?.left);
  const top = clampCrop(crop?.top);
  const keptX = 1 - left - clampCrop(crop?.right);
  const keptY = 1 - top - clampCrop(crop?.bottom);
  const masked = image.path !== undefined || keptX !== 1 || keptY !== 1;
  if (keptX <= 0 || keptY <= 0) return;
  if (masked) {
    ctx.save();
    buildImageOutline(ctx, image);
    ctx.clip();
  }
  const width = sourceWidth(source);
  const height = sourceHeight(source);
  if (width > 0 && height > 0) {
    ctx.drawImage(
      source,
      left * width,
      top * height,
      keptX * width,
      keptY * height,
      image.x,
      image.y,
      image.w,
      image.h
    );
  } else {
    ctx.drawImage(source, image.x, image.y, image.w, image.h);
  }
  if (masked) ctx.restore();
}

/** The picture's own outline when it has one, else its frame. */
function buildImageOutline(ctx: CanvasRenderingContext2D, image: ImagePrimitive): void {
  if (image.path) buildPath(ctx, image.path, image.x, image.y, image.w, image.h);
  else {
    ctx.beginPath();
    ctx.rect(image.x, image.y, image.w, image.h);
  }
}

/** `a:srcRect` also encodes outsets as negatives, which canvas cannot express. */
function clampCrop(value: number | undefined): number {
  if (value === undefined || !Number.isFinite(value)) return 0;
  return Math.min(Math.max(value, 0), 1);
}

function sourceWidth(source: CanvasImageSource): number {
  if ('naturalWidth' in source && typeof source.naturalWidth === 'number') return source.naturalWidth;
  if ('width' in source && typeof source.width === 'number') return source.width;
  return 0;
}

function sourceHeight(source: CanvasImageSource): number {
  if ('naturalHeight' in source && typeof source.naturalHeight === 'number')
    return source.naturalHeight;
  if ('height' in source && typeof source.height === 'number') return source.height;
  return 0;
}

function buildPath(
  ctx: CanvasRenderingContext2D,
  commands: GeometryPathCommand[],
  x: number,
  y: number,
  width: number,
  height: number
): void {
  ctx.beginPath();
  for (const command of commands) {
    switch (command.type) {
      case 'move':
        ctx.moveTo(x + command.x * width, y + command.y * height);
        break;
      case 'line':
        ctx.lineTo(x + command.x * width, y + command.y * height);
        break;
      case 'quad':
        ctx.quadraticCurveTo(
          x + command.cpx * width,
          y + command.cpy * height,
          x + command.x * width,
          y + command.y * height
        );
        break;
      case 'cubic':
        ctx.bezierCurveTo(
          x + command.cp1x * width,
          y + command.cp1y * height,
          x + command.cp2x * width,
          y + command.cp2y * height,
          x + command.x * width,
          y + command.y * height
        );
        break;
      case 'close':
        ctx.closePath();
        break;
    }
  }
}

function strokeCurrentPath(ctx: CanvasRenderingContext2D, stroke: Stroke): void {
  ctx.strokeStyle = stroke.color;
  ctx.lineWidth = stroke.width;
  ctx.setLineDash(stroke.dashed ? [Math.max(3, stroke.width * 2), Math.max(2, stroke.width)] : []);
  ctx.stroke();
}

function paintStyle(
  ctx: CanvasRenderingContext2D,
  paint: Paint,
  x: number,
  y: number,
  width: number,
  height: number
): string | CanvasGradient {
  if (paint.kind === 'solid') return paint.color;
  const radians = ((paint.angleDeg ?? 0) * Math.PI) / 180;
  const centerX = x + width / 2;
  const centerY = y + height / 2;
  const radius = Math.hypot(width, height) / 2;
  const gradient =
    paint.gradientType === 'linear'
      ? ctx.createLinearGradient(
          centerX - Math.cos(radians) * radius,
          centerY - Math.sin(radians) * radius,
          centerX + Math.cos(radians) * radius,
          centerY + Math.sin(radians) * radius
        )
      : ctx.createRadialGradient(centerX, centerY, 0, centerX, centerY, radius);
  for (const stop of paint.stops) gradient.addColorStop(Math.max(0, Math.min(1, stop.position)), stop.color);
  return gradient;
}

async function paintImage(
  ctx: CanvasRenderingContext2D,
  image: ImagePrimitive,
  resolver: CanvasImageResolver | undefined
): Promise<void> {
  if (image.assetId && resolver) {
    const source = await resolver(image.assetId);
    if (source) drawCropped(ctx, source, image);
  }
  if (image.stroke) {
    buildImageOutline(ctx, image);
    strokeCurrentPath(ctx, image.stroke);
  }
}

function paintTextBox(ctx: CanvasRenderingContext2D, textBox: TextBoxPrimitive): void {
  ctx.beginPath();
  ctx.rect(textBox.x, textBox.y, textBox.w, textBox.h);
  ctx.clip();
  ctx.textAlign = 'left';
  ctx.textBaseline = 'alphabetic';
  for (const line of textBox.lines) {
    for (const run of line.runs) paintTextRun(ctx, run, line.baseline);
  }
}

function paintTextRun(
  ctx: CanvasRenderingContext2D,
  run: PositionedTextRun,
  baseline: number
): void {
  const style = run.italic ? 'italic ' : '';
  const weight = run.bold ? 'bold ' : '';
  ctx.font = `${style}${weight}${run.fontSizePx}px ${quoteFamily(run.fontFamily)}`;
  ctx.fillStyle = run.color;
  const runBaseline = baseline - (run.baselineOffsetPx ?? 0);
  for (const chunk of positionedTextChunks(run)) {
    ctx.fillText(chunk.text, chunk.x, runBaseline);
  }
  if (run.underline) {
    ctx.fillRect(
      run.x,
      runBaseline + run.fontSizePx * 0.08,
      run.width,
      Math.max(1, run.fontSizePx * 0.05)
    );
  }
}

function positionedTextChunks(run: PositionedTextRun): Array<{ text: string; x: number }> {
  if (run.glyphs.length < 2) return [{ text: run.text, x: run.x }];
  const chunks: Array<{ text: string; x: number }> = [];
  let textStart = 0;
  let x = run.x;
  let expectedX = run.glyphs[0].x;
  for (const glyph of run.glyphs) {
    const offset = glyph.cluster - run.start;
    if (
      offset > textStart &&
      offset < run.text.length &&
      Math.abs(glyph.x - expectedX) > 0.0001
    ) {
      chunks.push({ text: run.text.slice(textStart, offset), x });
      textStart = offset;
      x = glyph.x;
    }
    expectedX = Math.fround(Math.fround(glyph.x) + Math.fround(glyph.advance));
  }
  chunks.push({ text: run.text.slice(textStart), x });
  return chunks;
}

function quoteFamily(family: string): string {
  return family.includes(' ') ? JSON.stringify(family) : family;
}

function paintPlaceholder(ctx: CanvasRenderingContext2D, placeholder: PlaceholderPrimitive): void {
  ctx.strokeStyle = '#8a94a6';
  ctx.lineWidth = 1;
  ctx.setLineDash([5, 4]);
  ctx.strokeRect(placeholder.x, placeholder.y, placeholder.w, placeholder.h);
  if (!placeholder.label) return;
  ctx.setLineDash([]);
  ctx.fillStyle = '#5d6675';
  ctx.font = '12px sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(
    placeholder.label,
    placeholder.x + placeholder.w / 2,
    placeholder.y + placeholder.h / 2,
    Math.max(0, placeholder.w - 12)
  );
}
