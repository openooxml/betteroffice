import type {
  GeometryPathCommand,
  ImagePrimitive,
  Paint,
  PlaceholderPrimitive,
  PositionedTextRun,
  ShapePrimitive,
  SlideDisplayList,
  SlidePrimitive,
  Stroke,
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
    }
  } finally {
    ctx.restore();
  }
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
  if (shape.stroke) strokeCurrentPath(ctx, shape.stroke);
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
    if (source) ctx.drawImage(source, image.x, image.y, image.w, image.h);
  }
  if (image.stroke) {
    ctx.beginPath();
    ctx.rect(image.x, image.y, image.w, image.h);
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
  ctx.fillText(run.text, run.x, baseline);
  if (run.underline) {
    ctx.fillRect(run.x, baseline + run.fontSizePx * 0.08, run.width, Math.max(1, run.fontSizePx * 0.05));
  }
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
