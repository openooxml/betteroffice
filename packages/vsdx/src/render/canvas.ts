import type { Affine, PageDisplayList, PagePrimitive, Paint, PlaceholderPrimitive, ShapePrimitive, Stroke, TextBoxPrimitive, TextRun } from '../types';

export type CanvasImageResolver = (assetId: string) => CanvasImageSource | Promise<CanvasImageSource | null> | null;
export interface PaintPageOptions { resolveImage?: CanvasImageResolver; signal?: AbortSignal; }
export interface PageCanvasLike { width: number; height: number; style: { width: string; height: string }; }
export function sizeCanvasForPage(canvas: PageCanvasLike, list: Pick<PageDisplayList, 'width' | 'height'>, dpr: number, scale = 1): void {
  canvas.width = Math.round(list.width * scale * dpr); canvas.height = Math.round(list.height * scale * dpr);
  canvas.style.width = `${list.width * scale}px`; canvas.style.height = `${list.height * scale}px`;
}
export interface ModelPoint { x: number; y: number; }
export function canvasPointToModel(paintTransform: Affine, x: number, y: number, scale = 1): ModelPoint {
  const determinant = paintTransform.a * paintTransform.d - paintTransform.b * paintTransform.c;
  if (!Number.isFinite(determinant) || determinant === 0) throw new Error('VSDX paint transform is not invertible');
  if (!Number.isFinite(scale) || scale <= 0) throw new Error('VSDX canvas scale must be a positive number');
  const px = x / scale - paintTransform.e, py = y / scale - paintTransform.f;
  return { x: (paintTransform.d * px - paintTransform.c * py) / determinant + 0, y: (paintTransform.a * py - paintTransform.b * px) / determinant + 0 };
}
const paintRequests = new WeakMap<CanvasRenderingContext2D, object>();
export async function paintPage(ctx: CanvasRenderingContext2D, list: PageDisplayList, dpr = 1, scale = 1, options: PaintPageOptions = {}): Promise<void> {
  if (list.contractVersion !== 4) throw new Error(`unsupported VSDX display-list contract version ${list.contractVersion}`);
  const request = {};
  paintRequests.set(ctx, request);
  const images = new Map<string, CanvasImageSource | null>();
  const pending = new Map<string, Promise<void>>();
  const collect = (primitives: PagePrimitive[], depth: number) => {
    if (depth >= 256) throw new Error('VSDX primitive nesting exceeds 256');
    for (const primitive of primitives) {
      if (primitive.kind === 'group') collect(primitive.primitives, depth + 1);
      if (primitive.kind === 'image' && !pending.has(primitive.assetId)) {
        pending.set(primitive.assetId, Promise.resolve(options.resolveImage?.(primitive.assetId) ?? null).then(source => { images.set(primitive.assetId, source); }));
      }
    }
  };
  collect(list.primitives, 0);
  await Promise.all(pending.values());
  if (options.signal?.aborted || paintRequests.get(ctx) !== request) return;
  ctx.save();
  try { ctx.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0); ctx.clearRect(0, 0, list.width, list.height); for (const primitive of [...list.primitives].sort((a, b) => a.zOrder - b.zOrder)) paintPrimitive(ctx, primitive, list.paintTransform, images); }
  finally { ctx.restore(); }
}
function paintPrimitive(ctx: CanvasRenderingContext2D, primitive: PagePrimitive, paintTransform: Affine, images: Map<string, CanvasImageSource | null>): void {
  ctx.save();
  try {
    const transform = 'transform' in primitive ? primitive.transform ?? identity() : identity(); ctx.transform(paintTransform.a, paintTransform.b, paintTransform.c, paintTransform.d, paintTransform.e, paintTransform.f); ctx.transform(transform.a, transform.b, transform.c, transform.d, transform.e, transform.f);
    switch (primitive.kind) {
      case 'shape': paintShape(ctx, primitive); break;
      case 'image': { const source = images.get(primitive.assetId); if (source) { ctx.translate(0, 2 * primitive.y + primitive.height); ctx.scale(1, -1); ctx.drawImage(source, primitive.x, primitive.y, primitive.width, primitive.height); } break; }
      case 'textBox': paintTextBox(ctx, primitive); break;
      case 'placeholder': paintPlaceholder(ctx, primitive); break;
      case 'group': for (const child of [...primitive.primitives].sort((a, b) => a.zOrder - b.zOrder)) paintPrimitive(ctx, child, identity(), images); break;
    }
  } finally { ctx.restore(); }
}
function identity(): Affine { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; }
function paintShape(ctx: CanvasRenderingContext2D, shape: ShapePrimitive): void { ctx.beginPath(); for (const command of shape.path) { if (command.type === 'move') ctx.moveTo(Number(command.x), Number(command.y)); else if (command.type === 'line') ctx.lineTo(Number(command.x), Number(command.y)); else if (command.type === 'quad') ctx.quadraticCurveTo(Number(command.cpx), Number(command.cpy), Number(command.x), Number(command.y)); else if (command.type === 'cubic') ctx.bezierCurveTo(Number(command.cp1x), Number(command.cp1y), Number(command.cp2x), Number(command.cp2y), Number(command.x), Number(command.y)); else if (command.type === 'close') ctx.closePath(); } if (shape.fill) { ctx.fillStyle = paintStyle(ctx, shape.fill); ctx.fill(); } if (shape.stroke) stroke(ctx, shape.stroke); }
function paintStyle(ctx: CanvasRenderingContext2D, paint: Paint): string | CanvasGradient { if (paint.kind === 'solid') return paint.color; const gradient = ctx.createLinearGradient(0, 0, 1, 1); for (const stop of paint.stops) gradient.addColorStop(Math.max(0, Math.min(1, stop.position)), stop.color); return gradient; }
function stroke(ctx: CanvasRenderingContext2D, value: Stroke): void { ctx.strokeStyle = value.color; ctx.lineWidth = value.width; ctx.setLineDash(value.dashed ? [Math.max(3, value.width * 2), Math.max(2, value.width)] : []); ctx.stroke(); }
function paintTextBox(ctx: CanvasRenderingContext2D, text: TextBoxPrimitive): void {
  ctx.translate(0, 2 * text.y + text.height); ctx.scale(1, -1); ctx.textBaseline = 'top';
  ctx.beginPath(); ctx.rect(text.x, text.y, text.width, text.height); ctx.clip();
  let offset = 0;
  const runs = text.paragraphs.flatMap(paragraph => paragraph.runs.map(run => {
    const start = offset;
    offset += utf8Length(run.text);
    return { run, start, end: offset };
  }));
  for (const line of text.lines) for (const entry of runs) {
    const start = Math.max(line.start, entry.start), end = Math.min(line.end, entry.end);
    if (start >= end) continue;
    const x = line.caretStops.find(stop => stop.position === start)?.x ?? line.x;
    paintTextRun(ctx, entry.run, utf8Slice(entry.run.text, start - entry.start, end - entry.start), x, line.y);
  }
}
function paintTextRun(ctx: CanvasRenderingContext2D, run: TextRun, value: string, x: number, y: number): void { ctx.font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.sizeIn}px ${quote(run.family)}`; ctx.fillStyle = run.color; ctx.fillText(value, x, y); }
function utf8Length(value: string): number { return new TextEncoder().encode(value).byteLength; }
function utf8Slice(value: string, start: number, end: number): string { return new TextDecoder().decode(new TextEncoder().encode(value).slice(start, end)); }
function quote(family: string): string { return family.includes(' ') ? JSON.stringify(family) : family; }
function paintPlaceholder(ctx: CanvasRenderingContext2D, value: PlaceholderPrimitive): void { ctx.strokeStyle = '#8a94a6'; ctx.lineWidth = 1; ctx.setLineDash([5, 4]); ctx.strokeRect(value.x, value.y, value.width, value.height); ctx.setLineDash([]); ctx.fillStyle = '#5d6675'; ctx.font = '12px sans-serif'; ctx.fillText(value.reason, value.x + 6, value.y + 16, Math.max(0, value.width - 12)); }
