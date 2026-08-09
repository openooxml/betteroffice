import type { Affine, ImagePrimitive, PageDisplayList, PagePrimitive, Paint, PlaceholderPrimitive, ShapePrimitive, Stroke, TextBoxPrimitive } from '../types';

export type CanvasImageResolver = (assetId: string) => CanvasImageSource | Promise<CanvasImageSource | null> | null;
export interface PaintPageOptions { resolveImage?: CanvasImageResolver; }
export interface PageCanvasLike { width: number; height: number; style: { width: string; height: string }; }
export function sizeCanvasForPage(canvas: PageCanvasLike, list: Pick<PageDisplayList, 'width' | 'height'>, dpr: number, scale = 1): void {
  canvas.width = Math.round(list.width * scale * dpr); canvas.height = Math.round(list.height * scale * dpr);
  canvas.style.width = `${list.width * scale}px`; canvas.style.height = `${list.height * scale}px`;
}
export async function paintPage(ctx: CanvasRenderingContext2D, list: PageDisplayList, dpr = 1, scale = 1, options: PaintPageOptions = {}): Promise<void> {
  if (list.contractVersion !== 3) throw new Error(`unsupported VSDX display-list contract version ${list.contractVersion}`);
  ctx.save();
  try { ctx.setTransform(dpr * scale, 0, 0, dpr * scale, 0, 0); ctx.clearRect(0, 0, list.width, list.height); for (const primitive of [...list.primitives].sort((a, b) => a.zOrder - b.zOrder)) await paintPrimitive(ctx, primitive, list.paintTransform, options, 0); }
  finally { ctx.restore(); }
}
async function paintPrimitive(ctx: CanvasRenderingContext2D, primitive: PagePrimitive, paintTransform: Affine, options: PaintPageOptions, depth: number): Promise<void> {
  if (depth >= 256) throw new Error('VSDX primitive nesting exceeds 256');
  ctx.save();
  try {
    const transform = 'transform' in primitive ? primitive.transform ?? identity() : identity(); ctx.transform(paintTransform.a, paintTransform.b, paintTransform.c, paintTransform.d, paintTransform.e, paintTransform.f); ctx.transform(transform.a, transform.b, transform.c, transform.d, transform.e, transform.f);
    switch (primitive.kind) {
      case 'shape': paintShape(ctx, primitive); break;
      case 'image': await paintImage(ctx, primitive, options.resolveImage); break;
      case 'textBox': paintTextBox(ctx, primitive); break;
      case 'placeholder': paintPlaceholder(ctx, primitive); break;
      case 'group': for (const child of [...primitive.primitives].sort((a, b) => a.zOrder - b.zOrder)) await paintPrimitive(ctx, child, identity(), options, depth + 1); break;
    }
  } finally { ctx.restore(); }
}
function identity(): Affine { return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 }; }
function paintShape(ctx: CanvasRenderingContext2D, shape: ShapePrimitive): void { ctx.beginPath(); for (const command of shape.path) { if (command.type === 'move') ctx.moveTo(Number(command.x), Number(command.y)); else if (command.type === 'line') ctx.lineTo(Number(command.x), Number(command.y)); else if (command.type === 'quad') ctx.quadraticCurveTo(Number(command.cpx), Number(command.cpy), Number(command.x), Number(command.y)); else if (command.type === 'cubic') ctx.bezierCurveTo(Number(command.cp1x), Number(command.cp1y), Number(command.cp2x), Number(command.cp2y), Number(command.x), Number(command.y)); else if (command.type === 'close') ctx.closePath(); } if (shape.fill) { ctx.fillStyle = paintStyle(ctx, shape.fill); ctx.fill(); } if (shape.stroke) stroke(ctx, shape.stroke); }
function paintStyle(ctx: CanvasRenderingContext2D, paint: Paint): string | CanvasGradient { if (paint.kind === 'solid') return paint.color; const gradient = ctx.createLinearGradient(0, 0, 1, 1); for (const stop of paint.stops) gradient.addColorStop(Math.max(0, Math.min(1, stop.position)), stop.color); return gradient; }
function stroke(ctx: CanvasRenderingContext2D, value: Stroke): void { ctx.strokeStyle = value.color; ctx.lineWidth = value.width; ctx.setLineDash(value.dashed ? [Math.max(3, value.width * 2), Math.max(2, value.width)] : []); ctx.stroke(); }
async function paintImage(ctx: CanvasRenderingContext2D, image: ImagePrimitive, resolve: CanvasImageResolver | undefined): Promise<void> { const source = resolve ? await resolve(image.assetId) : null; if (source) ctx.drawImage(source, image.x, image.y, image.width, image.height); }
function paintTextBox(ctx: CanvasRenderingContext2D, text: TextBoxPrimitive): void { ctx.beginPath(); ctx.rect(text.x, text.y, text.width, text.height); ctx.clip(); for (const paragraph of text.paragraphs) for (const run of paragraph.runs) { ctx.font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.sizeIn}px ${quote(run.family)}`; ctx.fillStyle = run.color; ctx.fillText(run.text, text.x, text.y); } }
function quote(family: string): string { return family.includes(' ') ? JSON.stringify(family) : family; }
function paintPlaceholder(ctx: CanvasRenderingContext2D, value: PlaceholderPrimitive): void { ctx.strokeStyle = '#8a94a6'; ctx.lineWidth = 1; ctx.setLineDash([5, 4]); ctx.strokeRect(value.x, value.y, value.width, value.height); ctx.setLineDash([]); ctx.fillStyle = '#5d6675'; ctx.font = '12px sans-serif'; ctx.fillText(value.reason, value.x + 6, value.y + 16, Math.max(0, value.width - 12)); }
