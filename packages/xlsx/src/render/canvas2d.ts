/**
 * Canvas2D backend for the display list — the browser paint target.
 *
 * DOM canvas types are used deliberately here; this file is the browser backend
 * and is excluded from the pure seam (viewport/ + display-list/). The function
 * is otherwise pure: it reads a plain display list and issues draw calls, with
 * no allocation of app state and no reads back from the canvas.
 */

import type { DisplayList, DrawCmd, LineCmd, TextCmd } from '../display-list/types';

const ALIGN_TO_TEXT_ALIGN: Record<NonNullable<TextCmd['align']>, CanvasTextAlign> = {
  left: 'left',
  center: 'center',
  right: 'right',
};

// dash patterns in device-independent px, matching the raster backend's stroke
// dashes so both targets read the same for a given line style.
const LINE_DASH: Record<'dashed' | 'dotted', number[]> = {
  dashed: [4, 2],
  dotted: [1, 2],
};

/**
 * Paint a display list into a 2D context. `dpr` maps device-independent list
 * coordinates onto the backing store, so callers size the canvas at
 * `width * dpr` × `height * dpr` and this sets the matching transform.
 */
export function paintDisplayList(
  ctx: CanvasRenderingContext2D,
  dl: DisplayList,
  dpr: number
): void {
  ctx.save();
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, dl.width, dl.height);
  for (const cmd of dl.commands) paintCommand(ctx, cmd);
  ctx.restore();
}

function paintCommand(ctx: CanvasRenderingContext2D, cmd: DrawCmd): void {
  switch (cmd.op) {
    case 'fillRect':
      ctx.fillStyle = cmd.color;
      ctx.fillRect(cmd.x, cmd.y, cmd.w, cmd.h);
      return;
    case 'line':
      paintLine(ctx, cmd);
      return;
    case 'text':
      paintText(ctx, cmd);
      return;
  }
}

// gridlines and borders. `dashed`/`dotted` set a dash pattern; `double` draws
// two thin parallel passes offset perpendicular to the (axis-aligned) line,
// matching the raster backend's double-border approximation.
function paintLine(ctx: CanvasRenderingContext2D, cmd: LineCmd): void {
  ctx.strokeStyle = cmd.color;
  if (cmd.style === 'double') {
    const off = Math.max(cmd.width * 0.8, 0.8);
    const horizontal = Math.abs(cmd.y1 - cmd.y2) <= Math.abs(cmd.x1 - cmd.x2);
    const [dx, dy] = horizontal ? [0, off] : [off, 0];
    ctx.lineWidth = Math.max(cmd.width * 0.6, 0.5);
    strokeSegment(ctx, cmd.x1 - dx, cmd.y1 - dy, cmd.x2 - dx, cmd.y2 - dy);
    strokeSegment(ctx, cmd.x1 + dx, cmd.y1 + dy, cmd.x2 + dx, cmd.y2 + dy);
    return;
  }
  ctx.lineWidth = cmd.width;
  ctx.save();
  ctx.setLineDash(cmd.style ? LINE_DASH[cmd.style] : []);
  strokeSegment(ctx, cmd.x1, cmd.y1, cmd.x2, cmd.y2);
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
  ctx.textAlign = ALIGN_TO_TEXT_ALIGN[cmd.align ?? 'left'];
  ctx.textBaseline = 'alphabetic';
  ctx.fillText(cmd.text, cmd.x, cmd.y);
  if (cmd.underline || cmd.strike) paintDecorations(ctx, cmd);
  ctx.restore();
}

// css font shorthand: `[italic] [bold] <size>px <family>`. Workbook fonts are
// often not installed on the host (Calibri on macOS); without a fallback the
// browser's unknown-family default is serif.
function fontString(cmd: TextCmd): string {
  const style = cmd.italic ? 'italic ' : '';
  const weight = cmd.bold ? 'bold ' : '';
  const family = cmd.fontFamily ? `"${cmd.fontFamily}", sans-serif` : 'sans-serif';
  return `${style}${weight}${cmd.fontSize}px ${family}`;
}

function paintDecorations(ctx: CanvasRenderingContext2D, cmd: TextCmd): void {
  const width = ctx.measureText(cmd.text).width;
  const align = cmd.align ?? 'left';
  const x0 = align === 'right' ? cmd.x - width : align === 'center' ? cmd.x - width / 2 : cmd.x;
  const thickness = Math.max(cmd.fontSize * 0.05, 0.5);
  ctx.fillStyle = cmd.color;
  if (cmd.underline) ctx.fillRect(x0, cmd.y + cmd.fontSize * 0.1, width, thickness);
  if (cmd.strike) ctx.fillRect(x0, cmd.y - cmd.fontSize * 0.26, width, thickness);
}
