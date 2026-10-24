import type { PositionedTextRun } from '@betteroffice/pptx';
import type { CommandGeometryLine } from '@betteroffice/pptx/worker';
import { applyPrimitiveTransform, retainCommandFrameRaster, type CommandFrameLease } from './commandFramePainter';
import type {
  PeerlessActiveLine, PeerlessConfirmedFrame, PeerlessEcho, PeerlessPaintCheckpoint, PeerlessRect,
  PeerlessViewPainters, PeerlessViewSnapshot,
} from './peerlessViewStore';

interface EchoPlate {
  references: number;
  active: PeerlessActiveLine;
  bounds: PeerlessRect;
  background: HTMLCanvasElement;
  foreground: HTMLCanvasElement;
}

interface EchoRun {
  text: string;
  style: PositionedTextRun;
}

let measurementContext: CanvasRenderingContext2D | null = null;

export function echoLineClip(active: CommandGeometryLine): PeerlessRect {
  const shadow = active.box.textShadow;
  const spread = 2 + (shadow?.blur ?? 0) * 3 + Math.abs(shadow?.dx ?? 0) + Math.abs(shadow?.dy ?? 0);
  return { x: active.box.x - spread, y: active.line.y - spread,
    width: active.box.w + spread * 2, height: active.line.height + spread * 2 };
}

export function measureEchoPosition(active: PeerlessActiveLine, echo: PeerlessEcho, position: number): number | null {
  if (typeof document === 'undefined') return null;
  if (!measurementContext) measurementContext = document.createElement('canvas').getContext('2d');
  if (!measurementContext || typeof measurementContext.measureText !== 'function') return null;
  return active.line.x + measurePrefix(measurementContext, echoRuns(active, echo), position - active.line.start);
}

export function createLocalEchoPainter(
  canvas: HTMLCanvasElement, options?: { scale?: number; dpr?: number; current?(): boolean },
): Pick<PeerlessViewPainters, 'paintEcho' | 'restoreLine'> & { checkpoint(): PeerlessPaintCheckpoint; attach(target: HTMLCanvasElement): void; dispose(): void } {
  const scale = options?.scale ?? 1;
  const factor = scale * (options?.dpr ?? canvas.ownerDocument.defaultView?.devicePixelRatio ?? 1);
  const current = () => !disposed && (options?.current?.() ?? true);
  let disposed = false;
  let paintId = 0;
  let frame: PeerlessConfirmedFrame | undefined;
  let held: CommandFrameLease | undefined;
  let plate: EchoPlate | undefined;
  let latest: PeerlessViewSnapshot | undefined;
  let unsubscribe: (() => void) | undefined;
  let damaged: PeerlessRect[] = [];
  const checkpoints = new Set<() => void>();

  function releasePlate(value: EchoPlate | undefined): void {
    if (!value || --value.references > 0) return;
    value.background.width = value.foreground.width = 0;
    value.background.height = value.foreground.height = 0;
  }

  function subscribeRaster(): (() => void) | undefined {
    const candidate = held;
    if (!candidate) return;
    const refresh = () => {
      void candidate.ready.then(() => {
        if (!current() || held !== candidate || !latest) return;
        releasePlate(plate); plate = undefined;
        paint(latest);
      }).catch(() => {});
    };
    const off = candidate.raster.subscribe(refresh);
    if (!candidate.raster.ready) refresh();
    return off;
  }

  function prepare(view: PeerlessViewSnapshot): CanvasRenderingContext2D | null {
    if (!current() || !view.confirmed) return null;
    paintId += 1;
    latest = view;
    if (frame !== view.confirmed || !held) {
      unsubscribe?.();
      releasePlate(plate); plate = undefined;
      held?.release();
      frame = view.confirmed;
      held = retainCommandFrameRaster(frame, factor);
      damaged = [];
      unsubscribe = subscribeRaster();
    }
    const raster = held?.raster;
    if (!raster?.ready) return null;
    const active = view.activeLine;
    if (active && (!plate || plate.active.box !== active.box || plate.active.line !== active.line)) {
      releasePlate(plate);
      const bounds = activeBounds(active);
      const right = Math.ceil((bounds.x + bounds.width) * factor) / factor;
      const bottom = Math.ceil((bounds.y + bounds.height) * factor) / factor;
      bounds.x = Math.floor(bounds.x * factor) / factor;
      bounds.y = Math.floor(bounds.y * factor) / factor;
      bounds.width = right - bounds.x;
      bounds.height = bottom - bounds.y;
      plate = { references: 1, active, bounds, ...raster.plate(active, bounds) };
    }
    return canvas.getContext('2d');
  }

  function restore(ctx: CanvasRenderingContext2D, view: PeerlessViewSnapshot): void {
    const raster = held?.raster;
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    for (const bounds of damaged) {
      const x = Math.max(0, Math.floor(bounds.x * factor));
      const y = Math.max(0, Math.floor(bounds.y * factor));
      const width = Math.min(canvas.width, Math.ceil((bounds.x + bounds.width) * factor)) - x;
      const height = Math.min(canvas.height, Math.ceil((bounds.y + bounds.height) * factor)) - y;
      if (!raster || width <= 0 || height <= 0) continue;
      ctx.clearRect(x, y, width, height);
      ctx.drawImage(raster.canvas, x, y, width, height, x, y, width, height);
    }
    damaged = [];
    ctx.restore();
    latest = view;
  }

  function paint(view: PeerlessViewSnapshot): void {
    const ctx = prepare(view);
    if (!ctx) return;
    restore(ctx, view);
    ctx.save();
    ctx.setTransform(factor, 0, 0, factor, 0, 0);
    if (view.echo && plate) {
      const { bounds, active } = plate;
      ctx.clearRect(bounds.x, bounds.y, bounds.width, bounds.height);
      ctx.drawImage(plate.background, bounds.x, bounds.y, bounds.width, bounds.height);
      damaged.push(bounds);
      ctx.save();
      for (const primitive of active.transforms) {
        applyPrimitiveTransform(ctx, primitive);
        if (primitive.kind === 'chart' || primitive.kind === 'table') {
          ctx.beginPath();
          ctx.rect(primitive.x, primitive.y, primitive.w, primitive.h);
          ctx.clip();
        }
      }
      ctx.beginPath();
      const clip = echoLineClip(active);
      ctx.rect(clip.x, clip.y, clip.width, clip.height);
      ctx.clip();
      const runs = echoRuns(active, view.echo);
      const measure = (position: number) => measurePrefix(ctx, runs, position - active.line.start);
      let x = active.line.x;
      for (const run of runs) {
        setRunStyle(ctx, run.style);
        const shadow = active.box.textShadow;
        ctx.shadowColor = shadow?.color ?? 'transparent';
        ctx.shadowBlur = (shadow?.blur ?? 0) * factor;
        ctx.shadowOffsetX = (shadow?.dx ?? 0) * factor;
        ctx.shadowOffsetY = (shadow?.dy ?? 0) * factor;
        ctx.fillText(run.text, x, active.line.baseline - (run.style.baselineOffsetPx ?? 0));
        const width = ctx.measureText(run.text).width;
        if (run.style.underline) {
          ctx.fillRect(x, active.line.baseline + run.style.fontSizePx * 0.08,
            width, Math.max(1, run.style.fontSizePx * 0.05));
        }
        x += width;
      }
      ctx.shadowColor = 'transparent';
      ctx.shadowBlur = 0;
      ctx.shadowOffsetX = 0;
      ctx.shadowOffsetY = 0;
      if (view.composition?.selection) {
        const start = Math.min(view.composition.selection.anchor, view.composition.selection.focus);
        const left = measure(start);
        const right = measure(start + view.composition.text.length);
        ctx.fillStyle = '#111827';
        ctx.fillRect(active.line.x + Math.min(left, right), active.line.baseline + 2,
          Math.abs(right - left), 1);
      }
      ctx.restore();
      ctx.drawImage(plate.foreground, bounds.x, bounds.y, bounds.width, bounds.height);
    } else if (view.selection && view.selection.anchor !== view.selection.focus) {
      ctx.fillStyle = '#3b82f640';
      for (const rect of view.selectionRects) {
        if (rect.width <= 0 || rect.height <= 0) continue;
        ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
        damaged.push(rect);
      }
    }
    ctx.restore();
  }

  return {
    checkpoint() {
      const savedPaintId = paintId;
      const savedHeld = held?.retain();
      const saved = { frame, plate, latest, damaged: [...damaged] };
      if (saved.plate) saved.plate.references += 1;
      let settled = false;
      const finish = () => {
        if (settled) return;
        settled = true;
        checkpoints.delete(finish);
        savedHeld?.release();
        releasePlate(saved.plate);
      };
      checkpoints.add(finish);
      return Object.assign(() => {
        if (settled) return;
        try {
          if (!current() || paintId > savedPaintId + 1) return;
          paintId += 1;
          unsubscribe?.();
          held?.release();
          releasePlate(plate);
          ({ frame, plate, latest, damaged } = saved);
          if (plate) plate.references += 1;
          held = savedHeld?.retain();
          unsubscribe = subscribeRaster();
        } finally { finish(); }
      }, { commit: finish });
    },
    attach(target) { canvas = target; },
    paintEcho: paint,
    restoreLine(view) {
      const ctx = prepare(view);
      if (ctx) restore(ctx, view);
    },
    dispose() {
      if (disposed) return;
      try {
        if (current() && latest && damaged.length) {
          const ctx = canvas.getContext('2d');
          if (ctx) restore(ctx, latest);
        }
      } finally {
        disposed = true;
        for (const finish of checkpoints) finish();
        unsubscribe?.(); unsubscribe = undefined;
        releasePlate(plate); plate = undefined;
        held?.release(); held = undefined;
        frame = undefined;
        latest = undefined;
        damaged = [];
      }
    },
  };
}

function activeBounds(active: PeerlessActiveLine): PeerlessRect {
  const clip = echoLineClip(active);
  const points = [{ x: clip.x, y: clip.y }, { x: clip.x + clip.width, y: clip.y },
    { x: clip.x, y: clip.y + clip.height }, { x: clip.x + clip.width, y: clip.y + clip.height }];
  for (let i = active.transforms.length - 1; i >= 0; i -= 1) {
    const primitive = active.transforms[i];
    const cx = primitive.x + primitive.w / 2; const cy = primitive.y + primitive.h / 2;
    const angle = (primitive.transform?.rotationDeg ?? 0) * Math.PI / 180;
    const cos = Math.cos(angle); const sin = Math.sin(angle);
    for (const point of points) {
      const x = (point.x - cx) * (primitive.transform?.flipH ? -1 : 1);
      const y = (point.y - cy) * (primitive.transform?.flipV ? -1 : 1);
      point.x = cx + x * cos - y * sin;
      point.y = cy + x * sin + y * cos;
    }
  }
  const x = Math.min(...points.map((point) => point.x));
  const y = Math.min(...points.map((point) => point.y));
  return { x, y, width: Math.max(...points.map((point) => point.x)) - x,
    height: Math.max(...points.map((point) => point.y)) - y };
}

function setRunStyle(ctx: CanvasRenderingContext2D, run: PositionedTextRun): void {
  ctx.font = `${run.italic ? 'italic ' : ''}${run.bold ? 'bold ' : ''}${run.fontSizePx}px ${JSON.stringify(run.fontFamily)}`;
  ctx.fillStyle = run.color;
  ctx.letterSpacing = `${run.letterSpacingPx ?? 0}px`;
  ctx.textBaseline = 'alphabetic';
}

function echoRuns(active: PeerlessActiveLine, echo: PeerlessEcho): EchoRun[] {
  const original = active.text;
  const prefix = echo.insertionOffset;
  const suffix = original.length - prefix;
  const runs: EchoRun[] = [];
  let position = 0;
  const first = active.line.runs[0];
  if (!first) return runs;
  const append = (value: string, style: PositionedTextRun) => {
    if (!value) return;
    const last = runs[runs.length - 1];
    if (last?.style === style) last.text += value;
    else runs.push({ text: value, style });
  };
  for (const run of active.line.runs) {
    const end = position + run.text.length;
    append(original.slice(position, Math.min(end, prefix)), run);
    position = end;
    if (position >= prefix) break;
  }
  append(echo.text.slice(prefix, echo.text.length - suffix), echo.insertionStyle);
  position = 0;
  for (const run of active.line.runs) {
    const end = position + run.text.length;
    const start = Math.max(position, original.length - suffix);
    if (start < end) append(original.slice(start, end), run);
    position = end;
  }
  return runs;
}

function measurePrefix(ctx: CanvasRenderingContext2D, runs: readonly EchoRun[], length: number): number {
  let width = 0;
  let remaining = Math.max(0, length);
  for (const run of runs) {
    if (remaining <= 0) break;
    setRunStyle(ctx, run.style);
    width += ctx.measureText(run.text.slice(0, remaining)).width;
    remaining -= run.text.length;
  }
  return width;
}
