import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import { drawGlyphRunPrimitive, GlyphCache, type GlyphRunPrimitive } from '@betteroffice/docx/layout/render';
import {
  indexEchoGlyphSet, leadingWordWidth, placeEchoGlyphs, type EchoGlyphIndex,
} from '@betteroffice/docx/peerless';
import type {
  DocxCommandEchoContext,
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxEchoGlyphSet,
  DocxEchoRect,
} from '@betteroffice/docx/yrs';
import { releaseBitmap, retainBitmap } from './bitmapOwnership';
import { decodeCommandFrameDeltas } from './frameDeltas';
import type { PeerlessTextIntent } from './inputEvents';

type GlyphSetEntry = { glyphs: EchoGlyphIndex; cache: GlyphCache };
type EchoLayout = { placement: NonNullable<ReturnType<typeof placeEchoGlyphs>>; cache: GlyphCache };

/** @internal Display-only drafts retain counts of original admissions after merging. */
export interface PeerlessEchoIntent {
  intent: PeerlessTextIntent;
  originalIntentCount: number;
  contextFrameEpoch: number;
}

/** @internal */
export interface PeerlessEchoPainterOptions {
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  /** Asks the worker to send everything again, here the glyph sets a context names but never arrived. */
  requestFullFrame?(): void;
}

/** @internal Painter state never updates selection or command targets. */
export interface PeerlessEchoPainter {
  installContext(context: DocxCommandEchoContext | undefined): void;
  restoreBackground(page: DocxCommandPageIdentity, backgroundId: string, rect: DocxEchoRect): boolean;
  preview(intent: PeerlessTextIntent): PeerlessEchoIntent | undefined;
  paint(pending: readonly EchoEntry<PeerlessEchoIntent>[]): void;
  drop(entries: readonly EchoEntry<PeerlessEchoIntent>[], reason: string): void;
  reconcile(
    frame: DocxCommandFrame,
    retired: readonly EchoEntry<PeerlessEchoIntent>[],
    remaining: readonly EchoEntry<PeerlessEchoIntent>[]
  ): void;
  disable(reason: string): void;
  dispose(): void;
}

/** @internal */
export function createEchoPainter(options: PeerlessEchoPainterOptions): PeerlessEchoPainter {
  const echoEncoder = new TextEncoder();
  let context: DocxCommandEchoContext | undefined;
  let staged: DocxCommandFrame['payload']['stagedPresentation'];
  let disabled = false;
  let disposed = false;
  let originalIntents = 0;
  let paintedPage: string | undefined;
  let paintedRect: { x: number; y: number; width: number; height: number } | undefined;
  let draft: EchoDraft | undefined;
  let baseline: { context: DocxCommandEchoContext; set: GlyphSetEntry; end: number } | undefined;
  let prepared: { context: DocxCommandEchoContext; set: GlyphSetEntry; text: string; value: EchoLayout } | undefined;
  let coverage = -1;
  const owner = {};
  const glyphSets = new Map<string, GlyphSetEntry>();
  const requested = new Set<string>();

  function receiveGlyphSets(sets: readonly DocxEchoGlyphSet[] | undefined): void {
    for (const set of sets ?? []) {
      glyphSets.set(set.key, { glyphs: indexEchoGlyphSet(set), cache: new GlyphCache({ provider: (fontId, glyphId) => {
        const outline = fontId === set.fontId ? set.outlines[glyphId] : undefined;
        if (outline === undefined) throw new Error('Echo glyph outline is unavailable');
        return outline;
      } }) });
      requested.delete(set.key);
    }
  }

  function contextGlyphSet(next: DocxCommandEchoContext) {
    const key = next.glyphs?.glyphSet;
    if (key === undefined) return undefined;
    const set = glyphSets.get(key);
    if (!set && !requested.has(key)) {
      requested.add(key);
      options.requestFullFrame?.();
    }
    return set;
  }

  function currentEnd(next: DocxCommandEchoContext, set: GlyphSetEntry): number | undefined {
    if (baseline?.context === next && baseline.set === set) return baseline.end;
    const line = next.glyphs;
    const placement = line && placeEchoGlyphs(set.glyphs, next.text, line.next, line.x);
    if (!placement) return undefined;
    baseline = { context: next, set, end: placement.end };
    return placement.end;
  }

  /** The draft line exactly as the engine will place it, or undefined when that cannot be known here. */
  function layout(text: string): EchoLayout | undefined {
    const line = context?.glyphs;
    const set = context && contextGlyphSet(context);
    if (!context || !line || !set) return undefined;
    if (!text && line.canEmpty === false) return undefined;
    if (line.next && (text.match(/ *$/)?.[0] !== context.text.match(/ *$/)?.[0] ||
      !text.endsWith(' ') && !line.next.startsWith(' '))) return undefined;
    if (prepared?.context === context && prepared.set === set && prepared.text === text) return prepared.value;
    const placement = placeEchoGlyphs(set.glyphs, text, line.next, line.x);
    const current = currentEnd(context, set);
    if (!placement || current === undefined) return undefined;
    if (placement.end > current && placement.end > line.right - 1) return undefined;
    if (placement.end < current && line.nextWord !== undefined && line.nextWord <= line.right - placement.end + 1) return undefined;
    if (line.previousSlack !== undefined && leadingWordWidth(text, placement.advances) <= line.previousSlack + 1) return undefined;
    try { for (const id of placement.glyphIds) set.cache.get(line.fontId, id); } catch { return undefined; }
    const value = { placement, cache: set.cache };
    prepared = { context, set, text, value };
    return value;
  }

  function restorePages(): void {
    for (const item of staged?.pages ?? []) {
      if (item.page.pageId !== paintedPage) continue;
      const canvas = options.canvasForPage(item.page);
      const ctx = canvas?.getContext('2d');
      if (!canvas || !ctx) continue;
      ctx.save();
      try {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        if (paintedRect && canvas.width === item.bitmap.width && canvas.height === item.bitmap.height) {
          const { x, y, width, height } = paintedRect;
          ctx.clearRect(x, y, width, height);
          ctx.drawImage(item.bitmap, x, y, width, height, x, y, width, height);
        } else {
          ctx.clearRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(item.bitmap, 0, 0, canvas.width, canvas.height);
        }
      } finally { ctx.restore(); }
    }
    paintedPage = undefined;
    paintedRect = undefined;
  }

  function closeStaged(keep?: DocxCommandFrame['payload']['stagedPresentation']): void {
    const retained = new Set<ImageBitmap>();
    for (const item of keep?.pages ?? []) {
      retained.add(item.bitmap);
      for (const background of item.backgrounds ?? []) retained.add(background.bitmap);
    }
    for (const bitmap of retained) retainBitmap(bitmap, owner);
    const bitmaps = new Set<ImageBitmap>();
    for (const item of staged?.pages ?? []) {
      bitmaps.add(item.bitmap);
      for (const background of item.backgrounds ?? []) bitmaps.add(background.bitmap);
    }
    for (const bitmap of bitmaps) if (!retained.has(bitmap)) releaseBitmap(bitmap, owner);
    staged = keep;
  }

  function prunePages(): void {
    if (!staged) return;
    closeStaged({ ...staged, pages: staged.pages.filter((item) => options.canvasForPage(item.page)) });
    if (!staged?.pages.some((item) => item.page.pageId === paintedPage)) {
      paintedPage = undefined;
      paintedRect = undefined;
    }
  }

  function installContext(next: DocxCommandEchoContext | undefined): void {
    if (disposed) return;
    context = next;
    baseline = undefined;
    prepared = undefined;
    draft = next ? contextDraft(next) : undefined;
    disabled = !next || next.shaping !== 'simple-ltr' || next.serializedBytes > 16 * 1024 ||
      echoEncoder.encode(JSON.stringify(next)).byteLength > 16 * 1024 || !contextGlyphSet(next);
    originalIntents = 0;
    if (disabled) restorePages();
    else if (next) currentEnd(next, contextGlyphSet(next)!);
  }

  /** Copies a staged background back onto its page canvas pixel for pixel. */
  function restoreBackground(page: DocxCommandPageIdentity, backgroundId: string, rect: DocxEchoRect): boolean {
    if (disposed) return false;
    prunePages();
    const item = staged?.pages.find((candidate) => candidate.page.pageId === page.pageId &&
      candidate.page.pageIndex === page.pageIndex);
    const background = item?.backgrounds?.find((candidate) => candidate.id === backgroundId);
    const canvas = options.canvasForPage(page);
    const ctx = canvas?.getContext('2d');
    if (!item || !background || !canvas || !ctx || canvas.width !== item.bitmap.width ||
      canvas.height !== item.bitmap.height || rect.width <= 0 || rect.height <= 0 ||
      rect.x < background.rect.x || rect.y < background.rect.y ||
      rect.x + rect.width > background.rect.x + background.rect.width ||
      rect.y + rect.height > background.rect.y + background.rect.height) return false;
    const ratio = background.bitmap.width / background.rect.width;
    const x = Math.round(background.rect.x * ratio);
    const y = Math.round(background.rect.y * ratio);
    if (paintedPage !== undefined && (paintedPage !== page.pageId || !paintedRect ||
      paintedRect.x !== x || paintedRect.y !== y || paintedRect.width !== background.bitmap.width ||
      paintedRect.height !== background.bitmap.height)) restorePages();
    ctx.save();
    try {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.clearRect(x, y, background.bitmap.width, background.bitmap.height);
      ctx.drawImage(background.bitmap, x, y);
      paintedPage = page.pageId;
      paintedRect = { x, y, width: background.bitmap.width, height: background.bitmap.height };
      return true;
    } finally { ctx.restore(); }
  }

  function disable(_reason: string): void {
    if (disposed) return;
    prunePages();
    restorePages();
    disabled = true;
  }

  function preview(intent: PeerlessTextIntent): PeerlessEchoIntent | undefined {
    if (!disposed) prunePages();
    if (disposed || disabled || !context || staged?.pages.find((item) =>
      item.page.pageId === context!.page.pageId && item.page.pageIndex === context!.page.pageIndex
    )?.frameEpoch !== context.frameEpoch) return undefined;
    if (intent.kind === 'paragraphBreak' ||
      (('text' in intent.payload) && !simpleText(intent.payload.text))) {
      disable('unsupported');
      return undefined;
    }
    const next = draft && applyIntent(draft, intent);
    if (!next || !layout(next.text)) {
      disable('unsupported');
      return undefined;
    }
    originalIntents += 1;
    if (originalIntents > 256) {
      disable('intent-cap');
      return undefined;
    }
    draft = next;
    return { intent, originalIntentCount: 1, contextFrameEpoch: context.frameEpoch };
  }

  function paint(pending: readonly EchoEntry<PeerlessEchoIntent>[]): void {
    if (!disposed) prunePages();
    if (disposed || disabled || !context) return;
    originalIntents = Math.max(originalIntents,
      pending.reduce((sum, entry) => sum + entry.value.originalIntentCount, 0));
    if (!Number.isSafeInteger(originalIntents) || originalIntents > 256 || pending.some(({ value }) =>
      !Number.isSafeInteger(value.originalIntentCount) || value.originalIntentCount < 1)) {
      disable('intent-cap');
      return;
    }
    draft = contextDraft(context);
    if (!pending.length) { restorePages(); return; }
    if (!draft) { disable('missing-boundary'); return; }
    for (const entry of pending) {
      const intent = entry.value.intent;
      if (intent.kind === 'paragraphBreak' || entry.value.contextFrameEpoch > context.frameEpoch) {
        disable('unsupported'); return;
      }
      const next = applyIntent(draft, intent);
      if (!next) { disable('unsupported'); return; }
      draft = next;
      if (echoEncoder.encode(draft.text).byteLength > 16 * 1024) { disable('context-cap'); return; }
    }
    const line = context.glyphs;
    const placed = line && layout(draft.text);
    const damage = context.damage[0];
    if (!line || !placed || !damage || context.damage.length !== 1) { disable('unsupported'); return; }
    if (!restoreBackground(context.page, damage.backgroundId, damage.rect)) {
      disable('missing-background'); return;
    }
    const canvas = options.canvasForPage(context.page)!;
    const ctx = canvas.getContext('2d')!;
    const glyphs = placed.placement;
    ctx.save();
    try {
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.beginPath();
      ctx.rect(line.device.x, line.device.y, line.device.width, line.device.height);
      ctx.clip();
      ctx.setTransform(line.pixelRatio, 0, 0, line.pixelRatio, 0, 0);
      if (glyphs.glyphIds.length) {
        const run: GlyphRunPrimitive = { kind: 'glyphRun', fontId: line.fontId, size: line.size, color: line.color,
          text: draft.text, glyphs: glyphs.glyphIds.map((id, index) => ({ id, x: glyphs.xs[index]!, y: line.y,
            cluster: 0, advance: glyphs.advances[index]! })) };
        drawGlyphRunPrimitive(ctx, run, { glyphCache: placed.cache });
      }
    } finally { ctx.restore(); }
  }

  return {
    installContext, restoreBackground, preview, paint, disable,
    drop(_entries, _reason) { disable('dropped'); },
    reconcile(frame, _retired, remaining) {
      if (disposed) return;
      const next = frame.payload.stagedPresentation;
      receiveGlyphSets(frame.payload.echoGlyphSets);
      const removed = new Set<string>();
      let fullPages: Set<string> | undefined;
      for (const delta of decodeCommandFrameDeltas(frame, true)) {
        if (delta.full && delta.operations.length) fullPages = new Set<string>();
        for (const operation of delta.operations) {
          if (operation.kind === 'remove') {
            removed.add(String(operation.pageId));
            fullPages?.delete(String(operation.pageId));
          }
          if (operation.kind === 'upsert') {
            removed.delete(String(operation.pageId));
            fullPages?.add(String(operation.pageId));
          }
        }
      }
      const replaced = new Set(next?.pages.map((item) => item.page.pageId));
      const reset = staged !== undefined && frame.revision <= staged.frameEpoch;
      const retainContext = !reset && frame.coversThroughSeq === coverage && context !== undefined &&
        !replaced.has(context.page.pageId) && frame.payload.echoContext === undefined;
      if (reset || (paintedPage !== undefined && replaced.has(paintedPage))) {
        paintedPage = undefined;
        paintedRect = undefined;
      }
      closeStaged(next ? { ...next, pages: [
        ...(reset ? [] : staged?.pages.filter((item) => !replaced.has(item.page.pageId) &&
          !removed.has(item.page.pageId) && (fullPages === undefined || fullPages.has(item.page.pageId)) &&
          options.canvasForPage(item.page) &&
          (frame.payload.layoutMeta === undefined || item.page.pageIndex < frame.payload.layoutMeta.pageCount)) ?? []),
        ...next.pages,
      ] } : staged);
      prunePages();
      if (!retainContext) installContext(frame.payload.echoContext);
      coverage = frame.coversThroughSeq;
      originalIntents = remaining.reduce((sum, entry) => sum + entry.value.originalIntentCount, 0);
      paint(remaining);
    },
    dispose() {
      if (disposed) return;
      restorePages();
      closeStaged();
      context = undefined;
      baseline = undefined;
      prepared = undefined;
      disposed = true;
    },
  };
}

type EchoDraft = { text: string; anchor: number; focus: number };

function contextDraft(context: DocxCommandEchoContext): EchoDraft | undefined {
  const anchor = context.boundaries.find((boundary) => boundary.grapheme === context.selection.anchorGrapheme)?.utf16;
  const focus = context.boundaries.find((boundary) => boundary.grapheme === context.selection.focusGrapheme)?.utf16;
  return anchor === undefined || focus === undefined ? undefined : { text: context.text, anchor, focus };
}

function applyIntent(draft: EchoDraft, intent: PeerlessTextIntent): EchoDraft | undefined {
  if (intent.kind === 'paragraphBreak') return undefined;
  let start = Math.min(draft.anchor, draft.focus);
  let end = Math.max(draft.anchor, draft.focus);
  let replacement = '';
  if (intent.kind === 'insertText' || intent.kind === 'replaceSelection') {
    replacement = intent.payload.text;
    if (!simpleText(replacement)) return undefined;
  } else if (start === end) {
    const boundaries = [0];
    for (const point of draft.text) boundaries.push(boundaries[boundaries.length - 1]! + point.length);
    const index = boundaries.indexOf(start);
    const count = intent.payload.count ?? 1;
    if (index < 0 || !Number.isSafeInteger(count) || count < 0) return undefined;
    if (intent.kind === 'deleteBackward') {
      if (count > index) return undefined;
      start = boundaries[index - count]!;
    } else {
      if (count > boundaries.length - 1 - index) return undefined;
      end = boundaries[index + count]!;
    }
  }
  const text = draft.text.slice(0, start) + replacement + draft.text.slice(end);
  return { text, anchor: start + replacement.length, focus: start + replacement.length };
}

function simpleText(text: string): boolean {
  return !/[\u0000-\u001f\u007f\u0590-\u08ff\u0900-\u109f\u200e-\u200f\u202a-\u202e\u2066-\u2069]/u.test(text);
}
