import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import { drawGlyphRunPrimitive, GlyphCache, type GlyphRunPrimitive } from '@betteroffice/docx/layout/render';
import {
  indexEchoGlyphSet, leadingWordWidth, placeEchoFaceGlyphs, placeEchoLine, type EchoGlyphIndex,
} from '@betteroffice/docx/worker';
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

type GlyphSetEntry = { glyphs: EchoGlyphIndex; cache: GlyphCache; set: DocxEchoGlyphSet };
type EchoLayout = { placement: NonNullable<ReturnType<typeof placeEchoLine>>; cache: GlyphCache;
  /** Face of each placed glyph and each face's set, for a line that mixes fonts. */
  faces?: { charFaces: readonly number[]; sets: readonly GlyphSetEntry[] } };

/** @internal Display-only drafts retain counts of original admissions after merging. */
export interface PeerlessEchoIntent {
  intent: PeerlessTextIntent;
  originalIntentCount: number;
  contextFrameEpoch: number;
}

/** @internal Page local caret predicted for the last painted pending entry. */
export interface PeerlessPredictedCaret {
  readonly throughSeq: number | undefined;
  pageIndex: number;
  rect: DocxEchoRect;
}

/** @internal */
export interface PeerlessEchoPainterOptions {
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  onPaint?(predicted?: PeerlessPredictedCaret): void;
  /** Asks the worker to send everything again, here the glyph sets a context names but never arrived. */
  requestFullFrame?(): void;
}

/** @internal Painter state never updates selection or command targets. */
export interface PeerlessEchoPainter {
  readonly contextThroughSeq?: number;
  installContext(context: DocxCommandEchoContext | undefined): void;
  installAckContext?(context: DocxCommandEchoContext, throughSeq: number): boolean;
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
  let prepared: { context: DocxCommandEchoContext; set: GlyphSetEntry; text: string; faces: string; value: EchoLayout } | undefined;
  /** The line metrics faces the context's line holds, which an edit must keep for the line height to hold. */
  let lineMetrics: string | undefined;
  let coverage = -1;
  const owner = {};
  const glyphSets = new Map<string, GlyphSetEntry>();
  /** Sets a grown set replaced, kept until no installed context names them. */
  const superseded = new Set<string>();
  const requested = new Set<string>();

  function receiveGlyphSets(sets: readonly DocxEchoGlyphSet[] | undefined): void {
    for (const incoming of sets ?? []) {
      let set: DocxEchoGlyphSet = incoming;
      if (incoming.extends !== undefined) {
        const base = glyphSets.get(incoming.extends)?.set;
        // without the set it grows, a context naming it asks for the whole set again
        if (!base) continue;
        const { extends: grown, ...rest } = incoming;
        set = { ...rest, kerning: [...base.kerning, ...incoming.kerning],
          unsafePairs: [...base.unsafePairs, ...incoming.unsafePairs], outlines: { ...base.outlines, ...incoming.outlines } };
        superseded.add(grown);
      }
      const outlines = set.outlines;
      glyphSets.set(set.key, { set, glyphs: indexEchoGlyphSet(set), cache: new GlyphCache({ provider: (fontId, glyphId) => {
        const outline = fontId === set.fontId ? outlines[glyphId] : undefined;
        if (outline === undefined) throw new Error('Echo glyph outline is unavailable');
        return outline;
      } }) });
      requested.delete(set.key);
    }
  }

  function clearPainted(): void {
    options.onPaint?.(undefined);
  }

  function predictedCaret(pending: readonly EchoEntry<PeerlessEchoIntent>[], placed: EchoLayout): PeerlessPredictedCaret | undefined {
    if (!context || !draft || draft.anchor !== draft.focus) return undefined;
    const last = pending.at(-1);
    const template = context.boundaries[0]?.caret;
    const damage = context.damage[0];
    const line = context.glyphs;
    if (!last || !template || !damage || !line) return undefined;
    let utf16 = 0;
    let x: number | undefined;
    const chars = [...draft.text];
    for (let index = 0; index < chars.length; index += 1) {
      if (utf16 === draft.focus) { x = placed.placement.xs[index]!; break; }
      utf16 += chars[index]!.length;
      if (utf16 > draft.focus) return undefined;
    }
    if (x === undefined) {
      if (utf16 !== draft.focus) return undefined;
      x = placed.placement.end;
    }
    if (x < damage.rect.x || x > damage.rect.x + damage.rect.width || x > line.right) return undefined;
    const ticket = last.ticket;
    return { get throughSeq() { return last.seq ?? ticket.seq; }, pageIndex: context.page.pageIndex,
      rect: { x, y: template.y, width: template.width, height: template.height } };
  }

  /** The glyph set of each face the context's line draws with; undefined until every one arrived. */
  function contextGlyphSets(next: DocxCommandEchoContext): GlyphSetEntry[] | undefined {
    const line = next.glyphs;
    if (!line) return undefined;
    const sets: GlyphSetEntry[] = [];
    for (const key of line.faces?.map((face) => face.glyphSet) ?? [line.glyphSet]) {
      const set = glyphSets.get(key);
      if (!set && !requested.has(key)) {
        requested.add(key);
        options.requestFullFrame?.();
      }
      if (!set) return undefined;
      sets.push(set);
    }
    return sets;
  }

  function contextGlyphSet(next: DocxCommandEchoContext) {
    return contextGlyphSets(next)?.[0];
  }

  /** Places `text` whose code points take `faces` on the context's line. */
  function place(next: DocxCommandEchoContext, sets: readonly GlyphSetEntry[], text: string, faces: readonly number[] | undefined) {
    const line = next.glyphs!;
    if (!line.faces) return placeEchoLine(sets[0]!.glyphs, text, line.next, line.x, line.center);
    if (!faces || line.shapingSize === undefined) return undefined;
    return placeEchoFaceGlyphs({ sets: sets.map((set) => set.glyphs), charFaces: faces,
      ...(line.nextFace === undefined ? {} : { nextFace: line.nextFace }),
      ...(line.letterSpacing === undefined ? {} : { letterSpacing: line.letterSpacing }),
      ...(line.autoSpaceOff === undefined ? {} : { autoSpaceOff: line.autoSpaceOff }),
      size: line.shapingSize }, text, line.next, line.x);
  }

  function currentEnd(next: DocxCommandEchoContext, set: GlyphSetEntry): number | undefined {
    if (baseline?.context === next && baseline.set === set) return baseline.end;
    const sets = contextGlyphSets(next);
    const placement = next.glyphs && sets && place(next, sets, next.text, next.glyphs.charFaces);
    if (!placement) return undefined;
    baseline = { context: next, set, end: placement.end };
    return placement.end;
  }

  function metricsOf(faces: readonly number[]): string | undefined {
    const line = context?.glyphs;
    if (!line?.faces) return undefined;
    return [...new Set(faces.map((face) => line.faces![face]!.metrics))].sort((a, b) => a - b).join(',');
  }

  /** The draft line exactly as the engine will place it, or undefined when that cannot be known here. */
  function layout(text: string, faces?: readonly number[]): EchoLayout | undefined {
    const line = context?.glyphs;
    const sets = context && contextGlyphSets(context);
    const set = sets?.[0];
    if (!context || !line || !sets || !set) return undefined;
    if (!text && line.canEmpty === false) return undefined;
    if (line.center !== undefined && text.endsWith(' ')) return undefined;
    // the first glyph of a continuation line shapes and kerns with the end of the line above
    if (line.previousSlack !== undefined && text[0] !== context.text[0]) return undefined;
    // a line wrapped inside a word rewraps as the word grows or shrinks; one wrapped between East
    // Asian characters keeps its break only while its last character stays (a break after it may not be)
    if (line.next && (text.match(/ *$/)?.[0] !== context.text.match(/ *$/)?.[0] ||
      !text.endsWith(' ') && !line.next.startsWith(' ') &&
      (/^[A-Za-z0-9]/.test(line.next) || [...text].at(-1) !== [...context.text].at(-1)))) return undefined;
    // the line keeps its height only while it holds the line metrics faces it had
    if (line.faces && (!faces || metricsOf(faces) !== lineMetrics)) return undefined;
    const faceKey = faces?.join(',') ?? '';
    if (prepared?.context === context && prepared.set === set && prepared.text === text && prepared.faces === faceKey) {
      return prepared.value;
    }
    const placement = place(context, sets, text, faces);
    const current = currentEnd(context, set);
    if (!placement || current === undefined) return undefined;
    if (placement.end > current && placement.end > line.right - 1) return undefined;
    let ink = line.x;
    [...text].forEach((char, index) => { if (char !== ' ') ink = placement.xs[index]! + placement.advances[index]!; });
    if (ink > line.right - 1) return undefined;
    if (placement.end < current && line.nextWord !== undefined && line.nextWord <= line.right - placement.end + 1) return undefined;
    if (line.previousSlack !== undefined && leadingWordWidth(text, placement.advances) <= line.previousSlack + 1) return undefined;
    try {
      placement.glyphIds.forEach((id, index) => {
        const face = faces?.[index] ?? 0;
        sets[face]!.cache.get(line.faces?.[face]?.fontId ?? line.fontId, id);
      });
    } catch { return undefined; }
    const value: EchoLayout = { placement, cache: set.cache, ...(line.faces && faces ? { faces: { charFaces: faces, sets } } : {}) };
    prepared = { context, set, text, faces: faceKey, value };
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
    lineMetrics = draft?.faces ? metricsOf(draft.faces) : undefined;
    const named = new Set(next?.glyphs?.faces?.map((face) => face.glyphSet) ?? [next?.glyphs?.glyphSet]);
    for (const key of superseded) {
      if (named.has(key)) continue;
      glyphSets.delete(key);
      superseded.delete(key);
    }
    disabled = !next || next.shaping !== 'simple-ltr' || next.serializedBytes > 16 * 1024 ||
      echoEncoder.encode(JSON.stringify(next)).byteLength > 16 * 1024 || !contextGlyphSet(next);
    originalIntents = 0;
    if (disabled) { restorePages(); clearPainted(); }
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
    clearPainted();
  }

  function preview(intent: PeerlessTextIntent): PeerlessEchoIntent | undefined {
    if (!disposed) prunePages();
    if (disposed || disabled || !context || staged?.pages.find((item) =>
      item.page.pageId === context!.page.pageId && item.page.pageIndex === context!.page.pageIndex
    )?.frameEpoch !== context.frameEpoch) return undefined;
    if (intent.kind === 'paragraphBreak' || intent.kind === 'lineBreak' ||
      (('text' in intent.payload) && !simpleText(intent.payload.text))) {
      disable('unsupported');
      return undefined;
    }
    const next = draft && applyIntent(draft, intent, context.glyphs);
    if (!next || !layout(next.text, next.faces)) {
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
    if (disposed) return;
    if (disabled || !context) { clearPainted(); return; }
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
      if (intent.kind === 'paragraphBreak' || intent.kind === 'lineBreak' || entry.value.contextFrameEpoch > context.frameEpoch) {
        disable('unsupported'); return;
      }
      const next = applyIntent(draft, intent, context.glyphs);
      if (!next) { disable('unsupported'); return; }
      draft = next;
      if (echoEncoder.encode(draft.text).byteLength > 16 * 1024) { disable('context-cap'); return; }
    }
    const line = context.glyphs;
    const placed = line && layout(draft.text, draft.faces);
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
      // a table cell's text paints inside its clip, as the page raster clips it
      if (line.clip) {
        ctx.beginPath();
        ctx.rect(line.clip.x, line.clip.y, line.clip.w, line.clip.h);
        ctx.clip();
      }
      const chars = [...draft.text];
      // one run per stretch of characters that take one face
      for (let from = 0; from < glyphs.glyphIds.length;) {
        const face = placed.faces?.charFaces[from] ?? 0;
        let to = from + 1;
        while (to < glyphs.glyphIds.length && (placed.faces?.charFaces[to] ?? 0) === face) to += 1;
        const run: GlyphRunPrimitive = { kind: 'glyphRun', fontId: line.faces?.[face]?.fontId ?? line.fontId,
          size: line.size, color: line.color, text: chars.slice(from, to).join(''),
          glyphs: glyphs.glyphIds.slice(from, to).map((id, index) => ({ id, x: glyphs.xs[from + index]!, y: line.y,
            cluster: 0, advance: glyphs.advances[from + index]! })) };
        drawGlyphRunPrimitive(ctx, run, { glyphCache: placed.faces?.sets[face]?.cache ?? placed.cache });
        from = to;
      }
    } finally { ctx.restore(); }
    const predicted = predictedCaret(pending, placed);
    options.onPaint?.(predicted);
  }

  return {
    get contextThroughSeq() { return context && !disabled ? coverage : undefined; },
    installContext, restoreBackground, preview, paint, disable,
    installAckContext(next, throughSeq) {
      if (disposed || staged?.pages.find((item) => item.page.pageId === next.page.pageId &&
        item.page.pageIndex === next.page.pageIndex)?.frameEpoch !== next.frameEpoch) return false;
      installContext(next);
      if (disabled) return false;
      coverage = throughSeq;
      return true;
    },
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
      clearPainted();
      context = undefined;
      baseline = undefined;
      prepared = undefined;
      disposed = true;
    },
  };
}

/**
 * The text a pending edit leaves on the echoed line. On a line that mixes fonts, `faces` holds each
 * code point's face and `own` whether it sits in the run typing goes into (typed characters do).
 */
type EchoDraft = { text: string; anchor: number; focus: number; faces?: number[]; own?: boolean[] };

function contextDraft(context: DocxCommandEchoContext): EchoDraft | undefined {
  const anchor = context.boundaries.find((boundary) => boundary.grapheme === context.selection.anchorGrapheme)?.utf16;
  const focus = context.boundaries.find((boundary) => boundary.grapheme === context.selection.focusGrapheme)?.utf16;
  if (anchor === undefined || focus === undefined) return undefined;
  const line = context.glyphs;
  if (!line?.faces) return { text: context.text, anchor, focus };
  const faces = line.charFaces ? [...line.charFaces] : [];
  if (faces.length !== [...context.text].length) return undefined;
  const own: boolean[] = [];
  let utf16 = 0;
  for (const point of context.text) {
    own.push(line.typedRun !== undefined && utf16 >= line.typedRun[0] && utf16 < line.typedRun[1]);
    utf16 += point.length;
  }
  return { text: context.text, anchor, focus, faces, own };
}

function applyIntent(draft: EchoDraft, intent: PeerlessTextIntent, line?: DocxCommandEchoContext['glyphs']): EchoDraft | undefined {
  if (intent.kind === 'paragraphBreak' || intent.kind === 'lineBreak') return undefined;
  let start = Math.min(draft.anchor, draft.focus);
  let end = Math.max(draft.anchor, draft.focus);
  let replacement = '';
  const boundaries = [0];
  for (const point of draft.text) boundaries.push(boundaries[boundaries.length - 1]! + point.length);
  if (intent.kind === 'insertText' || intent.kind === 'replaceSelection') {
    replacement = intent.payload.text;
    if (!simpleText(replacement)) return undefined;
  } else if (start === end) {
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
  const next: EchoDraft = { text, anchor: start + replacement.length, focus: start + replacement.length };
  if (!draft.faces || !draft.own) return next;
  // typed characters take the faces the engine routes them to in the run left of the caret
  const from = boundaries.indexOf(start);
  const to = boundaries.indexOf(end);
  if (from < 0 || to < 0) return undefined;
  const faces: number[] = [];
  for (const char of replacement) {
    const face = line?.typed?.findIndex((chars) => chars.includes(char)) ?? -1;
    if (face < 0) return undefined;
    faces.push(face);
  }
  if (faces.length && (from === 0 || !draft.own[from - 1])) return undefined;
  next.faces = [...draft.faces.slice(0, from), ...faces, ...draft.faces.slice(to)];
  next.own = [...draft.own.slice(0, from), ...faces.map(() => true), ...draft.own.slice(to)];
  return next;
}

function simpleText(text: string): boolean {
  return !/[\u0000-\u001f\u007f\u0590-\u08ff\u0900-\u109f\u200e-\u200f\u202a-\u202e\u2066-\u2069]/u.test(text);
}
