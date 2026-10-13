import type { EchoEntry } from '../../../../../../shared/office-session/commands/types';
import { decodeFrameDelta } from '@betteroffice/docx/layout/render';
import { closeCommandBitmap } from '@betteroffice/docx/peerless';
import type {
  DocxCommandEchoContext,
  DocxCommandFrame,
  DocxCommandPageIdentity,
  DocxEchoRect,
} from '@betteroffice/docx/yrs';
import type { PeerlessTextIntent } from './inputEvents';

/** @internal Display-only drafts retain counts of original admissions after merging. */
export interface PeerlessEchoIntent {
  intent: PeerlessTextIntent;
  originalIntentCount: number;
  contextFrameEpoch: number;
}

/** @internal */
export interface PeerlessEchoPainterOptions {
  canvasForPage(page: DocxCommandPageIdentity): HTMLCanvasElement | null;
  canEchoFont?(family: string): boolean;
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
  let context: DocxCommandEchoContext | undefined;
  let staged: DocxCommandFrame['payload']['stagedPresentation'];
  let disabled = false;
  let disposed = false;
  let originalIntents = 0;
  let paintedPage: string | undefined;
  let draft: EchoDraft | undefined;
  let coverage = -1;
  const closed = new WeakSet<ImageBitmap>();
  const pageSizes = new Map<string, readonly [number, number]>();

  function restorePages(): void {
    for (const item of staged?.pages ?? []) {
      if (item.page.pageId !== paintedPage) continue;
      const canvas = options.canvasForPage(item.page);
      const ctx = canvas?.getContext('2d');
      if (!canvas || !ctx) continue;
      ctx.save();
      try {
        ctx.setTransform(1, 0, 0, 1, 0, 0);
        ctx.clearRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(item.bitmap, 0, 0, canvas.width, canvas.height);
      } finally { ctx.restore(); }
    }
    paintedPage = undefined;
  }

  function closeStaged(keep?: DocxCommandFrame['payload']['stagedPresentation']): void {
    const retained = new Set<ImageBitmap>();
    for (const item of keep?.pages ?? []) {
      retained.add(item.bitmap);
      for (const background of item.backgrounds ?? []) retained.add(background.bitmap);
    }
    const bitmaps = new Set<ImageBitmap>();
    for (const item of staged?.pages ?? []) {
      bitmaps.add(item.bitmap);
      for (const background of item.backgrounds ?? []) bitmaps.add(background.bitmap);
    }
    for (const bitmap of bitmaps) if (!retained.has(bitmap) && !closed.has(bitmap)) {
      closed.add(bitmap);
      closeCommandBitmap(bitmap);
    }
    staged = keep;
    const pages = new Set(keep?.pages.map((item) => item.page.pageId));
    for (const id of pageSizes.keys()) if (!pages.has(id)) pageSizes.delete(id);
  }

  function prunePages(): void {
    if (!staged) return;
    closeStaged({ ...staged, pages: staged.pages.filter((item) => options.canvasForPage(item.page)) });
    if (!staged?.pages.some((item) => item.page.pageId === paintedPage)) paintedPage = undefined;
  }

  function installContext(next: DocxCommandEchoContext | undefined): void {
    if (disposed) return;
    context = next;
    draft = next ? contextDraft(next) : undefined;
    disabled = !next || next.shaping !== 'simple-ltr' || next.serializedBytes > 16 * 1024 ||
      new TextEncoder().encode(JSON.stringify(next)).byteLength > 16 * 1024;
    originalIntents = 0;
    if (disabled) restorePages();
  }

  function restoreBackground(page: DocxCommandPageIdentity, backgroundId: string, rect: DocxEchoRect): boolean {
    if (disposed) return false;
    prunePages();
    const item = staged?.pages.find((candidate) => candidate.page.pageId === page.pageId &&
      candidate.page.pageIndex === page.pageIndex);
    const background = item?.backgrounds?.find((candidate) => candidate.id === backgroundId);
    const canvas = options.canvasForPage(page);
    const ctx = canvas?.getContext('2d');
    if (!item || !background || !canvas || !ctx || rect.width <= 0 || rect.height <= 0 ||
      rect.x < background.rect.x || rect.y < background.rect.y ||
      rect.x + rect.width > background.rect.x + background.rect.width ||
      rect.y + rect.height > background.rect.y + background.rect.height) return false;
    const sx = background.bitmap.width / background.rect.width;
    const sy = background.bitmap.height / background.rect.height;
    const size = pageSizes.get(page.pageId);
    const dx = canvas.width / (size?.[0] ?? item.bitmap.width);
    const dy = canvas.height / (size?.[1] ?? item.bitmap.height);
    ctx.save();
    try {
      ctx.setTransform(dx, 0, 0, dy, 0, 0);
      ctx.clearRect(rect.x, rect.y, rect.width, rect.height);
      ctx.drawImage(background.bitmap, (rect.x - background.rect.x) * sx,
        (rect.y - background.rect.y) * sy, rect.width * sx, rect.height * sy,
        rect.x, rect.y, rect.width, rect.height);
      paintedPage = page.pageId;
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
    if (context?.styleRuns.some(run => options.canEchoFont?.(run.fontFamily) === false)) disable('worker-font');
    if (disposed || disabled || !context || staged?.pages.find((item) =>
      item.page.pageId === context!.page.pageId && item.page.pageIndex === context!.page.pageIndex
    )?.frameEpoch !== context.frameEpoch) return undefined;
    if (intent.kind === 'paragraphBreak' ||
      (('text' in intent.payload) && !simpleText(intent.payload.text))) {
      disable('unsupported');
      return undefined;
    }
    const next = draft && applyIntent(draft, intent);
    if (!next) {
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
    if (context?.styleRuns.some(run => options.canEchoFont?.(run.fontFamily) === false)) disable('worker-font');
    if (disposed || disabled || !context) return;
    originalIntents = Math.max(originalIntents,
      pending.reduce((sum, entry) => sum + entry.value.originalIntentCount, 0));
    if (!Number.isSafeInteger(originalIntents) || originalIntents > 256 || pending.some(({ value }) =>
      !Number.isSafeInteger(value.originalIntentCount) || value.originalIntentCount < 1)) {
      disable('intent-cap');
      return;
    }
    restorePages();
    draft = contextDraft(context);
    if (!pending.length) return;
    const style = context.styleRuns[0];
    if (!style || context.styleRuns.some((run) => !sameStyle(style, run))) {
      disable('unsupported-style');
      return;
    }
    if (!draft) { disable('missing-boundary'); return; }
    for (const entry of pending) {
      const intent = entry.value.intent;
      if (intent.kind === 'paragraphBreak' || entry.value.contextFrameEpoch > context.frameEpoch) {
        disable('unsupported'); return;
      }
      const next = applyIntent(draft, intent);
      if (!next) { disable('unsupported'); return; }
      draft = next;
      if (new TextEncoder().encode(draft.text).byteLength > 16 * 1024) { disable('context-cap'); return; }
    }
    const text = draft.text;
    const item = staged?.pages.find((candidate) => candidate.page.pageId === context!.page.pageId);
    const damage = context.damage[0];
    const background = item?.backgrounds?.find((candidate) => candidate.id === damage?.backgroundId);
    const canvas = options.canvasForPage(context.page);
    const ctx = canvas?.getContext('2d');
    if (!item || !damage || !background || !canvas || !ctx || context.damage.length !== 1) {
      disable('missing-background'); return;
    }
    if (!restoreBackground(context.page, damage.backgroundId, damage.rect)) {
      disable('missing-background'); return;
    }
    let reflow = false;
    const size = pageSizes.get(item.page.pageId);
    ctx.save();
    try {
      ctx.setTransform(canvas.width / (size?.[0] ?? item.bitmap.width),
        0, 0, canvas.height / (size?.[1] ?? item.bitmap.height), 0, 0);
      ctx.beginPath();
      ctx.rect(damage.rect.x, damage.rect.y, damage.rect.width, damage.rect.height);
      ctx.clip();
      ctx.transform(...context.transform);
      ctx.transform(...style.transform);
      ctx.font = `${style.italic ? 'italic ' : ''}${style.bold ? 'bold ' : ''}${style.fontSizePx}px ${style.fontFamily}`;
      ctx.fillStyle = style.color;
      ctx.textBaseline = 'alphabetic';
      ctx.textAlign = 'left';
      ctx.direction = 'ltr';
      const width = ctx.measureText(text).width;
      if (style.baseline.x + width > damage.rect.x + damage.rect.width) {
        reflow = true;
      } else {
        ctx.fillText(text, style.baseline.x, style.baseline.y);
        for (const y of [style.underline ? style.baseline.y + 1 : null,
          style.strike ? style.baseline.y - style.fontSizePx * 0.3 : null]) {
          if (y !== null) ctx.fillRect(style.baseline.x, y, width, Math.max(1, style.fontSizePx / 16));
        }
      }
    } finally { ctx.restore(); }
    if (reflow) disable('reflow');
  }

  return {
    installContext, restoreBackground, preview, paint, disable,
    drop(_entries, _reason) { disable('dropped'); },
    reconcile(frame, _retired, remaining) {
      if (disposed) return;
      const next = frame.payload.stagedPresentation;
      const removed = new Set<string>();
      let fullPages: Set<string> | undefined;
      const sizes = new Map<string, readonly [number, number]>();
      for (const bytes of [...(frame.payload.pageFrames ?? []), frame.payload.frame]) {
        if (!bytes.byteLength) continue;
        const delta = decodeFrameDelta(bytes);
        if (delta.full && delta.operations.length) fullPages = new Set<string>();
        for (const operation of delta.operations) {
          if (operation.kind === 'remove') {
            removed.add(String(operation.pageId));
            fullPages?.delete(String(operation.pageId));
          }
          if (operation.kind === 'upsert') {
            removed.delete(String(operation.pageId));
            fullPages?.add(String(operation.pageId));
            sizes.set(String(operation.pageId), [operation.page.width, operation.page.height]);
          }
        }
      }
      const replaced = new Set(next?.pages.map((item) => item.page.pageId));
      const reset = staged !== undefined && frame.revision <= staged.frameEpoch;
      const retainContext = !reset && frame.coversThroughSeq === coverage && context !== undefined &&
        !replaced.has(context.page.pageId) && frame.payload.echoContext === undefined;
      if (reset || (paintedPage !== undefined && replaced.has(paintedPage))) paintedPage = undefined;
      closeStaged(next ? { ...next, pages: [
        ...(reset ? [] : staged?.pages.filter((item) => !replaced.has(item.page.pageId) &&
          !removed.has(item.page.pageId) && (fullPages === undefined || fullPages.has(item.page.pageId)) &&
          options.canvasForPage(item.page) &&
          (frame.payload.layoutMeta === undefined || item.page.pageIndex < frame.payload.layoutMeta.pageCount)) ?? []),
        ...next.pages,
      ] } : staged);
      for (const item of next?.pages ?? []) {
        const meta = frame.payload.layoutMeta?.pageSizes;
        const size = meta ? [meta[item.page.pageIndex * 2], meta[item.page.pageIndex * 2 + 1]] : sizes.get(item.page.pageId);
        if (size && size[0] !== undefined && size[1] !== undefined && size[0] > 0 && size[1] > 0) {
          pageSizes.set(item.page.pageId, [size[0], size[1]]);
        }
      }
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

function sameStyle(a: DocxCommandEchoContext['styleRuns'][number], b: DocxCommandEchoContext['styleRuns'][number]): boolean {
  return a.fontFamily === b.fontFamily && a.fontSizePx === b.fontSizePx && a.color === b.color &&
    a.bold === b.bold && a.italic === b.italic && a.underline === b.underline && a.strike === b.strike &&
    a.baseline.y === b.baseline.y && a.transform.every((value, index) => value === b.transform[index]);
}
