import { retainedFramePageById, type RetainedFrame } from '../layout/render/frameDelta';
import type { DisplayList, DisplayPage } from '../layout/render/displayList';
import { createDisplayListQueries, type DisplayListQueries, type DisplayListVisualLine } from '../layout/render/displayListQueries';
import { loadedRustDisplayListQueryEngine } from '../layout/render/rustDisplayList';
import { findWordBoundaries } from '../utils/textSelection';
import { DisplayPositionIndex } from './displayPositionIndex';
import { displayPositionToYrsLoc } from './inputPositionMap';
import { commandSelectionRange, readCommandSelection, resolveCommandSelection } from './residentCommandAnchors';
import type { YrsLoc, YrsStoryRange } from './index';
import type {
  DocxCommandApplyContext,
  DocxCommandIntent,
  DocxCommandSelection,
  ResidentDocxCommandEngine,
  ResidentDocxCommandState,
} from './residentCommandProtocol';

/** @internal */
export type DocxSelectionCommandIntent = Extract<
  DocxCommandIntent, { kind: 'navigate' | 'selectPoint' | 'extendSelection' }
>;

/** @internal Layout read after the predecessor command has resolved. */
export interface DocxCommandSelectionLayout {
  frame: RetainedFrame;
  layoutRevision: number;
}

/** @internal */
export interface ResolveDocxSelectionIntentInput {
  engine: ResidentDocxCommandEngine;
  state: ResidentDocxCommandState;
  command: DocxSelectionCommandIntent;
  context: DocxCommandApplyContext;
  readLayout?: () => DocxCommandSelectionLayout;
}

/** @internal Stale pointers always resolve or clamp; they cannot refuse a dependency. */
export type ResolveDocxSelectionIntentResult = {
  status: 'resolved' | 'clamped';
  selection: DocxCommandSelection;
  range: YrsStoryRange;
  goalX?: number;
};

interface SelectionReads {
  version: string;
  indexes: Map<string, DisplayPositionIndex>;
  positions: Map<string, Array<{ from: number; to: number; at: YrsLoc; length: number }>>;
  roots: Map<string, string>;
  parts: Map<string, number>;
  paragraphs: Map<string, {
    spans: ReturnType<ResidentDocxCommandEngine['geometryReader']['paragraphSpans']>;
    indices: Map<string, number>;
  }>;
  frame?: RetainedFrame;
  frameEpoch?: number;
  queries: Map<string, DisplayListQueries>;
}

const selectionReads = new WeakMap<ResidentDocxCommandEngine, SelectionReads>();

function readsFor(engine: ResidentDocxCommandEngine): SelectionReads {
  const reader = engine.geometryReader;
  const version = reader.version();
  let reads = selectionReads.get(engine);
  if (reads?.version !== version) {
    for (const queries of reads?.queries.values() ?? []) queries.dispose();
    reads = {
      version, indexes: new Map(), positions: new Map(),
      roots: new Map(), parts: reads?.parts ?? new Map(), paragraphs: new Map(), queries: new Map(),
    };
    selectionReads.set(engine, reads);
  }
  return reads;
}

function indexFor(reads: SelectionReads, engine: ResidentDocxCommandEngine, root: string): DisplayPositionIndex {
  let index = reads.indexes.get(root);
  if (!index) {
    if (reads.indexes.size >= 4) {
      const oldest = reads.indexes.keys().next().value!;
      reads.indexes.delete(oldest);
      reads.positions.delete(oldest);
    }
    index = new DisplayPositionIndex({
      ...engine.geometryReader, selectionText: engine.selectionText, resolveComment: engine.resolveComment,
    });
    reads.indexes.set(root, index);
  }
  return index;
}

function paragraphsFor(reads: SelectionReads, engine: ResidentDocxCommandEngine, story: string) {
  let paragraphs = reads.paragraphs.get(story);
  if (!paragraphs) {
    const spans = engine.geometryReader.paragraphSpans(story);
    paragraphs = { spans, indices: new Map(spans.map(({ paraId }, ordinal) => [paraId, ordinal])) };
    reads.paragraphs.set(story, paragraphs);
  }
  return paragraphs;
}

function paragraphText(reads: SelectionReads, engine: ResidentDocxCommandEngine, at: YrsLoc): string {
  const ordinal = paragraphsFor(reads, engine, at.story).indices.get(at.paraId);
  if (ordinal === undefined) throw new Error('Selection lacks paragraph');
  return (engine.paragraphSegments(at.story, [ordinal])[0] ?? []).map((segment) =>
    segment.kind === 'text' ? segment.text : segment.kind === 'embed' ? '\uFFFC' : ''
  ).join('');
}

function locationAt(reads: SelectionReads, engine: ResidentDocxCommandEngine, root: string, position: number): YrsLoc | null {
  const index = indexFor(reads, engine, root);
  const projection = index.projection(root);
  if (!projection) return null;
  const bounded = Math.max(0, Math.min(position, projection.size));
  const boundary = projection.blockBoundaryAt(bounded);
  if (boundary) return boundary;
  let positions = reads.positions.get(root);
  if (!positions) {
    positions = [];
    for (let from = 0; from <= projection.size;) {
      const node = projection.nodeAt(from);
      if (node?.kind === 'paragraph') {
        const paraId = String(node.attrs.paraId);
        const paragraphs = paragraphsFor(reads, engine, node.story);
        const ordinal = paragraphs.indices.get(paraId);
        if (ordinal === undefined) throw new Error('Projected selection lacks paragraph');
        const length = node.nodeSize - 2;
        const leading = paragraphs.spans[ordinal]!.length - length;
        positions.push({ from, to: from, at: { story: node.story, paraId, offset: Math.max(0, leading - 1) }, length: 0 });
        positions.push({ from: from + 1, to: from + node.nodeSize - 1, at: { story: node.story, paraId, offset: leading }, length });
        reads.roots.set(node.story, root);
        from += node.nodeSize;
      } else {
        const target = projection.targetAt(from);
        const map = index.inputMap(target.story);
        const at = map ? displayPositionToYrsLoc(map, target.displayPosition) : null;
        if (at) positions.push({ from, to: from, at, length: 0 });
        from++;
      }
    }
    reads.positions.set(root, positions);
  }
  let low = 0;
  let high = positions.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (positions[middle]!.from <= bounded) low = middle + 1;
    else high = middle;
  }
  const entry = positions[low - 1];
  return entry && bounded <= entry.to
    ? { ...entry.at, offset: entry.at.offset + Math.min(entry.length, bounded - entry.from) } : null;
}

function scopedList(list: DisplayList, root: string, pageIndex?: number): DisplayList {
  if (root === 'body') return list;
  const noteKind = root.startsWith('fn:') ? 'footnote' : 'endnote';
  return {
    ...list,
    pages: (root.startsWith('hf:')
      ? [list.pages.find((page) => !page.unbuilt && page.pageIndex === pageIndex &&
          [page.header, page.footer].some((region) => region?.rId === root.slice(3)))
        ?? list.pages.find((page) => !page.unbuilt &&
          [page.header, page.footer].some((region) => region?.rId === root.slice(3)))].filter((page): page is DisplayPage => page !== undefined)
      : list.pages).map((page) => ({
      ...page,
      primitives: root.startsWith('hf:')
        ? [page.header, page.footer].find((region) => region?.rId === root.slice(3))?.primitives ?? []
        : (page.noteAreas ?? []).filter((area) => (area.kind ?? 'footnote') === noteKind)
          .flatMap((area) => area.primitives ?? [])
          .filter((primitive) => primitive.groupId === `${noteKind}-${root.slice(3)}`),
      header: undefined, footer: undefined, noteAreas: undefined, positionSpan: undefined,
    })),
  };
}

function queriesFor(reads: SelectionReads, frame: RetainedFrame, root = 'body'): DisplayListQueries {
  if (reads.frame !== frame || reads.frameEpoch !== frame.frameEpoch) {
    for (const queries of reads.queries.values()) queries.dispose();
    reads.queries.clear();
    reads.frame = frame;
    reads.frameEpoch = frame.frameEpoch;
  }
  let queries = reads.queries.get(root);
  if (!queries) {
    if (reads.queries.size >= 4) {
      const oldest = reads.queries.keys().next().value!;
      reads.queries.get(oldest)!.dispose();
      reads.queries.delete(oldest);
    }
    queries = createDisplayListQueries(scopedList(frame.displayList, root, reads.parts.get(root)), loadedRustDisplayListQueryEngine() ?? undefined);
    reads.queries.set(root, queries);
  }
  return queries;
}

function lineAt(queries: DisplayListQueries, position: number): DisplayListVisualLine | null {
  const caret = queries.caretRect(position);
  if (!caret) return null;
  const center = caret.y + caret.height / 2;
  const distance = (line: DisplayListVisualLine) => Math.abs(line.y + line.height / 2 - center);
  let best: DisplayListVisualLine | null = null;
  for (const line of queries.visualLinesOnPage(caret.pageIndex)) {
    if (position < line.from || position > line.to) continue;
    if (!best || distance(line) < distance(best)) best = line;
  }
  return best;
}

/** @internal */
export function resolveDocxSelectionIntent(
  input: ResolveDocxSelectionIntentInput
): ResolveDocxSelectionIntentResult {
  const { engine, state, command, context } = input;
  const reader = engine.geometryReader;
  const previous = command.kind === 'navigate'
    ? resolveCommandSelection(engine, state, command, context)
    : readCommandSelection(engine, state, context.predecessorSelection ?? state.selection);
  const fallback = (): YrsLoc => {
    const story = reader.hasStory('body') ? 'body' : reader.storyIds().find((story) => reader.paragraphSpans(story).length);
    const first = story ? reader.paragraphSpans(story)[0] : undefined;
    if (!story || !first) throw new Error('Document has no valid caret');
    return { story, paraId: first.paraId, offset: 0 };
  };
  const base = previous.ok ? previous.selection : { anchor: fallback(), focus: fallback() };
  const preserve = (): ResolveDocxSelectionIntentResult => {
    const range = previous.ok ? previous.range : commandSelectionRange(engine, base);
    if (!range) throw new Error('Document has no valid selection');
    return { status: 'clamped', selection: base, range };
  };
  const reads = readsFor(engine);
  const rootFor = (at: YrsLoc): string => {
    const cached = reads.roots.get(at.story);
    if (cached) return cached;
    const roots = /^(body|(?:hf|fn|en):[^:]+)$/u.test(at.story)
      ? [at.story] : ['body', ...reader.storyIds().filter((story) => /^(hf|fn|en):/u.test(story))];
    for (const story of roots) {
      if (indexFor(reads, engine, story).projection(story)?.positionForLoc(at) != null) {
        if (reads.roots.size >= 32) reads.roots.clear();
        reads.roots.set(at.story, story);
        return story;
      }
    }
    return at.story;
  };
  const atDisplay = (root: string, position: number): YrsLoc | null => {
    return locationAt(reads, engine, root, position);
  };
  let clamped = !previous.ok;
  let goalX: number | undefined;
  let anchor: YrsLoc = base.anchor;
  let focus: YrsLoc = base.focus;
  if (command.kind === 'navigate') {
    const navigation = command.payload;
    const backward = navigation.direction === 'backward' || navigation.direction === 'up';
    const paragraphs = paragraphsFor(reads, engine, focus.story);
    const { spans, indices } = paragraphs;
    let paragraph = indices.get(focus.paraId) ?? -1;
    if (paragraph < 0) throw new Error('Navigation selection lacks paragraph');
    const range = commandSelectionRange(engine, base)!;
    const collapsed = range.start.paraId === range.end.paraId && range.start.offset === range.end.offset;
    if (!navigation.extend && !collapsed && (navigation.unit === 'character' || navigation.unit === 'word')) {
      focus = { story: range.story, ...(backward ? range.start : range.end) };
    } else if (navigation.unit === 'document-boundary') {
      const edge = backward ? spans[0]! : spans.at(-1)!;
      focus = { story: focus.story, paraId: edge.paraId, offset: backward ? 0 : edge.length };
    } else if (navigation.unit === 'line' && input.readLayout) {
      const layout = input.readLayout();
      const root = rootFor(focus);
      const queries = queriesFor(reads, layout.frame, root);
      const position = indexFor(reads, engine, root).positionOf(focus, root);
      if (position !== null && queries.isReady()) {
        if (navigation.direction === 'up' || navigation.direction === 'down') {
          const movement = queries.verticalMove(position, navigation.direction, navigation.goalX);
          if (movement) {
            focus = atDisplay(root, movement.position) ?? focus;
            goalX = movement.goalX;
          }
        } else {
          const line = lineAt(queries, position);
          if (line) focus = atDisplay(root, backward ? line.from : line.to) ?? focus;
        }
      }
    } else if (navigation.unit === 'line' || navigation.unit === 'paragraph') {
      if (navigation.direction === 'up' || navigation.direction === 'down') {
        paragraph = Math.max(0, Math.min(spans.length - 1, paragraph + (backward ? -1 : 1)));
        focus = { story: focus.story, paraId: spans[paragraph]!.paraId, offset: Math.min(focus.offset, spans[paragraph]!.length) };
        goalX = navigation.goalX;
      } else {
        focus = { story: focus.story, paraId: focus.paraId, offset: backward ? 0 : spans[paragraph]!.length };
      }
    } else {
      const text = paragraphText(reads, engine, focus);
      let offset = focus.offset;
      if (backward && offset === 0 && paragraph > 0) {
        paragraph--;
        offset = spans[paragraph]!.length;
      } else if (!backward && offset === spans[paragraph]!.length && paragraph + 1 < spans.length) {
        paragraph++;
        offset = 0;
      } else if (navigation.unit === 'word') {
        if (backward) {
          while (offset > 0 && /\s/u.test(text[offset - 1]!)) offset--;
          while (offset > 0 && !/\s/u.test(text[offset - 1]!)) offset--;
        } else {
          while (offset < text.length && !/\s/u.test(text[offset]!)) offset++;
          while (offset < text.length && /\s/u.test(text[offset]!)) offset++;
        }
      } else if (backward && offset > 0) {
        offset--;
        if (offset > 0 && /[\uDC00-\uDFFF]/u.test(text[offset]!) && /[\uD800-\uDBFF]/u.test(text[offset - 1]!)) offset--;
      } else if (!backward && offset < text.length) {
        offset += text.codePointAt(offset)! > 0xffff ? 2 : 1;
      }
      focus = { story: focus.story, paraId: spans[paragraph]!.paraId, offset };
    }
    if (navigation.extend && anchor.story !== focus.story) {
      focus = base.focus;
      clamped = true;
    }
    if (!navigation.extend) anchor = focus;
  } else {
    const point = command.payload;
    const layout = input.readLayout?.();
    const pages = layout?.frame.pages ?? [];
    const exact = layout && /^\d+$/u.test(point.page.pageId)
      ? retainedFramePageById(layout.frame, BigInt(point.page.pageId)) : undefined;
    const page = exact ?? pages[Math.max(0, Math.min(pages.length - 1, Math.trunc(point.page.pageIndex)))];
    clamped = !exact || (point.story !== undefined && !reader.hasStory(point.story));
    if (!page || !layout || page.page.unbuilt) return preserve();
    const queries = queriesFor(reads, layout.frame);
    if (!queries.isReady()) return preserve();
    const x = Math.max(0, Math.min(point.x, page.page.width));
    const y = !exact && point.page.pageIndex >= pages.length ? page.page.height : Math.max(0, Math.min(point.y, page.page.height));
    const hit = queries.hitTestRegions(page.pageIndex, x, y);
    const hitRoot = hit?.region === 'header' || hit?.region === 'footer' ? hit.rId ? `hf:${hit.rId}` : null
      : hit?.region === 'footnote' || hit?.region === 'endnote' ? hit.noteId === undefined ? null
        : `${hit.region === 'footnote' ? 'fn' : 'en'}:${hit.noteId}` : 'body';
    const projection = hitRoot ? indexFor(reads, engine, hitRoot).projection(hitRoot) : null;
    const at = hitRoot && projection && hit?.pos != null && Number.isFinite(hit.pos) && hit.pos >= 0 && hit.pos <= projection.size
      ? atDisplay(hitRoot, hit.pos) : null;
    if (at) {
      focus = at;
      if (hitRoot?.startsWith('hf:') && reads.parts.get(hitRoot) !== page.pageIndex) {
        reads.parts.set(hitRoot, page.pageIndex);
        reads.queries.get(hitRoot)?.dispose();
        reads.queries.delete(hitRoot);
      }
    } else {
      clamped = true;
      const lines = queries.visualLinesOnPage(page.pageIndex);
      if (exact && lines.length === 0) return preserve();
      const nearest = lines.reduce<typeof lines[number] | undefined>((best, line) => {
        const distance = (entry: typeof line) => Math.max(entry.y - y, 0, y - entry.y - entry.height);
        return !best || distance(line) < distance(best) ? line : best;
      }, undefined);
      const body = reader.hasStory('body') ? reader.paragraphSpans('body') : [];
      const edge = point.page.pageIndex >= pages.length ? body.at(-1) : body[0];
      focus = nearest ? atDisplay('body', x > nearest.x + nearest.width / 2 ? nearest.to : nearest.from) ?? fallback()
        : edge ? { story: 'body', paraId: edge.paraId, offset: point.page.pageIndex >= pages.length ? edge.length : 0 } : fallback();
    }
    if (command.kind === 'extendSelection' || point.modifiers.shift) {
      if (anchor.story !== focus.story) {
        focus = base.focus;
        clamped = true;
      }
    } else {
      anchor = focus;
      if (point.clickCount >= 2) {
        const text = paragraphText(reads, engine, focus);
        const [start, end] = point.clickCount >= 3 ? [0, text.length] : findWordBoundaries(text, focus.offset);
        anchor = { story: focus.story, paraId: focus.paraId, offset: start };
        focus = { story: focus.story, paraId: focus.paraId, offset: end };
      }
    }
  }
  let selection: DocxCommandSelection = { anchor, focus };
  let range = commandSelectionRange(engine, selection);
  if (!range) {
    focus = fallback();
    selection = { anchor: focus, focus };
    range = commandSelectionRange(engine, selection);
    clamped = true;
  }
  if (!range) throw new Error('Document has no valid selection');
  return { status: clamped ? 'clamped' : 'resolved', selection, range, ...(goalX === undefined ? {} : { goalX }) };
}
