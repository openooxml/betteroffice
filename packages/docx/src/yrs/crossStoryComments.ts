/** The comment ranges of an opened document that start in one story and end in another. @internal */

import type { BlockContent, Document, Paragraph, ParagraphContent } from '../types/document';
import type { YrsSession } from './index';
import { headerFooterStory } from './sessionInternals';
import { seededCellIndices } from './tableRowSpans';

/** @internal */
export type CommentMarkKind = 'start' | 'end' | 'reference';
type Kind = CommentMarkKind;

/** Where one comment mark sits in the source document. @internal */
export interface SourceCommentMark {
  story: string;
  /** The root story holding `story`: the body, a header or footer, or a note. */
  part: string;
  /** The mark's place among all marks of its part, in document order. */
  order: number;
  /** Whether the mark lies in a drawing's own markup, as in a text box, rather than in a story. */
  carried: boolean;
  /** Whether the mark sits directly in its paragraph: a range mark, or a reference first in a run of it. */
  direct: boolean;
  paragraph?: Paragraph;
  index?: number;
  carrier?: object;
}

/** A source comment whose one start, one end and one reference make a range across stories. @internal */
export interface CrossingComment {
  start: SourceCommentMark;
  end: SourceCommentMark;
  reference: SourceCommentMark;
}

const KINDS = new Map<unknown, Kind>([
  ['commentRangeStart', 'start'],
  ['commentRangeEnd', 'end'],
  ['commentReference', 'reference'],
]);
const XML_KINDS = new Map<string, Kind>([
  ['RangeStart', 'start'],
  ['RangeEnd', 'end'],
  ['Reference', 'reference'],
]);
const XML_MARKS = /<w:comment(RangeStart|RangeEnd|Reference)\b[^>]*?\bw:id=\\?"(\d+)\\?"/g;
const BETWEEN_END_AND_REFERENCE = new Set(['commentRangeStart', 'commentRangeEnd', 'bookmarkStart', 'bookmarkEnd']);
const CARRIERS = new Set(['opaqueDrawing', 'shape', 'drawing', 'chart']);

/** Every comment mark `value` holds at any depth, as objects or as markup in its strings. */
function marksIn(value: unknown): Array<[Kind, number]> {
  const found = new Map<string, [Kind, number]>();
  const visit = (node: unknown): void => {
    if (typeof node === 'string') {
      for (const match of node.matchAll(XML_MARKS)) {
        const kind = XML_KINDS.get(match[1]!)!;
        found.set(`${kind}:${match[2]}`, [kind, Number(match[2])]);
      }
      return;
    }
    if (!node || typeof node !== 'object' || ArrayBuffer.isView(node) || node instanceof ArrayBuffer) return;
    if (Array.isArray(node)) {
      for (const child of node) visit(child);
      return;
    }
    const entry = node as { type?: unknown; id?: unknown };
    const kind = KINDS.get(entry.type);
    if (kind && typeof entry.id === 'number') found.set(`${kind}:${entry.id}`, [kind, entry.id]);
    for (const child of Object.values(node)) visit(child);
  };
  visit(value);
  return [...found.values()];
}

/** Whether `value` holds the comment mark at any depth, in an object, in markup or in a JSON string. */
export function holdsCommentMark(value: unknown, kind: CommentMarkKind, id: number): boolean {
  if (typeof value === 'string') {
    if (marksIn(value).some(([held, heldId]) => held === kind && heldId === id)) return true;
    if (!value.startsWith('{') && !value.startsWith('[')) return false;
    try {
      return holdsCommentMark(JSON.parse(value), kind, id);
    } catch {
      return false;
    }
  }
  if (!value || typeof value !== 'object' || ArrayBuffer.isView(value) || value instanceof ArrayBuffer) return false;
  const entry = value as { type?: unknown; id?: unknown };
  if (KINDS.get(entry.type) === kind && entry.id === id) return true;
  return Object.values(value).some((child) => holdsCommentMark(child, kind, id));
}

/** Whether only range marks and reference runs lie between the end and the reference of its paragraph. */
function referenceFollowsEnd(end: SourceCommentMark, reference: SourceCommentMark): boolean {
  if (!end.direct || !reference.direct || end.paragraph !== reference.paragraph) return false;
  if (end.index === undefined || reference.index === undefined || end.index >= reference.index) return false;
  return end.paragraph!.content
    .slice(end.index + 1, reference.index)
    .every((item: ParagraphContent) =>
      BETWEEN_END_AND_REFERENCE.has(item.type) ||
      (item.type === 'run' && item.content.every((content) => content.type === 'commentReference'))
    );
}

/** Whether the end and the reference are carried by the same drawing. */
const carriedTogether = (end: SourceCommentMark, reference: SourceCommentMark): boolean =>
  end.carried && reference.carried && end.carrier === reference.carrier;

/**
 * The comments of `document`, as opened, whose range starts in one story and ends in another (or
 * in a drawing's own markup), keyed by comment id. A comment is listed only when the source holds
 * exactly one start, one end and one reference of it, all three in one part, the start before the
 * end, and the reference directly after the end in its paragraph or in the drawing that carries
 * the end.
 */
export function crossStoryComments(
  session: YrsSession,
  document: Readonly<Document>
): Map<number, CrossingComment> {
  const marks = new Map<number, Record<Kind, SourceCommentMark[]>>();
  const add = (kind: Kind, id: number, mark: SourceCommentMark): void => {
    const entry = marks.get(id) ?? { start: [], end: [], reference: [] };
    entry[kind].push(mark);
    marks.set(id, entry);
  };
  let order = 0;
  const seen = new Set<string>();

  const paragraph = (story: string, part: string, block: Paragraph): void => {
    block.content.forEach((item, index) => {
      const place = { story, part, carried: false, direct: true, paragraph: block, index };
      const kind = KINDS.get(item.type);
      if (kind && kind !== 'reference' && typeof (item as { id?: unknown }).id === 'number') {
        add(kind, (item as { id: number }).id, { ...place, order: ++order });
        return;
      }
      if (item.type !== 'run') {
        for (const [held, id] of marksIn(item)) add(held, id, { ...place, order: ++order, direct: false });
        return;
      }
      const contents = item.content;
      contents.forEach((content, at) => {
        if (content.type === 'commentReference') {
          const first = contents.slice(0, at).every((before) => before.type === 'commentReference');
          if (content.id !== undefined) add('reference', content.id, { ...place, order: ++order, direct: first });
        } else if (CARRIERS.has(content.type)) {
          for (const [held, id] of marksIn(content)) {
            add(held, id, { ...place, order: ++order, carried: true, direct: false, carrier: content });
          }
        }
      });
    });
  };

  const blocks = (story: string, part: string, content: readonly BlockContent[]): void => {
    if (seen.has(story)) return;
    seen.add(story);
    let tableIndex = 0;
    let sdtIndex = 0;
    for (const block of content) {
      if (block.type === 'paragraph') {
        paragraph(story, part, block);
      } else if (block.type === 'blockSdt') {
        blocks(`${story}:sdt${sdtIndex++}`, part, block.content);
      } else if (block.type === 'table') {
        const table = tableIndex++;
        seededCellIndices(block).forEach((cells, rowIndex) => {
          cells.forEach((cellIndex, ordinal) => {
            blocks(`${story}:t${table}:r${rowIndex}c${ordinal}`, part, block.rows[rowIndex]!.cells[cellIndex]!.content);
          });
        });
      } else {
        for (const [held, id] of marksIn(block)) {
          add(held, id, { story, part, order: ++order, carried: false, direct: false });
        }
      }
    }
  };

  const root = (story: string, content: readonly BlockContent[]): void => blocks(story, story, content);
  root('body', document.package.document.content);
  for (const [rId, part] of document.package.headers ?? []) root(headerFooterStory(session, rId), part.content);
  for (const [rId, part] of document.package.footers ?? []) root(headerFooterStory(session, rId), part.content);
  for (const note of document.package.footnotes ?? []) root(`fn:${note.id}`, note.content);
  for (const note of document.package.endnotes ?? []) root(`en:${note.id}`, note.content);

  const crossing = new Map<number, CrossingComment>();
  for (const [id, found] of marks) {
    const once = found.start.length === 1 && found.end.length === 1 && found.reference.length === 1;
    if (!once) continue;
    const [start, end, reference] = [found.start[0]!, found.end[0]!, found.reference[0]!];
    const acrossStories = start.story !== end.story || start.carried || end.carried;
    const onePart = start.part === end.part && end.part === reference.part;
    const startFirst = start.order < end.order;
    const startHeld = start.direct || start.carried;
    const referenceAfterEnd = referenceFollowsEnd(end, reference) || carriedTogether(end, reference);
    if (acrossStories && onePart && startFirst && startHeld && referenceAfterEnd) {
      crossing.set(id, { start, end, reference });
    }
  }
  return crossing;
}

const NO_FALLBACK: ReadonlySet<number> = new Set();
const NONE_WHOLE: ReadonlyMap<number, string> = new Map();
const fallbacks = new WeakMap<object, ReadonlySet<number>>();
const writtenWhole = new WeakMap<object, ReadonlyMap<number, string>>();

/**
 * What `project` returns, its projections of `session` writing the comments `ids` as the session
 * holds them, cut at the end of the paragraph they start in, rather than across stories. @internal
 */
export function withCrossingFallback<T>(session: object, ids: ReadonlySet<number>, project: () => T): T {
  const previous = fallbacks.get(session);
  fallbacks.set(session, ids);
  try {
    return project();
  } finally {
    if (previous) fallbacks.set(session, previous);
    else fallbacks.delete(session);
  }
}

/** The comments a projection of `session` writes as the session holds them. @internal */
export const crossingFallback = (session: object): ReadonlySet<number> => fallbacks.get(session) ?? NO_FALLBACK;

/** Records the comments the projection `document` wrote across stories, each with the root story it starts in. @internal */
export function noteWrittenWhole(document: object, whole: ReadonlyMap<number, string>): void {
  if (whole.size > 0) writtenWhole.set(document, whole);
}

/** The comments the projection `document` wrote across stories, each with the root story it starts in. @internal */
export const crossingWrittenWhole = (document: object): ReadonlyMap<number, string> =>
  writtenWhole.get(document) ?? NONE_WHOLE;
