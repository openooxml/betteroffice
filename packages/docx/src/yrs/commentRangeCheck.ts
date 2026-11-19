/** Whether the comment ranges a save writes across stories come out whole in its final model. @internal */

import type { Comment } from '../types/content';
import type { BlockContent, Document, Paragraph } from '../types/document';
import { crossingWrittenWhole, withCrossingFallback } from './crossStoryComments';

type Kind = 'start' | 'end' | 'reference';

/** One comment mark of a save's model, in document order. @internal */
export interface ModelCommentMark {
  kind: Kind;
  /** The part holding the mark: the body, a header or footer, the footnotes or the endnotes. */
  part: string;
  order: number;
  /** The paragraph holding the mark as an item, or holding the run of a reference. */
  paragraph?: Paragraph;
  index?: number;
  /** A reference's place in its run. */
  inRun?: number;
  /** The markup holding the mark, as a carried drawing holds it, and where the mark starts and ends there. */
  markup?: string;
  from?: number;
  to?: number;
}

type Placed = Omit<ModelCommentMark, 'part' | 'order'>;

const KINDS = new Map<unknown, Kind>([
  ['commentRangeStart', 'start'],
  ['commentRangeEnd', 'end'],
  ['commentReference', 'reference'],
]);
const MARKUP_KINDS = new Map<string, Kind>([
  ['RangeStart', 'start'],
  ['RangeEnd', 'end'],
  ['Reference', 'reference'],
]);
const MARKUP = /<w:comment(RangeStart|RangeEnd|Reference)\b[^>]*?\bw:id=\\?"(\d+)\\?"[^>]*>/g;
const MARKUP_CONTENT = /<w:(?:t|tab|br|cr|sym|noBreakHyphen|softHyphen|delText|instrText|fldChar|fldSimple|drawing|pict|object|p)\b/;
const BETWEEN = new Set(['commentRangeStart', 'commentRangeEnd', 'bookmarkStart', 'bookmarkEnd']);
const CARRIERS = new Set(['opaqueDrawing', 'shape', 'drawing', 'chart']);

const parsedJson = (text: string): unknown => {
  try {
    return JSON.parse(text);
  } catch {
    return undefined;
  }
};

function inMarkup(text: string, id: number): Placed[] {
  const json = text.startsWith('{') || text.startsWith('[') ? parsedJson(text) : undefined;
  if (json !== undefined) return carried(json, id);
  if (!text.includes('<w:comment')) return [];
  const marks: Placed[] = [];
  for (const match of text.matchAll(MARKUP)) {
    if (Number(match[2]) !== id) continue;
    const from = match.index ?? 0;
    marks.push({ kind: MARKUP_KINDS.get(match[1]!)!, markup: text, from, to: from + match[0].length });
  }
  return marks;
}

/**
 * The marks of `id` that content carried as it is holds, counted in the one form of it that holds
 * the most: its markup, or the typed content beside it, never both.
 */
function carried(value: unknown, id: number): Placed[] {
  const strings: string[] = [];
  const typed: Placed[] = [];
  const visit = (node: unknown): void => {
    if (typeof node === 'string') {
      strings.push(node);
      return;
    }
    if (!node || typeof node !== 'object' || ArrayBuffer.isView(node) || node instanceof ArrayBuffer) return;
    if (Array.isArray(node)) {
      for (const child of node) visit(child);
      return;
    }
    const entry = node as { type?: unknown; id?: unknown; content?: unknown };
    if (entry.type === 'paragraph' && Array.isArray(entry.content)) {
      typed.push(...inParagraph(node as Paragraph, id));
      return;
    }
    const kind = KINDS.get(entry.type);
    if (kind && entry.id === id) typed.push({ kind });
    for (const child of Object.values(node)) visit(child);
  };
  visit(value);
  let best: Placed[] = [];
  for (const text of strings) {
    const marks = inMarkup(text, id);
    if (marks.length > best.length) best = marks;
  }
  return typed.length > best.length ? typed : best;
}

function inParagraph(paragraph: Paragraph, id: number): Placed[] {
  const marks: Placed[] = [];
  paragraph.content.forEach((item, index) => {
    if (item.type === 'commentRangeStart' || item.type === 'commentRangeEnd') {
      if (item.id === id) marks.push({ kind: KINDS.get(item.type)!, paragraph, index });
      return;
    }
    if (item.type !== 'run') {
      marks.push(...carried(item, id));
      return;
    }
    item.content.forEach((content, inRun) => {
      if (content.type === 'commentReference') {
        if (content.id === id) marks.push({ kind: 'reference', paragraph, index, inRun });
      } else if (CARRIERS.has(content.type)) {
        marks.push(...carried(content, id));
      }
    });
  });
  return marks;
}

function inBlocks(blocks: readonly BlockContent[], id: number): Placed[] {
  const marks: Placed[] = [];
  for (const block of blocks) {
    if (block.type === 'paragraph') marks.push(...inParagraph(block, id));
    else if (block.type === 'table') {
      for (const row of block.rows) for (const cell of row.cells) marks.push(...inBlocks(cell.content, id));
    } else if (block.type === 'blockSdt') marks.push(...inBlocks(block.content, id));
    else marks.push(...carried(block, id));
  }
  return marks;
}

/** Every mark of the comment `id` in `document`, part by part, in document order. @internal */
export function modelCommentMarks(document: Readonly<Document>, id: number): ModelCommentMark[] {
  const parts: Array<[string, readonly BlockContent[]]> = [['body', document.package.document.content]];
  for (const [rId, part] of document.package.headers ?? []) parts.push([`header:${rId}`, part.content]);
  for (const [rId, part] of document.package.footers ?? []) parts.push([`footer:${rId}`, part.content]);
  for (const note of document.package.footnotes ?? []) parts.push(['footnotes', note.content]);
  for (const note of document.package.endnotes ?? []) parts.push(['endnotes', note.content]);
  const marks: ModelCommentMark[] = [];
  for (const [part, content] of parts) {
    for (const mark of inBlocks(content, id)) marks.push({ ...mark, part, order: marks.length });
  }
  return marks;
}

/** Whether nothing but range marks and reference runs lies between `end` and `reference`. */
function referenceFollows(end: ModelCommentMark, reference: ModelCommentMark): boolean {
  if (end.markup !== undefined) {
    return (
      end.markup === reference.markup &&
      end.to! <= reference.from! &&
      !MARKUP_CONTENT.test(end.markup.slice(end.to, reference.from))
    );
  }
  const { paragraph } = end;
  if (!paragraph || paragraph !== reference.paragraph || reference.inRun === undefined) return false;
  const run = paragraph.content[reference.index!];
  if (end.index! >= reference.index! || run?.type !== 'run') return false;
  return (
    paragraph.content
      .slice(end.index! + 1, reference.index!)
      .every(
        (item) =>
          BETWEEN.has(item.type) ||
          (item.type === 'run' && item.content.every((content) => content.type === 'commentReference'))
      ) && run.content.slice(0, reference.inRun).every((content) => content.type === 'commentReference')
  );
}

/**
 * Whether `document`, a save's final model, holds the comment `id` well-formed: its marks all in one
 * part, exactly one start and one end, the start first, and at most one reference, right after the
 * end with nothing but range marks and reference runs between. Without a reference the end must sit
 * directly in a paragraph, where the serializer writes the reference right after it. It walks block
 * structure and comment marks, and searches strings only inside content carried as it is. @internal
 */
export function commentRangeWellFormed(document: Readonly<Document>, id: number): boolean {
  const marks = modelCommentMarks(document, id);
  if (new Set(marks.map((mark) => mark.part)).size !== 1) return false;
  const of = (kind: Kind) => marks.filter((mark) => mark.kind === kind);
  const [starts, ends, references] = [of('start'), of('end'), of('reference')];
  if (starts.length !== 1 || ends.length !== 1 || references.length > 1) return false;
  const [start, end] = [starts[0]!, ends[0]!];
  if (start.order >= end.order) return false;
  const reference = references[0];
  if (!reference) return end.paragraph !== undefined && end.inRun === undefined;
  return end.order < reference.order && referenceFollows(end, reference);
}

/** A save's model and the comments its projections wrote across stories, each with its root story. @internal */
export interface SaveModel {
  document: Document;
  whole: ReadonlyMap<number, string>;
}

/** `id` and the comments replying to it, at any depth. */
function threadOf(comments: readonly Comment[], id: number): number[] {
  const thread = [id];
  for (let at = 0; at < thread.length; at += 1) {
    for (const comment of comments) {
      if (comment.parentId === thread[at] && !thread.includes(comment.id)) thread.push(comment.id);
    }
  }
  return thread;
}

/**
 * The model `build` makes for a save, with no comment on the fallback. Where a comment written
 * across stories, or a reply on it, is not well-formed in it, `build` makes it once more with those
 * comments, each with the root story it starts in, written as the session holds them, and that
 * model is kept as it comes. @internal
 */
export function checkedSaveModel<T extends SaveModel>(build: (fallback: ReadonlyMap<number, string>) => T): T {
  const built = build(new Map());
  const comments = built.document.package.document.comments ?? [];
  const failed = new Map<number, string>();
  for (const [id, root] of built.whole) {
    const thread = threadOf(comments, id);
    if (thread.every((member) => commentRangeWellFormed(built.document, member))) continue;
    for (const member of thread) {
      const memberRoot = built.whole.get(member);
      if (memberRoot !== undefined) failed.set(member, memberRoot);
    }
    failed.set(id, root);
  }
  return failed.size === 0 ? built : build(failed);
}

/** The projection `project` makes of `session` for a save that serializes it as it is, checked. @internal */
export function checkedProjection(session: object, project: () => Document): Document {
  return checkedSaveModel((fallback) => {
    const document = withCrossingFallback(session, new Set(fallback.keys()), project);
    return { document, whole: crossingWrittenWhole(document) };
  }).document;
}
