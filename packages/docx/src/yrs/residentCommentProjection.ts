import type { Comment } from '../types/content';
import type { DocxEditFailureCode } from './edits';
import type { YrsCommentInfo, YrsStoryRange } from './index';
import type { ResidentEngineSession } from './residentEngineSession';

export interface ResidentCommentProjection {
  comments: Comment[];
  ids: Record<string, number>;
  sessionIds: Map<number, string>;
}

function commentContent(body: unknown, fallback: Comment['content']): Comment['content'] {
  if (Array.isArray(body) && body.every(entry => entry?.type === 'paragraph')) return body;
  return typeof body === 'string' ? [{ type: 'paragraph', formatting: {}, content: [{ type: 'run', formatting: {}, content: [{ type: 'text', text: body }] }] }] : fallback;
}

/**
 * The IDs a save gives the session's comments, in synced comment order: a source comment keeps its
 * ID, and any other comment keeps its own ID when that is a free OOXML comment ID and takes the next
 * free one after the source IDs otherwise.
 */
export function savedCommentIds(sourceIds: readonly number[], native: readonly Pick<YrsCommentInfo, 'id'>[]): Map<string, number> {
  const used = new Set(sourceIds);
  const keys = new Set(sourceIds.map(String));
  const ids = new Map(native.filter(comment => keys.has(comment.id)).map(comment => [comment.id, Number(comment.id)]));
  let next = Math.max(0, ...used) + 1;
  for (const comment of native) {
    if (ids.has(comment.id)) continue;
    let id = /^\d+$/u.test(comment.id) ? Number(comment.id) : Number.NaN;
    if (!Number.isSafeInteger(id) || used.has(id)) {
      while (used.has(next)) next += 1;
      id = next;
    }
    used.add(id);
    ids.set(comment.id, id);
  }
  return ids;
}

/**
 * The comments a save writes: the source package's, each with the fields the session wrote over its
 * seeded placeholders and without the ones it removed, then the comments added in the session.
 * With `initialized`, native resolved state replaces the source state.
 * Session comments keep the IDs `aliases` handed out, in the order this replica first projected them;
 * with `fresh`, they take {@link savedCommentIds} instead, so replicas holding the same comments save
 * the same IDs.
 */
export function projectResidentComments(engine: ResidentEngineSession, aliases: Map<string, number>, initialized = false, fresh = false): ResidentCommentProjection {
  const source = engine.propertyDocument().package.document.comments ?? [];
  const native = engine.listComments();
  for (const comment of native) {
    const id = /^\d+$/u.test(comment.id) ? Number(comment.id) : Number.NaN;
    if (Number.isSafeInteger(id)) aliases.set(comment.id, id);
  }
  if (!fresh) {
    let next = 1;
    for (const id of [...source.map(comment => comment.id), ...aliases.values()]) next = Math.max(next, id + 1);
    for (const comment of native) if (!aliases.has(comment.id)) { aliases.set(comment.id, next); next++; }
  }
  const numbers = fresh ? savedCommentIds(source.map(comment => comment.id), native) : aliases;
  const ids = Object.fromEntries(native.map(comment => [comment.id, numbers.get(comment.id)!]));
  const held = new Map(native.map(comment => [numbers.get(comment.id)!, comment]));
  const parentOf = (comment: YrsCommentInfo) => comment.parentId === null ? undefined
    : numbers.get(comment.parentId) ?? (/^\d+$/u.test(comment.parentId) ? Number(comment.parentId) : undefined);
  const comments = source.flatMap((original): Comment[] => {
    const comment = held.get(original.id);
    if (!comment) return aliases.has(String(original.id)) ? [] : [original];
    const parentId = parentOf(comment);
    const done = initialized ? comment.done : comment.done || (original.done ?? false);
    return [{ ...original, ...(comment.author ? { author: comment.author } : {}), ...(comment.date ? { date: comment.date } : {}),
      ...(comment.body == null ? {} : { content: commentContent(comment.body, original.content) }),
      ...(done === (original.done ?? false) ? {} : { done,
        ...(original.status === undefined ? {} : { status: done ? 'resolved' as const : 'active' as const }) }),
      ...(parentId === undefined ? {} : { parentId }) }];
  });
  const sourceIds = new Set(source.map(comment => comment.id));
  for (const comment of native) {
    const id = numbers.get(comment.id)!;
    if (!sourceIds.has(id)) comments.push({ id, author: comment.author, date: comment.date || undefined,
      content: commentContent(comment.body, []), done: comment.done, parentId: parentOf(comment) });
  }
  return { comments, ids, sessionIds: new Map(native.map(comment => [numbers.get(comment.id)!, comment.id])) };
}

interface CommentInsertionRefusal {
  status: 'refused';
  reason: Extract<DocxEditFailureCode, 'missing-target' | 'ambiguous-target' | 'invalid-step'>;
}

export type ResidentCommentInsertionRange = YrsStoryRange | CommentInsertionRefusal;

const refuse = (reason: CommentInsertionRefusal['reason']): CommentInsertionRefusal => ({ status: 'refused', reason });

export function readCommentInsertionRange(engine: ResidentEngineSession, paraId: string, search?: string): ResidentCommentInsertionRange {
  const matches = engine.geometryReader.storyIds().flatMap(story => engine.geometryReader.paragraphSpans(story)
    .filter(paragraph => paragraph.paraId === paraId).map(paragraph => ({ story, paragraph })));
  if (matches.length !== 1) return refuse(matches.length ? 'ambiguous-target' : 'missing-target');
  if (search === '') return refuse('invalid-step');
  const { story, paragraph } = matches[0]!;
  const at = { story, paraId, offset: 0 };
  if (search === undefined) return { story, start: at, end: { ...at, offset: paragraph.length } };
  const text = engine.selectionText({ story, start: at, end: at }).after;
  const start = text.indexOf(search);
  if (start < 0) return refuse('missing-target');
  if (text.indexOf(search, start + 1) >= 0) return refuse('ambiguous-target');
  const nativeOffset = (offset: number) => {
    let low = 0;
    let high = paragraph.length;
    while (low < high) {
      const middle = Math.floor((low + high) / 2);
      const point = { ...at, offset: middle };
      if (engine.selectionText({ story, start: point, end: point }).before.length < offset) low = middle + 1;
      else high = middle;
    }
    return low;
  };
  return { story, start: { ...at, offset: nativeOffset(start + 1) - 1 }, end: { ...at, offset: nativeOffset(start + search.length) } };
}
