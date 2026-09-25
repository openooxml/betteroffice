/** Session save that reports where each saved paragraph can be found after reopening. */

import type { RepackOptions } from '../docx/rezip';
import { writeDocumentWithRust } from '../docx/rustSaveFacade';
import type { Comment, Paragraph } from '../types/content';
import type { BlockContent, Document } from '../types/document';
import type { YrsSession } from './index';
import type {
  DocxParagraphIdentitySnapshot,
  DocxPersistedParagraphAnchor,
  DocxSessionParagraphAnchor,
} from './paragraphIdentity';
import { projectedSessionKey, yrsToDocument } from './yrsToDocument';

/** A saved paragraph's session anchor and the persisted anchor that resolves to it once reopened. */
export interface DocxSavedParagraph {
  session: DocxSessionParagraphAnchor;
  persisted: DocxPersistedParagraphAnchor;
}

export interface DocxSavedDocument {
  bytes: Uint8Array;
  /**
   * Every session paragraph the bytes hold with a Word paragraph ID, in
   * document order. Stories that share a part are views of its paragraphs:
   * each view the part holds is listed, and its persisted anchor resolves to
   * the first story's view once reopened.
   */
  paragraphs: DocxSavedParagraph[];
  /**
   * Saved paragraphs whose Word paragraph ID the live session reassigned
   * while the save was written, such as by a duplicate repair after a remote
   * update: their persisted anchor finds them in these bytes, not in the
   * live session.
   */
  conflicts: DocxSavedParagraph[];
}

function storyBlocks(document: Document): BlockContent[][] {
  const pkg = document.package;
  return [
    pkg.document.content,
    ...[...(pkg.headers?.values() ?? [])].map((part) => part.content),
    ...[...(pkg.footers?.values() ?? [])].map((part) => part.content),
    ...(pkg.footnotes ?? []).map((note) => note.content),
    ...(pkg.endnotes ?? []).map((note) => note.content),
  ];
}

const COMMENTS_PART = 'word/comments.xml';

function commentContent(body: unknown): Paragraph[] {
  if (
    Array.isArray(body) &&
    body.length > 0 &&
    body.every((entry) => (entry as { type?: unknown } | null)?.type === 'paragraph')
  ) {
    return body as Paragraph[];
  }
  const text = typeof body === 'string' ? body : '';
  return [
    {
      type: 'paragraph',
      formatting: {},
      content: text ? [{ type: 'run', formatting: {}, content: [{ type: 'text', text }] }] : [],
    },
  ];
}

/**
 * The comments a save writes: the source package's, each with any author,
 * date or body the session set on it, then every comment added to the
 * session, under its own ID when that is a free OOXML comment ID and the
 * next free one otherwise. Replies and resolution are the source package's;
 * the session holds neither.
 */
function savedComments(
  session: YrsSession,
  base: Document
): { comments: Comment[]; sessionIds: Map<number, string>; changed: boolean } {
  const source = base.package.document.comments ?? [];
  const held = new Map(session.listComments().map((comment) => [comment.id, comment]));
  let changed = false;
  const comments = source.map((comment) => {
    const entry = held.get(String(comment.id));
    held.delete(String(comment.id));
    if (!entry || (!entry.author && !entry.date && entry.body == null)) return comment;
    changed = true;
    return {
      ...comment,
      ...(entry.author ? { author: entry.author } : {}),
      ...(entry.date ? { date: entry.date } : {}),
      ...(entry.body == null ? {} : { content: commentContent(entry.body) }),
    };
  });
  const used = new Set(source.map((comment) => comment.id));
  const sessionIds = new Map<number, string>();
  let next = Math.max(0, ...used) + 1;
  for (const entry of held.values()) {
    changed = true;
    let id = /^\d+$/.test(entry.id) ? Number(entry.id) : Number.NaN;
    if (!Number.isSafeInteger(id) || used.has(id)) {
      while (used.has(next)) next += 1;
      id = next;
    }
    used.add(id);
    if (String(id) !== entry.id) sessionIds.set(id, entry.id);
    comments.push({
      id,
      author: entry.author,
      ...(entry.date ? { date: entry.date } : {}),
      content: commentContent(entry.body),
    });
  }
  return { comments, sessionIds, changed };
}

/** The session paragraphs a projection writes with a Word paragraph ID. */
function savedParagraphs(
  document: Document,
  identities: DocxParagraphIdentitySnapshot
): DocxSavedParagraph[] {
  const byKey = new Map(
    identities.paragraphs.flatMap((identity) =>
      identity.session ? [[identity.session.paraId, identity] as const] : []
    )
  );
  const saved: DocxSavedParagraph[] = [];
  const visit = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === 'table') {
        for (const row of block.rows) for (const cell of row.cells) visit(cell.content);
      } else if (block.type === 'blockSdt') {
        visit(block.content);
      } else if (block.type === 'paragraph' && block.paraId) {
        const key = projectedSessionKey(block);
        const identity = key === undefined ? undefined : byKey.get(key);
        if (!identity?.session || !identity.persisted) continue;
        saved.push({
          session: identity.session,
          persisted: { ...identity.persisted, paraId: block.paraId },
        });
      }
    }
  };
  for (const blocks of storyBlocks(document)) visit(blocks);
  return saved;
}

/**
 * Saves a session opened from DOCX bytes and reports each saved paragraph's
 * persisted anchor, read from the same projection the bytes are written
 * from and kept only where the written part holds its Word paragraph ID.
 * Parts whose stories are unchanged since the session opened keep
 * their source bytes, paragraph IDs aside. The comments the session holds
 * are saved with their anchors, those added to it included; replies and
 * resolution are saved as the source package has them. A source paragraph without a
 * Word paragraph ID saves without one unless the host called
 * `persistParagraphIds()` first; the IDs a save writes are recorded as
 * saved, so they keep them against unsaved claims from other replicas, once
 * reconciled with the live session.
 */
export async function saveYrsDocx(
  session: YrsSession,
  options: RepackOptions = {}
): Promise<DocxSavedDocument> {
  const base = session.materializeDocx();
  if (!base?.originalBuffer) {
    throw new Error('saveYrsDocx requires a session opened from DOCX bytes');
  }
  const identities = session.paragraphIdentities();
  const plan = session.paragraphSavePlan();
  const comments = savedComments(session, base);
  const document = yrsToDocument(
    session,
    comments.changed
      ? {
          ...base,
          package: {
            ...base.package,
            document: { ...base.package.document, comments: comments.comments },
          },
        }
      : base,
    { commentIds: comments.sessionIds }
  );
  const paragraphs = savedParagraphs(document, identities);
  const { buffer } = await writeDocumentWithRust(
    document,
    base.originalBuffer,
    options,
    undefined,
    undefined,
    comments.changed
      ? { ...plan, patchedParts: plan.patchedParts.filter(({ part }) => part !== COMMENTS_PART) }
      : plan
  );
  const bytes = new Uint8Array(buffer);
  const held = session.writtenParagraphIds(bytes);
  const isWritten = (partUri: string, paraId: string) =>
    held[partUri]?.includes(paraId.toUpperCase()) ?? false;
  const written = paragraphs.filter(({ persisted }) =>
    isWritten(persisted.story.partUri, persisted.paraId)
  );
  const stale = new Set(
    session
      .recordSavedParagraphIds([
        ...written.map(
          ({ session: anchor, persisted }) => [anchor.paraId, persisted.paraId] as const
        ),
        ...plan.assignments
          .filter(({ part, paraId }) => isWritten(`/${part}`, paraId))
          .map(({ part, ordinal, paraId }) => [`/${part}#${ordinal}`, paraId] as const),
      ])
      .map(([owner, paraId]) => `${owner}\u0000${paraId}`)
  );
  const isStale = ({ session: anchor, persisted }: DocxSavedParagraph) =>
    stale.has(`${anchor.paraId}\u0000${persisted.paraId}`);
  return {
    bytes,
    paragraphs: written.filter((paragraph) => !isStale(paragraph)),
    conflicts: written.filter(isStale),
  };
}
