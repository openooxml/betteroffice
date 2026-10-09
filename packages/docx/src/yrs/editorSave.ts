/** The editor's save: the session projected with its host's metadata and comments. @internal */

import { createDocx, repackDocx } from '../docx/rezip';
import { injectReplyRangeMarkers, injectTCReplyRangeMarkers } from '../docx/injectReplyRangeMarkers';
import type { Comment } from '../types/content';
import type { Document, Endnote, Footnote, HeaderFooter, Section } from '../types/document';
import { editorSaveKeys } from './editorSaveKeys';
import type { YrsSession } from './index';
import { dirtyProjectionStory } from './dirtyProjectionStories';
import { captureSessionSave, writeSessionSave, type DocxSessionSave } from './saveYrsDocx';
import { sessionSourcePackage } from './sessionInternals';
import { ownProjectedParagraphs, yrsToDocument } from './yrsToDocument';

export {
  DirtyProjectionStories,
  dirtyProjectionStory,
  EditorDirtyStories,
  proposalProjectionStories,
  serialWorkerSaves,
} from './dirtyProjectionStories';

/** What the editor's earlier saves of a session leave for its next save. @internal */
export interface EditorSaveRecord {
  /** A save wrote every part; every later save does too. */
  full: boolean;
  /** The bytes the last save wrote from the source package. */
  saved?: ArrayBuffer;
}

const records = new WeakMap<YrsSession, EditorSaveRecord>();

function sessionRecord(session: YrsSession): EditorSaveRecord {
  let record = records.get(session);
  if (!record) {
    record = { full: false };
    records.set(session, record);
  }
  return record;
}

function mergeHeaderFooterMaps(
  full: Map<string, HeaderFooter> | undefined,
  host: Map<string, HeaderFooter> | undefined
): Map<string, HeaderFooter> | undefined {
  if (host === undefined) return undefined;
  return new Map(
    [...host].map(([relationshipId, metadata]) => {
      const existing = full?.get(relationshipId);
      return [relationshipId, existing ? { ...metadata, content: existing.content } : metadata];
    })
  );
}

function mergeNotes<T extends Footnote | Endnote>(
  full: T[] | undefined,
  host: T[] | undefined
): T[] | undefined {
  if (host === undefined) return undefined;
  return host.map((metadata) => {
    const existing = full?.find((note) => note.id === metadata.id);
    return existing ? { ...existing, ...metadata, content: existing.content } : metadata;
  });
}

function mergeSections(
  full: Section[] | undefined,
  host: Section[] | undefined
): Section[] | undefined {
  if (host === undefined) return undefined;
  return host.map((metadata, index) => {
    const existing =
      full?.find((section) => section.id !== undefined && section.id === metadata.id) ??
      full?.[index];
    return existing ? { ...metadata, content: existing.content } : metadata;
  });
}

/** @internal */
export function mergeDocxHostMetadata(full: Document, host: Document): Document {
  const fullPackage = full.package;
  const hostPackage = host.package;
  return {
    ...full,
    contractVersion: host.contractVersion ?? full.contractVersion,
    originalBuffer: full.originalBuffer ?? host.originalBuffer,
    warnings: host.warnings,
    package: {
      ...fullPackage,
      contractVersion: hostPackage.contractVersion ?? fullPackage.contractVersion,
      styles: hostPackage.styles,
      theme: hostPackage.theme,
      settings: hostPackage.settings,
      fontTable: hostPackage.fontTable,
      relationships: hostPackage.relationships,
      headers: mergeHeaderFooterMaps(fullPackage.headers, hostPackage.headers),
      footers: mergeHeaderFooterMaps(fullPackage.footers, hostPackage.footers),
      footnotes: mergeNotes(fullPackage.footnotes, hostPackage.footnotes),
      endnotes: mergeNotes(fullPackage.endnotes, hostPackage.endnotes),
      document: {
        ...fullPackage.document,
        sections: mergeSections(fullPackage.document.sections, hostPackage.document.sections),
        finalSectionProperties: hostPackage.document.finalSectionProperties,
        comments: hostPackage.document.comments,
      },
    },
  };
}

/** Host metadata read by `mergeDocxHostMetadata`. @internal */
export function hostSaveMetadata(host: Document): Document {
  const pkg = host.package;
  return {
    contractVersion: host.contractVersion,
    warnings: host.warnings,
    package: {
      contractVersion: pkg.contractVersion,
      styles: pkg.styles,
      theme: pkg.theme,
      settings: pkg.settings,
      fontTable: pkg.fontTable,
      relationships: pkg.relationships,
      headers: pkg.headers,
      footers: pkg.footers,
      footnotes: pkg.footnotes,
      endnotes: pkg.endnotes,
      document: {
        content: [],
        sections: pkg.document.sections,
        finalSectionProperties: pkg.document.finalSectionProperties,
        comments: pkg.document.comments,
      },
    },
  };
}

/** Writes the editor's document, through the session save when it has one. */
async function writeEditorDocument(
  document: Document,
  session: YrsSession | null,
  capture: DocxSessionSave | null,
  comments: Comment[],
  injectedMarkers: boolean,
  record?: EditorSaveRecord
): Promise<ArrayBuffer> {
  const original = document.originalBuffer;
  if (!original) return createDocx(document);
  if (!session || !capture) return repackDocx(document);
  const saves = record ?? sessionRecord(session);
  const source = sessionSourcePackage(session);
  const keys = editorSaveKeys(document, comments);
  if (
    !source ||
    keys.metadata !== source.keys.metadata ||
    source.keys.commentIds.some((id) => !new Set(keys.commentIds).has(id)) ||
    saves.full ||
    (original !== saves.saved && !sameBytes(original, source.buffer))
  ) {
    saves.full = true;
    const { bytes } = await writeSessionSave(session, document, capture, original, {}, () => false);
    return bytes.buffer as ArrayBuffer;
  }
  const commentsChanged = keys.comments !== source.keys.comments;
  const bodyPart = capture.identities.paragraphs
    .find(({ session: anchor, source }) => anchor?.story === 'body' && source)
    ?.source?.partUri.slice(1);
  const patches = (part: string): boolean =>
    (!commentsChanged || part !== 'word/comments.xml') &&
    (!injectedMarkers || (bodyPart !== undefined && part !== bodyPart));
  const { bytes } = await writeSessionSave(
    session,
    document,
    capture,
    source.buffer,
    {},
    patches,
    true,
    () => true
  );
  const saved = bytes.buffer as ArrayBuffer;
  saves.saved = saved;
  return saved;
}

function sameBytes(a: ArrayBuffer, b: ArrayBuffer): boolean {
  if (a === b) return true;
  if (a.byteLength !== b.byteLength) return false;
  const left = new Uint8Array(a);
  const right = new Uint8Array(b);
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) return false;
  }
  return true;
}

/**
 * `document` with the comments a save writes, the stories they are anchored
 * in projected again with them: the editor projects its host's comments.
 * With replies, its body paragraphs are its own for their range markers.
 */
function withSavedComments(document: Document, session: YrsSession, comments: Comment[]): Document {
  const base: Document = {
    ...document,
    package: { ...document.package, document: { ...document.package.document, comments } },
  };
  const storyIds = new Set<string>();
  for (const comment of comments) {
    try {
      for (const anchor of session.resolveComment(String(comment.id))) {
        storyIds.add(dirtyProjectionStory(anchor.story));
      }
    } catch {
      // Replies and comments whose anchors are gone hold no range.
    }
  }
  const saved = storyIds.size > 0 ? yrsToDocument(session, base, { storyIds }) : base;
  if (!comments.some((comment) => comment.parentId != null)) return saved;
  const body = saved.package.document;
  return {
    ...saved,
    package: {
      ...saved.package,
      document: { ...body, content: ownProjectedParagraphs(body.content) },
    },
  };
}

/**
 * Saves `projected`, the editor's projection of `session`, with the host's
 * `comments`. `record` holds the earlier saves; the session's own by default.
 * @internal
 */
export async function saveEditorDocument(
  session: YrsSession,
  projected: Document,
  comments: Comment[],
  record?: EditorSaveRecord
): Promise<ArrayBuffer> {
  const capture = projected.originalBuffer ? captureSessionSave(session) : null;
  const document = withSavedComments(projected, session, comments);

  // Inject commentRangeStart/End for reply comments that share the parent's range.
  // Pages/Word require every comment (including replies) to have range markers in document.xml.
  const injectedReplies = injectReplyRangeMarkers(document.package.document.content, comments);
  // Also inject range markers for comments that reply to tracked changes.
  const injectedTCReplies = injectTCReplyRangeMarkers(document.package.document.content, comments);

  return writeEditorDocument(
    document,
    session,
    capture,
    comments,
    injectedReplies || injectedTCReplies,
    record
  );
}
