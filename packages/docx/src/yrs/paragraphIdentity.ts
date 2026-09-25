import type { RustParagraphIds } from '../docx/rustSaveFacade';

/**
 * Paragraph identity DTOs for `YrsSession` identity reads and saves.
 *
 * A session key (`YrsLoc.paraId`) addresses a paragraph in one collaborative
 * document session. A Word paragraph ID (`w14:paraId`) is saved in the file;
 * a paragraph has one once its source carried it, it was authored in a
 * session, or the host called `YrsSession.persistParagraphIds()`.
 */

/** A Word story qualified by the package part it is read from. */
export type DocxSourceStory =
  | { partUri: string; kind: 'body' | 'header' | 'footer' }
  | { partUri: string; kind: 'footnote' | 'endnote' | 'comment'; itemId: string };

/** A paragraph in one collaborative document session, on any of its replicas. */
export interface DocxSessionParagraphAnchor {
  kind: 'session';
  sessionId: string;
  story: string;
  paraId: string;
}

/**
 * A `w:p` occurrence in the exact source package the session was opened
 * from: its zero-based position among the part's paragraphs, nested ones
 * included, as the source bytes hold them rather than as edited since.
 */
export interface DocxSourceParagraphAnchor {
  kind: 'source';
  packageSha256: string;
  /** OPC part name, such as `/word/document.xml`. */
  partUri: string;
  paragraphOrdinal: number;
}

/**
 * A saved Word paragraph ID within its source story. It resolves after the
 * saved file is reopened, but it is scoped to a document the host chose: an
 * eight-digit ID is not a document identifier.
 */
export interface DocxPersistedParagraphAnchor {
  kind: 'persisted';
  story: DocxSourceStory;
  paraId: string;
}

export type DocxParagraphAnchor =
  | DocxSessionParagraphAnchor
  | DocxSourceParagraphAnchor
  | DocxPersistedParagraphAnchor;

/** A paragraph a resolution, assignment or diagnostic names. */
export type DocxParagraphRef = DocxSessionParagraphAnchor | DocxSourceParagraphAnchor;

/**
 * Where a Word paragraph ID comes from: `source` IDs are already in the source
 * package; the others are prepared by the session and appear on save.
 */
export type DocxParagraphIdOrigin = 'source' | 'authored' | 'persisted' | 'repaired';

/** Where a paragraph comes from; `synthetic` ones are editor-only until authored into. */
export type DocxParagraphOrigin = 'source' | 'authored' | 'synthetic';

export interface DocxParagraphIdentity {
  /** `null` for a source paragraph outside the session stories. */
  session: DocxSessionParagraphAnchor | null;
  origin: DocxParagraphOrigin;
  /** The Word paragraph ID the paragraph saves with, or `null` when it saves without one. */
  ooxmlParaId: string | null;
  idOrigin: DocxParagraphIdOrigin | null;
  /** `null` without a Word paragraph ID or a source story. */
  persisted: DocxPersistedParagraphAnchor | null;
  /** `null` unless the paragraph occurs in the retained source package. */
  source: DocxSourceParagraphAnchor | null;
}

/**
 * Every paragraph's identities: session paragraphs with stories sorted and in
 * document order, then source paragraphs outside the stories in package order.
 */
export interface DocxParagraphIdentitySnapshot {
  sessionId: string;
  /** SHA-256 of the retained source package, or `null` without one. */
  packageSha256: string | null;
  paragraphs: DocxParagraphIdentity[];
}

/** A paragraph `YrsSession.persistParagraphIds()` gave a Word paragraph ID. */
export interface DocxParagraphIdAssignment {
  paragraph: DocxParagraphRef;
  /** The duplicate session key the paragraph carried before repair. */
  replacedParaId: string | null;
  /** The duplicate Word paragraph ID it replaces. */
  previousOoxmlParaId: string | null;
  ooxmlParaId: string;
  idOrigin: 'persisted' | 'repaired';
  /** `null` when the paragraph has no source story. */
  persisted: DocxPersistedParagraphAnchor | null;
}

export type DocxParagraphIdDiagnostic =
  | {
      /** Saved claims to one ID conflict; the paragraphs keep it and resolve as ambiguous. */
      kind: 'conflicting-saved-ids';
      ooxmlParaId: string;
      paragraphs: DocxParagraphRef[];
    }
  | {
      /** No retained source package, so source paragraphs outside the stories were not covered. */
      kind: 'no-source-package';
    };

/**
 * Why `YrsSession.persistParagraphIds()` changed nothing: a duplicated
 * comment paragraph ID the comment companion parts reference, where which
 * comment they mean is ambiguous.
 */
export interface DocxParagraphIdRefusal {
  kind: 'ambiguous-comment-reference';
  ooxmlParaId: string;
  commentIds: string[];
}

export type DocxParagraphIdentityReceipt =
  | {
      status: 'applied';
      assignments: DocxParagraphIdAssignment[];
      diagnostics: DocxParagraphIdDiagnostic[];
    }
  | { status: 'refused'; refusal: DocxParagraphIdRefusal };

export type DocxParagraphAnchorResult =
  | { status: 'found'; anchor: DocxParagraphRef }
  | { status: 'missing' }
  | { status: 'ambiguous'; candidates: DocxParagraphRef[] }
  | { status: 'unsupported'; reason: 'foreign-session' | 'foreign-package' | 'no-source-package' };

/** The paragraph IDs a session save applies, as the package writer's `paragraphIds`. @internal */
export type DocxParagraphSavePlan = RustParagraphIds;
