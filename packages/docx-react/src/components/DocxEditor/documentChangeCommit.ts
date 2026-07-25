import type { Document } from '@betteroffice/docx/types/document';

type ProjectDocument = (baseDocument?: Document | null) => Document | null;

export interface DocumentChangeSinks {
  push(document: Document): void;
  notify?: (document: Document) => void;
}

/** Runs the projection later, dropping any projection still pending. */
export type ProjectionScheduler = (project: () => void) => void;

/** Matches the Yrs-origin change coalescing window in PagedEditor. */
export const LEGACY_PROJECTION_DELAY_MS = 100;

function commitDocument(document: Document, sinks: DocumentChangeSinks): Document {
  sinks.push(document);
  sinks.notify?.(document);
  return document;
}

/**
 * Publishes a legacy document synchronously and defers the save projection to
 * the scheduler, so a ruler drag pushes one object per event instead of
 * projecting the whole document per event.
 */
export function commitLegacyDocumentChange(
  newDocument: Document,
  projectDocument: ProjectDocument,
  sinks: DocumentChangeSinks,
  scheduleProjection: ProjectionScheduler
): Document {
  sinks.push(newDocument);
  if (!sinks.notify) return newDocument;
  scheduleProjection(() => {
    commitDocument(projectDocument(newDocument) ?? newDocument, sinks);
  });
  return newDocument;
}

export function commitYrsDocumentChange(
  projectDocument: ProjectDocument,
  sinks: DocumentChangeSinks
): Document | null {
  const document = projectDocument();
  return document ? commitDocument(document, sinks) : null;
}
