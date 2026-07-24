import type { Document } from '@betteroffice/docx/types/document';

type ProjectDocument = (baseDocument?: Document | null) => Document | null;

export interface DocumentChangeSinks {
  push(document: Document): void;
  notify?: (document: Document) => void;
}

function commitDocument(document: Document, sinks: DocumentChangeSinks): Document {
  sinks.push(document);
  sinks.notify?.(document);
  return document;
}

export function commitLegacyDocumentChange(
  newDocument: Document,
  projectDocument: ProjectDocument,
  sinks: DocumentChangeSinks
): Document {
  return commitDocument(projectDocument(newDocument) ?? newDocument, sinks);
}

export function commitYrsDocumentChange(
  projectDocument: ProjectDocument,
  sinks: DocumentChangeSinks
): Document | null {
  const document = projectDocument();
  return document ? commitDocument(document, sinks) : null;
}
