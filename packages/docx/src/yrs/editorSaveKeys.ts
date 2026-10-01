import type { Comment } from '../types/content';
import type { Document, Endnote, Footnote } from '../types/document';

/** @internal */
export function editorSaveKeys(
  document: Document,
  comments?: readonly Comment[]
): { metadata: string; comments: string; commentIds: number[] } {
  const withoutContent = <T extends { content: unknown }>(part: T) => {
    const { content, ...metadata } = part;
    return metadata;
  };
  const parts = (map: Document['package']['headers']) =>
    [...(map ?? [])].map(([id, part]) => [id, withoutContent(part)]);
  const notes = (entries: (Footnote | Endnote)[] | undefined) =>
    (entries ?? []).map((note) => {
      const { verbatimXml, sourceOrdinal, ...metadata } = withoutContent(note);
      return metadata;
    });
  const { package: pkg } = document;
  const metadata = {
    finalSectionProperties: pkg.document.finalSectionProperties,
    sections: pkg.document.sections?.map(withoutContent),
    headers: parts(pkg.headers),
    footers: parts(pkg.footers),
    footnotes: notes(pkg.footnotes),
    endnotes: notes(pkg.endnotes),
    relationships: [...(pkg.relationships ?? [])],
  };
  const stringify = (value: unknown): string =>
    JSON.stringify(value, (_key, entry) => (entry instanceof Map ? [...entry] : entry));
  const saved = comments ?? pkg.document.comments ?? [];
  return {
    metadata: stringify(metadata),
    comments: stringify(saved),
    commentIds: saved.map((comment) => comment.id),
  };
}
