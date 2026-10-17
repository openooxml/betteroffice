import type { SectionProperties } from '../types/document';
import type { YrsLoc, YrsSession } from './index';

export type SectionStart = NonNullable<SectionProperties['sectionStart']>;

const sectionOf = (attrs: Record<string, unknown>) =>
  (attrs.sectPr ?? attrs._sectionProperties) as SectionProperties | undefined;

/**
 * Splits the paragraph at `at` into a new section boundary: the first half
 * closes with `properties(closing)` and the section's closing paragraph, when
 * one follows, starts with `start`. `closed` is false when the split falls in
 * the final section, whose start lives in the part metadata.
 */
export function splitSection(
  session: Pick<YrsSession, 'paragraphs' | 'splitParagraph' | 'setParagraphAttrs'>,
  at: YrsLoc,
  properties: (closing: SectionProperties | undefined) => SectionProperties,
  start: SectionStart
): { at: YrsLoc; closed: boolean } {
  const paragraphs = session.paragraphs(at.story);
  const index = paragraphs.findIndex((paragraph) => paragraph.paraId === at.paraId);
  if (index < 0) throw new Error('Section break paragraph is unavailable');
  const closing = paragraphs.slice(index).find((paragraph) => sectionOf(paragraph.properties));
  const closingProperties = closing && sectionOf(closing.properties);
  const first = properties(closingProperties);
  const split = session.splitParagraph(at);
  const set = (paraId: string, sectPr: SectionProperties) =>
    session.setParagraphAttrs(
      { story: at.story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
      { other: { sectPr, sectionBreakType: sectPr.sectionStart ?? null } }
    );
  set(split.firstParaId, first);
  if (closing) {
    set(closing.paraId === at.paraId ? split.secondParaId : closing.paraId, {
      ...closingProperties,
      sectionStart: start,
    });
  }
  return { at: { story: at.story, paraId: split.secondParaId, offset: 0 }, closed: Boolean(closing) };
}
