import type { SectionProperties } from '../types/document';
import type { YrsLoc, YrsSession, YrsStorySegment } from './index';
import type { EditSession } from './wasm/index';

export type SectionStart = NonNullable<SectionProperties['sectionStart']>;

type BreakSession = Pick<YrsSession, 'paragraphs' | 'splitParagraph' | 'setParagraphAttrs'>;
type BreakWasm = Pick<EditSession, 'section_margins_json' | 'set_section_margins' | 'story_segment_units'>;

const BLOCK_EMBEDS = new Set(['table', 'blockSdt', 'pageBreak', 'columnBreak']);

const sectionOf = (attrs: Record<string, unknown>) =>
  (attrs.sectPr ?? attrs._sectionProperties) as SectionProperties | undefined;

/** Count of the block embeds (tables, block SDTs, page and column breaks) leading a paragraph. */
export function leadingBlockCount(segments: readonly YrsStorySegment[]): number {
  const index = segments.findIndex((segment) => segment.kind !== 'embed' || !BLOCK_EMBEDS.has(segment.embedKind));
  return index < 0 ? segments.length : index;
}

function paragraphIndex(session: BreakSession, at: YrsLoc): number {
  const index = session.paragraphs(at.story).findIndex((paragraph) => paragraph.paraId === at.paraId);
  if (index < 0) throw new Error('Break paragraph is unavailable');
  return index;
}

/**
 * Where a page break at `at` can land as a block: past the paragraph's
 * leading blocks, splitting the paragraph first when inline content precedes `at`.
 */
export function pageBreakTarget(session: BreakSession, wasm: BreakWasm, at: YrsLoc): YrsLoc {
  const index = paragraphIndex(session, at);
  const segments = JSON.parse(wasm.story_segment_units(at.story, Uint32Array.from([index]))) as YrsStorySegment[][];
  const leading = leadingBlockCount(segments[0] ?? []);
  if (at.offset <= leading) return { ...at, offset: leading };
  return { story: at.story, paraId: session.splitParagraph(at).secondParaId, offset: 0 };
}

/**
 * Splits the body paragraph at `at` into a new section boundary: the first half
 * closes with `properties(closing)` and the section's closing paragraph, when
 * one follows, starts with `start`. `closed` is false when the split falls in
 * the final section, whose start lives in the part metadata.
 */
export function splitSection(
  session: BreakSession,
  wasm: BreakWasm,
  at: YrsLoc,
  properties: (closing: SectionProperties | undefined) => SectionProperties,
  start: SectionStart
): { at: YrsLoc; closed: boolean } {
  if (at.story !== 'body') throw new Error('Section breaks belong to body paragraphs');
  const paragraphs = session.paragraphs(at.story);
  const closing = paragraphs.slice(paragraphIndex(session, at)).find((paragraph) => sectionOf(paragraph.properties));
  const margins = JSON.parse(wasm.section_margins_json()) as Record<string, SectionProperties>;
  const closingMargins = closing ? margins[closing.paraId] : undefined;
  const closingProperties = closing && { ...sectionOf(closing.properties), ...closingMargins };
  const first = properties(closingProperties);
  const split = session.splitParagraph(at);
  const set = (paraId: string, sectPr: SectionProperties) =>
    session.setParagraphAttrs(
      { story: at.story, start: { paraId, offset: 0 }, end: { paraId, offset: 0 } },
      { other: { sectPr, sectionBreakType: sectPr.sectionStart ?? null } }
    );
  set(split.firstParaId, first);
  if (closing) {
    const moved = closing.paraId === at.paraId ? split.secondParaId : closing.paraId;
    set(moved, { ...sectionOf(closing.properties), sectionStart: start });
    if (moved !== closing.paraId && closingMargins) wasm.set_section_margins(moved, JSON.stringify(closingMargins));
  }
  return { at: { story: at.story, paraId: split.secondParaId, offset: 0 }, closed: Boolean(closing) };
}
