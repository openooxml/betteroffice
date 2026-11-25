import type { YrsLoc, YrsParagraphLength, YrsStorySegment } from './index';
import { tableAnchors, type TablePayload } from './storyPlainText';
import { isUnavailableTablePayload } from './yrsPositionProjection';

interface TableDeleteReader {
  paragraphSpans(story: string): readonly YrsParagraphLength[];
  paragraphSegments(story: string, indices: readonly number[]): YrsStorySegment[][];
  /**
   * Index of the first paragraph of `story` with `paraId`, in `paragraphSpans` order, or -1.
   * When given, a delete inside a paragraph reads only the paragraphs it touches, not every
   * span of the story.
   */
  paragraphUnit?(story: string, paraId: string): number;
}

/** @internal A story's paragraph spans, read one paragraph at a time when the reader can find a paragraph by id. */
export interface StoryParagraphSpans {
  /** Index of the first paragraph with `paraId`, in `paragraphSpans` order, or -1. */
  indexOf(paraId: string): number;
  /** The span of the paragraph at `index`, which must be one of the story's paragraphs. */
  at(index: number): YrsParagraphLength;
  segmentsAt(index: number): readonly YrsStorySegment[];
  /** How many paragraphs the story has. Reads every span of the story. */
  count(): number;
  tableBoundaryTarget(at: YrsLoc, direction: 'backward' | 'forward'): YrsLoc | null;
}

/** @internal */
export function storyParagraphSpans(reader: TableDeleteReader, story: string): StoryParagraphSpans {
  const unit = reader.paragraphUnit;
  const spans = unit ? undefined : reader.paragraphSpans(story);
  const readSpans = new Map<number, YrsParagraphLength>();
  const segments = new Map<number, YrsStorySegment[]>();
  const segmentsAt = (index: number) => {
    let paragraph = segments.get(index);
    if (paragraph === undefined) {
      paragraph = reader.paragraphSegments(story, [index])[0] ?? [];
      segments.set(index, paragraph);
    }
    return paragraph;
  };
  let count: number | undefined;
  const paragraphs: StoryParagraphSpans = {
    indexOf: paraId => unit ? unit.call(reader, story, paraId) : spans!.findIndex(span => span.paraId === paraId),
    at: index => {
      if (spans) return spans[index]!;
      let span = readSpans.get(index);
      if (!span) {
        span = paragraphSpan(segmentsAt(index));
        readSpans.set(index, span);
      }
      return span;
    },
    segmentsAt,
    count: () => spans?.length ?? (count ??= reader.paragraphSpans(story).length),
    tableBoundaryTarget: (at, direction) => tableBoundaryTarget(reader, paragraphs, at, direction),
  };
  return paragraphs;
}

/** A paragraph's span from its segments, counted as `paragraphSpans` counts it. */
function paragraphSpan(segments: readonly YrsStorySegment[]): YrsParagraphLength {
  let length = 0;
  for (const segment of segments) {
    if (segment.kind === 'pilcrow') return { paraId: segment.paraId, length };
    length += segment.kind === 'text' ? segment.text.length : 1;
  }
  throw new Error('Paragraph segments end without a pilcrow');
}

/** @internal */
export function tableBoundaryDeleteTarget(
  reader: TableDeleteReader,
  at: YrsLoc,
  direction: 'backward' | 'forward'
): YrsLoc | null {
  return storyParagraphSpans(reader, at.story).tableBoundaryTarget(at, direction);
}

function tableBoundaryTarget(
  reader: TableDeleteReader,
  paragraphs: StoryParagraphSpans,
  at: YrsLoc,
  direction: 'backward' | 'forward'
): YrsLoc | null {
  let index = paragraphs.indexOf(at.paraId);
  if (index < 0) return null;
  let offset = at.offset;
  if (direction === 'forward' && offset === paragraphs.at(index).length) {
    index += 1;
    offset = 0;
    if (index === paragraphs.count()) return null;
  }
  let start = 0;
  for (const segment of paragraphs.segmentsAt(index)) {
    const width = segment.kind === 'text' ? segment.text.length : 1;
    if (segment.kind === 'embed' && segment.embedKind === 'table' &&
      (direction === 'backward' ? offset === start + width || offset === start : offset === start)) {
      if (direction === 'forward' || isUnavailableTablePayload(segment.payload) || !Array.isArray(segment.payload.rows)) return at;
      const cell = tableAnchors(segment.payload as unknown as TablePayload).anchors.at(-1);
      const paragraph = cell && reader.paragraphSpans(cell.story).at(-1);
      return cell && paragraph
        ? { story: cell.story, paraId: paragraph.paraId, offset: paragraph.length }
        : at;
    }
    start += width;
    if (start > offset) break;
  }
  return null;
}
