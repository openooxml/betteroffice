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
  /** How many paragraphs the story has. Reads every span of the story. */
  count(): number;
}

/** @internal */
export function storyParagraphSpans(reader: TableDeleteReader, story: string): StoryParagraphSpans {
  const unit = reader.paragraphUnit;
  if (!unit) {
    const spans = reader.paragraphSpans(story);
    return {
      indexOf: paraId => spans.findIndex(span => span.paraId === paraId),
      at: index => spans[index]!,
      count: () => spans.length,
    };
  }
  const read = new Map<number, YrsParagraphLength>();
  let count: number | undefined;
  return {
    indexOf: paraId => unit.call(reader, story, paraId),
    at: index => {
      let span = read.get(index);
      if (!span) {
        span = paragraphSpan(reader.paragraphSegments(story, [index])[0] ?? []);
        read.set(index, span);
      }
      return span;
    },
    count: () => (count ??= reader.paragraphSpans(story).length),
  };
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
  const paragraphs = storyParagraphSpans(reader, at.story);
  let index = paragraphs.indexOf(at.paraId);
  if (index < 0) return null;
  let offset = at.offset;
  if (direction === 'forward' && offset === paragraphs.at(index).length) {
    index += 1;
    offset = 0;
    if (index === paragraphs.count()) return null;
  }
  let start = 0;
  for (const segment of reader.paragraphSegments(at.story, [index])[0] ?? []) {
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
