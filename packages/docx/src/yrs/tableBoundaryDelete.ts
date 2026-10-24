import type { YrsLoc, YrsParagraphLength, YrsStorySegment } from './index';
import { tableAnchors, type TablePayload } from './storyPlainText';

interface TableDeleteReader {
  paragraphSpans(story: string): readonly YrsParagraphLength[];
  paragraphSegments(story: string, indices: readonly number[]): YrsStorySegment[][];
}

/** @internal */
export function tableBoundaryDeleteTarget(
  reader: TableDeleteReader,
  at: YrsLoc,
  direction: 'backward' | 'forward'
): YrsLoc | null {
  const spans = reader.paragraphSpans(at.story);
  let index = spans.findIndex(span => span.paraId === at.paraId);
  if (index < 0) return null;
  let offset = at.offset;
  if (direction === 'forward' && offset === spans[index]!.length) {
    index += 1;
    offset = 0;
    if (index === spans.length) return null;
  }
  let start = 0;
  for (const segment of reader.paragraphSegments(at.story, [index])[0] ?? []) {
    const width = segment.kind === 'text' ? segment.text.length : 1;
    if (segment.kind === 'embed' && segment.embedKind === 'table' &&
      (direction === 'backward' ? offset === start + width || offset === start : offset === start)) {
      if (direction === 'forward' || !Array.isArray(segment.payload.rows)) return at;
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
