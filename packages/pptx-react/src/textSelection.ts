export interface TextRange {
  start: number;
  end: number;
}

export type TextSelectionGranularity = 'word' | 'paragraph';

const wordSegmenter = new Intl.Segmenter(undefined, { granularity: 'word' });

export function textRangeAt(
  text: string,
  index: number,
  granularity: TextSelectionGranularity
): TextRange {
  return granularity === 'word'
    ? wordRangeAt(text, index)
    : paragraphRangeAt(text, index);
}

export function wordRangeAt(text: string, index: number): TextRange {
  const position = clampedIndex(text, index);
  const segments = [...wordSegmenter.segment(text)];
  const selected =
    segments.find(
      (segment) =>
        position >= segment.index &&
        position < segment.index + segment.segment.length
    ) ?? segments[segments.length - 1];
  if (!selected) return { start: position, end: position };
  return {
    start: selected.index,
    end: selected.index + selected.segment.length,
  };
}

export function paragraphRangeAt(text: string, index: number): TextRange {
  const position = clampedIndex(text, index);
  const start = text.lastIndexOf('\n', Math.max(0, position - 1)) + 1;
  const nextBreak = text.indexOf('\n', position);
  return {
    start,
    end: nextBreak === -1 ? text.length : nextBreak,
  };
}

export function extendTextRange(
  initial: TextRange,
  target: TextRange
): OrientedTextRange {
  if (target.start < initial.start) {
    return { anchor: initial.end, focus: target.start };
  }
  if (target.end > initial.end) {
    return { anchor: initial.start, focus: target.end };
  }
  return { anchor: initial.start, focus: initial.end };
}

interface OrientedTextRange {
  anchor: number;
  focus: number;
}

function clampedIndex(text: string, index: number): number {
  if (!Number.isFinite(index)) return 0;
  return Math.max(0, Math.min(text.length, Math.trunc(index)));
}

/** The laid-out lines a caret can move through, as the display list reports
 *  them. Only the fields caret movement reads are required. */
export interface CaretLine {
  start: number;
  end: number;
  caretStops: ReadonlyArray<{ position: number; x: number }>;
}

export interface CaretLocation {
  position: number;
  lineIndex?: number;
}

export interface CaretGoalKey extends CaretLocation {
  shapeId: string;
}

export function sameCaretGoalKey(left: CaretGoalKey, right: CaretGoalKey): boolean {
  return (
    left.shapeId === right.shapeId &&
    left.position === right.position &&
    left.lineIndex === right.lineIndex
  );
}

/** Returns the caret's visual column. */
export function caretGoalX(
  lines: readonly CaretLine[],
  caret: CaretLocation
): number | undefined {
  const line = lines[caretLineIndex(lines, caret)];
  return line?.caretStops.find((stop) => stop.position === caret.position)?.x;
}

/** Moves vertically while retaining `goalX`. */
export function verticalCaretMove(
  lines: readonly CaretLine[],
  caret: CaretLocation,
  direction: 'up' | 'down',
  goalX?: number
): Required<CaretLocation> {
  if (lines.length === 0) return { position: caret.position, lineIndex: 0 };
  const index = caretLineIndex(lines, caret);
  const targetIndex = index + (direction === 'up' ? -1 : 1);
  const target = lines[targetIndex];
  if (!target) return { position: caret.position, lineIndex: index };
  const x = goalX ?? caretGoalX(lines, caret);
  if (x === undefined) return { position: target.start, lineIndex: targetIndex };
  const nearest = target.caretStops.reduce<{ position: number; x: number } | null>(
    (best, stop) => (best === null || Math.abs(stop.x - x) < Math.abs(best.x - x) ? stop : best),
    null
  );
  return { position: nearest?.position ?? target.start, lineIndex: targetIndex };
}

/** The first or last position of the line the caret sits in. */
export function lineEdge(
  lines: readonly CaretLine[],
  caret: CaretLocation,
  edge: 'start' | 'end'
): Required<CaretLocation> {
  const lineIndex = caretLineIndex(lines, caret);
  const line = lines[lineIndex];
  if (!line) return { position: caret.position, lineIndex };
  return { position: edge === 'start' ? line.start : line.end, lineIndex };
}

/** The next word boundary in `direction`, for a word-wise caret jump. */
export function wordBoundary(text: string, index: number, direction: -1 | 1): number {
  const position = clampedIndex(text, index);
  const starts = [...wordSegmenter.segment(text)]
    .filter((segment) => segment.isWordLike)
    .flatMap((segment) => [segment.index, segment.index + segment.segment.length]);
  const candidates = direction < 0 ? starts.filter((s) => s < position) : starts.filter((s) => s > position);
  if (candidates.length === 0) return direction < 0 ? 0 : text.length;
  return direction < 0 ? Math.max(...candidates) : Math.min(...candidates);
}

export function caretLineIndex(
  lines: readonly CaretLine[],
  caret: CaretLocation
): number {
  const preferred = caret.lineIndex;
  if (
    preferred !== undefined &&
    preferred >= 0 &&
    preferred < lines.length &&
    caret.position >= lines[preferred].start &&
    caret.position <= lines[preferred].end
  ) {
    return preferred;
  }
  let index = -1;
  for (let candidate = 0; candidate < lines.length; candidate += 1) {
    const line = lines[candidate];
    if (caret.position >= line.start && caret.position <= line.end) index = candidate;
  }
  return index === -1 ? Math.max(0, lines.length - 1) : index;
}
