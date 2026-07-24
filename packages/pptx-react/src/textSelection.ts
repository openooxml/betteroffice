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
