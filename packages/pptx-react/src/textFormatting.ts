import type {
  ParagraphAlignment,
  StorySnapshot,
  TextBoxPrimitive,
  TextStyleSnapshot,
} from '@betteroffice/pptx';
import type { SelectionFormatting } from './components/Toolbar';

export interface EffectiveTextStyle {
  bold: boolean;
  italic: boolean;
  underline: string;
  fontSizePt: number;
  color: string;
  fontFamily: string;
}

interface StyledSpan {
  start: number;
  end: number;
  style: EffectiveTextStyle;
}

export function selectionFormattingFromStory(
  story: StorySnapshot,
  anchor: number,
  focus: number,
  fallback: EffectiveTextStyle
): SelectionFormatting {
  const styles = selectedStyles(story, anchor, focus, fallback);
  const underline = commonValue(styles, 'underline');
  return {
    bold: commonValue(styles, 'bold'),
    italic: commonValue(styles, 'italic'),
    underline: underline === undefined ? undefined : underline !== 'none',
    fontSize: commonValue(styles, 'fontSizePt'),
    textColor: commonValue(styles, 'color'),
    fontFamily: commonValue(styles, 'fontFamily'),
  };
}

export function storyFormattingFromStory(
  story: StorySnapshot,
  fallback: EffectiveTextStyle
): SelectionFormatting {
  return selectionFormattingFromStory(story, 0, story.length, fallback);
}

const RESOLVED_ALIGNMENTS: Record<string, ParagraphAlignment> = {
  left: 'l',
  center: 'ctr',
  right: 'r',
  justify: 'just',
};

const STORED_ALIGNMENTS: Record<string, ParagraphAlignment> = {
  l: 'l',
  ctr: 'ctr',
  r: 'r',
  just: 'just',
  justLow: 'just',
  dist: 'just',
  thaiDist: 'just',
};

/** The alignment shared by every paragraph the selection touches, or
 *  `undefined` when they differ. The laid-out text box supplies the value a
 *  paragraph inherits from its layout or master. */
export function paragraphAlignmentFromSelection(
  story: StorySnapshot,
  textBox: TextBoxPrimitive | undefined,
  anchor: number,
  focus: number
): ParagraphAlignment | undefined {
  const ranges = storyTextRanges(story);
  const limit = ranges[ranges.length - 1]?.end ?? 0;
  const start = Math.min(Math.min(anchor, focus), limit);
  const end = Math.max(anchor, focus);
  const alignments = ranges
    .map((range, index) => ({ range, index }))
    .filter(({ range }) =>
      start === end
        ? start >= range.start && start <= range.end
        : start <= range.end && end > range.start
    )
    .map(({ index }) => paragraphAlignment(story, textBox, index));
  if (alignments.length === 0) return undefined;
  const first = alignments[0];
  return alignments.every((alignment) => alignment === first) ? first : undefined;
}

function paragraphAlignment(
  story: StorySnapshot,
  textBox: TextBoxPrimitive | undefined,
  index: number
): ParagraphAlignment {
  const stored = story.paragraphs[index]?.alignment;
  if (stored) return STORED_ALIGNMENTS[stored] ?? 'l';
  const resolved = textBox?.paragraphs[index]?.align;
  return resolved ? RESOLVED_ALIGNMENTS[resolved] ?? 'l' : 'l';
}

export function storyTextRanges(story: StorySnapshot): Array<{ start: number; end: number }> {
  let start = 0;
  return story.paragraphs.map((paragraph) => {
    const end =
      start + paragraph.runs.reduce((length, run) => length + run.text.length, 0);
    const range = { start, end };
    start = end + 1;
    return range;
  });
}

export function effectiveStyleFromSelection(
  story: StorySnapshot,
  anchor: number,
  focus: number,
  fallback: EffectiveTextStyle
): EffectiveTextStyle {
  const formatting = selectionFormattingFromStory(story, anchor, focus, fallback);
  return {
    bold: formatting.bold ?? fallback.bold,
    italic: formatting.italic ?? fallback.italic,
    underline:
      formatting.underline === undefined
        ? fallback.underline
        : formatting.underline
          ? 'sng'
          : 'none',
    fontSizePt: formatting.fontSize ?? fallback.fontSizePt,
    color: formatting.textColor ?? fallback.color,
    fontFamily: formatting.fontFamily ?? fallback.fontFamily,
  };
}

function selectedStyles(
  story: StorySnapshot,
  anchor: number,
  focus: number,
  fallback: EffectiveTextStyle
): EffectiveTextStyle[] {
  const spans = storySpans(story, fallback);
  if (spans.length === 0) return [fallback];
  const start = Math.min(anchor, focus);
  const end = Math.max(anchor, focus);
  if (start === end) {
    const current = spans.find((span) => start >= span.start && start < span.end);
    if (current) return [current.style];
    const previous = [...spans].reverse().find((span) => span.end <= start);
    return [previous?.style ?? spans[0].style];
  }
  const selected = spans
    .filter((span) => span.end > start && span.start < end)
    .map((span) => span.style);
  if (selected.length > 0) return selected;
  const previous = [...spans].reverse().find((span) => span.end <= start);
  return [previous?.style ?? spans[0].style];
}

function storySpans(story: StorySnapshot, fallback: EffectiveTextStyle): StyledSpan[] {
  const spans: StyledSpan[] = [];
  let position = 0;
  story.paragraphs.forEach((paragraph, paragraphIndex) => {
    for (const run of paragraph.runs) {
      const start = position;
      position += run.text.length;
      if (position > start) {
        spans.push({
          start,
          end: position,
          style: resolveTextStyle(run.style, fallback),
        });
      }
    }
    if (paragraphIndex < story.paragraphs.length - 1) position += 1;
  });
  return spans;
}

function resolveTextStyle(
  style: TextStyleSnapshot,
  fallback: EffectiveTextStyle
): EffectiveTextStyle {
  return {
    bold: style.bold ?? fallback.bold,
    italic: style.italic ?? fallback.italic,
    underline: style.underline ?? fallback.underline,
    fontSizePt: style.fontSizePt ?? fallback.fontSizePt,
    color: style.color ?? fallback.color,
    fontFamily: style.fontFamily ?? fallback.fontFamily,
  };
}

function commonValue<K extends keyof EffectiveTextStyle>(
  styles: EffectiveTextStyle[],
  key: K
): EffectiveTextStyle[K] | undefined {
  const first = styles[0]?.[key];
  return styles.every((style) => style[key] === first) ? first : undefined;
}
