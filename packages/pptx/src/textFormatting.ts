import type { StorySnapshot, TextStyle, TextStyleSnapshot } from './types';

export interface EffectiveTextStyle {
  bold: boolean;
  italic: boolean;
  underline: string;
  fontSizePt: number;
  color: string;
  fontFamily: string;
}

export const DEFAULT_EDITOR_TEXT_STYLE: EffectiveTextStyle = {
  bold: false,
  italic: false,
  underline: 'none',
  fontSizePt: 24,
  color: '#111827',
  fontFamily: 'Arial',
};

export interface SelectionTextFormatting {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  fontSize?: number;
  textColor?: string;
  fontFamily?: string;
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
): SelectionTextFormatting {
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
): SelectionTextFormatting {
  return selectionFormattingFromStory(story, 0, story.length, fallback);
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

export function insertionStyleFromSelection(
  story: StorySnapshot, anchor: number, focus: number, storedMarks: TextStyle,
): TextStyle {
  return { ...effectiveStyleFromSelection(story, anchor, focus, DEFAULT_EDITOR_TEXT_STYLE), ...storedMarks };
}

export function insertionEchoRun(
  story: StorySnapshot, anchor: number, focus: number, storedMarks: TextStyle,
): { neighbour: 'before' | 'after'; start: number; end: number; paragraphId: string } | null {
  if (anchor !== focus) return null;
  const inserted = { ...insertionStyleFromSelection(story, anchor, focus, storedMarks), caps: null };
  const identical = (style: TextStyleSnapshot) => {
    const properties = new Set([...Object.keys(style), ...Object.keys(inserted)]);
    return [...properties].every((key) =>
      (style[key as keyof TextStyleSnapshot] ?? null) === (inserted[key as keyof typeof inserted] ?? null));
  };
  let position = 0;
  for (const paragraph of story.paragraphs) {
    let previous: { style: TextStyleSnapshot; start: number; end: number } | undefined;
    for (const run of paragraph.runs) {
      const start = position;
      position += run.text.length;
      if (position === start) continue;
      if (anchor >= start && anchor < position) {
        if (!identical(run.style) || anchor === start && previous && !identical(previous.style)) return null;
        return { neighbour: 'after', start, end: position, paragraphId: paragraph.id };
      }
      previous = { style: run.style, start, end: position };
    }
    if (anchor === position && previous && identical(previous.style)) {
      return { neighbour: 'before', start: previous.start, end: previous.end, paragraphId: paragraph.id };
    }
    position += 1;
  }
  return null;
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
