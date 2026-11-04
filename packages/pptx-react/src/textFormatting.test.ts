import { describe, expect, it } from 'bun:test';
import type { StorySnapshot, TextBoxPrimitive, TextStyleSnapshot } from '@betteroffice/pptx';
import {
  effectiveStyleFromSelection,
  insertionStyle,
  paragraphAlignmentFromSelection,
  selectionFormattingFromStory,
  storyFormattingFromStory,
  storyTextRanges,
} from './textFormatting';

const fallback = {
  bold: false,
  italic: false,
  underline: 'none',
  fontSizePt: 24,
  color: '#111827',
  fontFamily: 'Arial',
};

describe('pptx text formatting', () => {
  it('reads the caret style from the current run', () => {
    const formatting = selectionFormattingFromStory(story(), 7, 7, fallback);
    expect(formatting).toEqual({
      bold: true,
      italic: false,
      underline: true,
      strikethrough: false,
      fontSize: 28,
      textColor: '#325ee6',
      fontFamily: 'Aptos',
    });
  });

  it('leaves mixed selection properties unset', () => {
    const formatting = selectionFormattingFromStory(story(), 0, 10, fallback);
    expect(formatting.bold).toBeUndefined();
    expect(formatting.fontSize).toBeUndefined();
    expect(formatting.fontFamily).toBe('Aptos');
  });

  it('summarizes every run in a shape text story', () => {
    expect(storyFormattingFromStory(shapeStory(), fallback)).toEqual({
      bold: undefined,
      italic: false,
      underline: undefined,
      strikethrough: false,
      fontSize: undefined,
      textColor: undefined,
      fontFamily: undefined,
    });
  });

  it('reports single and double strike as active and noStrike as inactive', () => {
    const marked = story();
    const runs = marked.paragraphs[0].runs;
    expect(effectiveStyleFromSelection(marked, 7, 7, fallback).strike).toBeUndefined();
    runs[0].style.strike = 'sngStrike';
    runs[1].style.strike = 'dblStrike';
    expect(selectionFormattingFromStory(marked, 0, 10, fallback).strikethrough).toBe(true);
    expect(effectiveStyleFromSelection(marked, 7, 7, fallback).strike).toBe('dblStrike');
    runs[1].style.strike = 'noStrike';
    expect(effectiveStyleFromSelection(marked, 7, 7, fallback).strike).toBeUndefined();
    expect(insertionStyle(effectiveStyleFromSelection(marked, 7, 7, fallback)).strike).toBe('noStrike');
    expect(selectionFormattingFromStory(marked, 7, 7, fallback).strikethrough).toBe(false);
    expect(selectionFormattingFromStory(marked, 0, 10, fallback).strikethrough).toBeUndefined();
  });

  it('maps every shape story paragraph to its text range', () => {
    expect(storyTextRanges(shapeStory())).toEqual([
      { start: 0, end: 10 },
      { start: 11, end: 15 },
    ]);
  });

  it('reads the alignment stored on the paragraph the caret sits in', () => {
    const aligned = alignedStory('ctr', null);
    expect(paragraphAlignmentFromSelection(aligned, undefined, 3, 3)).toBe('ctr');
    expect(paragraphAlignmentFromSelection(aligned, undefined, 12, 12)).toBe('l');
  });

  it('falls back to the alignment the laid out text box resolved', () => {
    const inherited = alignedStory(null, null);
    expect(paragraphAlignmentFromSelection(inherited, textBox('center', 'left'), 3, 3)).toBe('ctr');
  });

  it('marks the last paragraph when the caret sits at the story end', () => {
    const aligned = alignedStory('ctr', 'r');
    expect(paragraphAlignmentFromSelection(aligned, undefined, 15, 15)).toBe('r');
    expect(paragraphAlignmentFromSelection(aligned, undefined, 16, 16)).toBe('r');
  });

  it('normalizes the justified variants', () => {
    expect(paragraphAlignmentFromSelection(alignedStory('dist', null), undefined, 3, 3)).toBe('just');
  });

  it('leaves a selection spanning differently aligned paragraphs unset', () => {
    const mixed = alignedStory('ctr', 'r');
    expect(paragraphAlignmentFromSelection(mixed, undefined, 0, 15)).toBeUndefined();
    expect(paragraphAlignmentFromSelection(mixed, undefined, 0, 10)).toBe('ctr');
  });

  it('uses the fallback style for an empty story', () => {
    const empty = { ...story(), length: 0, paragraphs: [{ id: 'p', alignment: null, level: 0, bulletJson: null, runs: [] }] };
    expect(effectiveStyleFromSelection(empty, 0, 0, fallback)).toEqual(fallback);
  });
});

function story(): StorySnapshot {
  return {
    id: 'story',
    length: 10,
    paragraphs: [
      {
        id: 'paragraph',
        alignment: null,
        level: 0,
        bulletJson: null,
        runs: [
          { text: 'Hello', style: style({ fontFamily: 'Aptos' }) },
          {
            text: 'World',
            style: style({
              bold: true,
              underline: 'sng',
              fontSizePt: 28,
              color: '#325ee6',
              fontFamily: 'Aptos',
            }),
          },
        ],
      },
    ],
  };
}

function shapeStory(): StorySnapshot {
  const first = story();
  return {
    ...first,
    length: 15,
    paragraphs: [
      ...first.paragraphs,
      {
        id: 'second-paragraph',
        alignment: null,
        level: 0,
        bulletJson: null,
        runs: [
          {
            text: 'More',
            style: style({
              color: '#db2777',
              fontFamily: 'Calibri',
            }),
          },
        ],
      },
    ],
  };
}

function alignedStory(first: string | null, second: string | null): StorySnapshot {
  const base = shapeStory();
  return {
    ...base,
    paragraphs: [
      { ...base.paragraphs[0], alignment: first },
      { ...base.paragraphs[1], alignment: second },
    ],
  };
}

function textBox(
  ...aligns: Array<'left' | 'center' | 'right' | 'justify'>
): TextBoxPrimitive {
  return {
    kind: 'textBox',
    objectId: 1,
    x: 0,
    y: 0,
    w: 100,
    h: 100,
    anchor: 'top',
    paragraphs: aligns.map((align) => ({ align, level: 0, runs: [] })),
    lines: [],
  };
}

function style(overrides: Partial<TextStyleSnapshot>): TextStyleSnapshot {
  return {
    bold: false,
    italic: false,
    fontSizePt: 24,
    color: '#111827',
    fontFamily: null,
    underline: null,
    spacingPt: null,
    ...overrides,
  };
}
