import { describe, expect, it } from 'bun:test';
import type { StorySnapshot, TextStyleSnapshot } from '@betteroffice/pptx';
import {
  effectiveStyleFromSelection,
  selectionFormattingFromStory,
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

function style(overrides: Partial<TextStyleSnapshot>): TextStyleSnapshot {
  return {
    bold: false,
    italic: false,
    fontSizePt: 24,
    color: '#111827',
    fontFamily: null,
    underline: null,
    ...overrides,
  };
}
