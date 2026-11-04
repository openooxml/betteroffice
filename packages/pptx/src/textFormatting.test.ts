import { expect, test } from 'bun:test';
import {
  DEFAULT_EDITOR_TEXT_STYLE,
  effectiveStyleFromSelection,
  insertionEchoRun,
  insertionStyle,
  insertionStyleFromSelection,
  selectionFormattingFromStory,
  typingStyle,
} from './textFormatting';
import type { StorySnapshot } from './types';

function story(strike?: string): StorySnapshot {
  return {
    id: 'story', length: 5,
    paragraphs: [{ id: 'paragraph', alignment: null, level: 0, bulletJson: null,
      runs: [{ text: 'alpha', style: { ...DEFAULT_EDITOR_TEXT_STYLE, strike } }] }],
  };
}

test('plain insertion leaves strike undefined', () => {
  expect(DEFAULT_EDITOR_TEXT_STYLE.strike).toBeUndefined();
  expect(effectiveStyleFromSelection(story(), 2, 2, DEFAULT_EDITOR_TEXT_STYLE).strike).toBeUndefined();
  expect(insertionStyleFromSelection(story(), 2, 2, {}).strike).toBeUndefined();
  expect(selectionFormattingFromStory(story(), 2, 2, DEFAULT_EDITOR_TEXT_STYLE).strikethrough).toBe(false);
  expect(insertionEchoRun(story(), 2, 2, {})).not.toBeNull();
});

for (const strike of ['sngStrike', 'dblStrike']) {
  test(`caret insertion inherits ${strike} and preserves an explicit clear`, () => {
    const marked = story(strike);
    expect(insertionStyleFromSelection(marked, 2, 2, {}).strike).toBe(strike);
    expect(selectionFormattingFromStory(marked, 2, 2, DEFAULT_EDITOR_TEXT_STYLE).strikethrough).toBe(true);
    expect(insertionStyleFromSelection(marked, 2, 2, { strike: 'noStrike' }).strike).toBe('noStrike');
    expect(insertionEchoRun(marked, 2, 2, { strike: 'noStrike' })).toBeNull();
  });
}

test('an explicit clear remains available at the caret', () => {
  expect(effectiveStyleFromSelection(story('noStrike'), 2, 2, DEFAULT_EDITOR_TEXT_STYLE).strike).toBe('noStrike');
});


test('typing keeps strike clears as an off mark until insertion', () => {
  const off = typingStyle({ strike: 'noStrike' });
  expect(off.strike).toBeUndefined();
  expect(Object.hasOwn(off, 'strike')).toBe(true);
  expect(insertionStyle(off).strike).toBe('noStrike');
  expect(Object.hasOwn(typingStyle({ bold: true } as { bold: boolean; strike?: string }), 'strike')).toBe(false);
  expect(insertionStyle({ strike: 'dblStrike' }).strike).toBe('dblStrike');
});
