import { describe, expect, it } from 'bun:test';
import {
  extendTextRange,
  paragraphRangeAt,
  textRangeAt,
  wordRangeAt,
} from './textSelection';

describe('pptx text selection boundaries', () => {
  it('uses Unicode word boundaries', () => {
    expect(wordRangeAt('Hello, café world', 8)).toEqual({ start: 7, end: 11 });
    expect(textRangeAt('Hello, café world', 14, 'word')).toEqual({
      start: 12,
      end: 17,
    });
  });

  it('selects the whitespace segment under the pointer', () => {
    expect(wordRangeAt('Hello,   world', 7)).toEqual({ start: 6, end: 9 });
  });

  it('selects the paragraph containing the caret position', () => {
    const text = 'First line\nSecond paragraph\nThird';
    expect(paragraphRangeAt(text, 3)).toEqual({ start: 0, end: 10 });
    expect(paragraphRangeAt(text, 10)).toEqual({ start: 0, end: 10 });
    expect(textRangeAt(text, 18, 'paragraph')).toEqual({ start: 11, end: 27 });
    expect(paragraphRangeAt(text, text.length)).toEqual({
      start: 28,
      end: 33,
    });
  });

  it('extends a unit selection without splitting either boundary', () => {
    expect(
      extendTextRange({ start: 6, end: 10 }, { start: 0, end: 5 })
    ).toEqual({ anchor: 10, focus: 0 });
    expect(
      extendTextRange({ start: 6, end: 10 }, { start: 11, end: 15 })
    ).toEqual({ anchor: 6, focus: 15 });
  });
});
