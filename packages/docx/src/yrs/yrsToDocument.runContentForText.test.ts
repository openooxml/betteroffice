import { expect, test } from 'bun:test';
import type { RunContent, TextFormatting } from '../types/document';
import { runContentForText } from './yrsToDocument';

function referenceRunContentForText(text: string, formatting: TextFormatting): RunContent[] {
  const content: RunContent[] = [];
  let plainText = '';
  const flushText = (): void => {
    if (!plainText) return;
    content.push({ type: 'text', text: plainText });
    plainText = '';
  };

  const fonts = formatting.fontFamily
    ? [
        formatting.fontFamily.ascii,
        formatting.fontFamily.hAnsi,
        formatting.fontFamily.eastAsia,
        formatting.fontFamily.cs,
      ].filter((font): font is string => Boolean(font))
    : [];
  const symbolFont = fonts.length > 0 && fonts.every((font) => font === fonts[0]) ? fonts[0] : null;

  for (const char of text) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (char === '\u00ad') {
      flushText();
      content.push({ type: 'softHyphen' });
    } else if (char === '\u2011') {
      flushText();
      content.push({ type: 'noBreakHyphen' });
    } else if (symbolFont && codePoint >= 0xf000 && codePoint <= 0xf8ff) {
      flushText();
      content.push({
        type: 'symbol',
        font: symbolFont,
        char: codePoint.toString(16).toUpperCase().padStart(4, '0'),
      });
    } else {
      plainText += char;
    }
  }
  flushText();
  return content;
}

const FORMATTINGS: TextFormatting[] = [
  {},
  { fontFamily: {} },
  { fontFamily: { ascii: 'Symbol' } },
  { fontFamily: { hAnsi: 'Wingdings', eastAsia: 'Wingdings' } },
  { fontFamily: { ascii: 'Symbol', hAnsi: 'Symbol', eastAsia: 'Symbol', cs: 'Symbol' } },
  { fontFamily: { ascii: 'Symbol', hAnsi: 'Arial' } },
  { fontFamily: { ascii: '', hAnsi: '', cs: '' } },
];

test('runContentForText matches the original loop for Unicode and symbol fonts', () => {
  const texts = [
    '', 'plain text', 'caf\u00e9\u4e2d', '\n\r\t\u0000', '\ud83d\ude00',
    '\ud800', '\udfff', 'a\ud800b\udfff', '\u00ad', '\u2011',
    'before\u00adafter', 'before\u2011after', '\u00ad\u2011\u00ad',
    '\uefff\uf000\uf001\uf8fe\uf8ff\uf900',
    '\uf000\uf8ff', 'a\uf041b\uf042c',
    '\udb80\udc00\udbff\udfff',
    'a\u00ad\ud83d\ude00\uf041\u2011\ud800z',
  ];
  for (const text of texts) {
    for (const formatting of FORMATTINGS) {
      const expected = referenceRunContentForText(text, formatting);
      expect(runContentForText(text, formatting)).toEqual(expected);
      expect(runContentForText(text, () => formatting)).toEqual(expected);
    }
  }
});

test('runContentForText matches the original loop for deterministic mixed text', () => {
  let state = 0x51f00d;
  const random = (limit: number): number => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0;
    return state % limit;
  };
  const chunks = [
    'abc', ' ', '\n', '\u00ad', '\u2011', '\uf000', '\uf8ff', '\uefff', '\uf900',
    '\ud83d\ude00', '\ud800', '\udfff', '\udb80\udc00', '\u4e2d',
  ];
  for (let index = 0; index < 1000; index += 1) {
    let text = '';
    const length = random(32);
    for (let offset = 0; offset < length; offset += 1) text += chunks[random(chunks.length)];
    const formatting = FORMATTINGS[random(FORMATTINGS.length)]!;
    expect(runContentForText(text, () => formatting))
      .toEqual(referenceRunContentForText(text, formatting));
  }
});

test('runContentForText does not resolve formatting on the fast path', () => {
  const formatting = (): TextFormatting => {
    throw new Error('formatting must stay lazy');
  };
  expect(runContentForText('', formatting)).toEqual([]);
  for (const text of ['plain', '\ud83d\ude00', '\ud800\udfff', '\uefff\uf900']) {
    expect(runContentForText(text, formatting)).toEqual([{ type: 'text', text }]);
  }
});

test('runContentForText resolves slow-path formatting once', () => {
  const formatting: TextFormatting = { fontFamily: { ascii: 'Symbol' } };
  let calls = 0;
  const text = 'a\uf041\u00ad\u2011\uf8ff';
  const content = runContentForText(text, () => {
    calls += 1;
    return formatting;
  });
  expect(calls).toBe(1);
  expect(content).toEqual(referenceRunContentForText(text, formatting));
});
