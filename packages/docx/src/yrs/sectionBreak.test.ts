import { expect, test } from 'bun:test';

import { breakRefusedWhileSuggesting, finalSectionStartMetadata } from './sectionBreak';

test('page and section breaks are refused only while suggesting', () => {
  expect(breakRefusedWhileSuggesting('insertPageBreak', true)).toBe(true);
  expect(breakRefusedWhileSuggesting('insertSectionBreak', true)).toBe(true);
  expect(breakRefusedWhileSuggesting('insertPageBreak', false)).toBe(false);
  expect(breakRefusedWhileSuggesting('insertSectionBreak', false)).toBe(false);
  expect(breakRefusedWhileSuggesting('insertText', true)).toBe(false);
});

test('final section start metadata keeps every part of the stored metadata', () => {
  const header = { type: 'header', hdrFtrType: 'default', content: [] };
  const stored = {
    body: { sections: [{ id: 'p1', properties: { sectionStart: 'evenPage' as const }, content: [] }], finalSectionProperties: { sectionStart: 'evenPage' as const } },
    headers: [['rIdHeader', header]],
    relationships: [['rIdHeader', { id: 'rIdHeader', type: 'header', target: 'header.xml' }]],
  } as unknown as Parameters<typeof finalSectionStartMetadata>[0];
  const metadata = finalSectionStartMetadata(stored, undefined, [], 'nextPage')!;
  expect(metadata.headers).toEqual(stored.headers);
  expect(metadata.relationships).toEqual(stored.relationships);
  expect(metadata.body.sections!.at(-1)!.properties.sectionStart).toBe('nextPage');
  expect(metadata.body.finalSectionProperties!.sectionStart).toBe('nextPage');
});

test('final section start metadata falls back to the saved keys with sections keyed by their closing paragraphs', () => {
  const saved = JSON.stringify({
    finalSectionProperties: { marginLeft: 2000 },
    sections: [{ properties: { sectionStart: 'continuous' } }, { properties: { marginLeft: 2000 } }],
    headers: [['rIdHeader', { type: 'header', hdrFtrType: 'default' }]],
    footers: [],
    footnotes: [{ id: 1, type: 'normal' }],
    endnotes: [],
    relationships: [],
  });
  const metadata = finalSectionStartMetadata({}, saved, ['p1'], 'oddPage')!;
  expect(metadata.body.sections!.map(section => section.id)).toEqual(['p1', undefined]);
  expect(metadata.body.sections!.at(-1)!.properties).toEqual({ marginLeft: 2000, sectionStart: 'oddPage' });
  expect(metadata.body.finalSectionProperties).toEqual({ marginLeft: 2000, sectionStart: 'oddPage' });
  expect(metadata.headers![0]![1].content).toEqual([]);
  expect(metadata.footnotes![0]!.content).toEqual([]);
  expect(finalSectionStartMetadata({}, saved, ['p1', 'p2'], 'oddPage')!.body.sections!.map(section => section.id)).toEqual(['p1', undefined]);
});

test('final section start metadata is undefined without stored metadata or saved keys', () => {
  expect(finalSectionStartMetadata({}, undefined, [], 'nextPage')).toBeUndefined();
});
