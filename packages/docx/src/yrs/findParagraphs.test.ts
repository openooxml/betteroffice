import { expect, test } from 'bun:test';
import { findParagraphs } from './findParagraphs';

function reader(stories: Record<string, string[]>) {
  return {
    storyIds: () => Object.keys(stories),
    paragraphs: (story: string) => stories[story].map((text, index) => ({ paraId: `${story}-${index}`, text })),
  };
}

test('matches body stories in order, preserves case and skips repeated occurrences', () => {
  const document = reader({
    body: ['A Needle here', 'needle and NEEDLE', 'absent'],
    'hf:rId1': ['Needle'],
    'body:t0:r0c0': ['Another NEEDLE'],
    'fn:1': ['needle'],
    bodyguard: ['needle'],
  });
  expect(findParagraphs(document, 'needle')).toEqual([
    { paraId: 'body-0', match: 'Needle', before: 'A ', after: ' here' },
    { paraId: 'body:t0:r0c0-0', match: 'NEEDLE', before: 'Another ', after: '' },
  ]);
  expect(findParagraphs(document, 'Needle', { caseSensitive: true })).toEqual([
    { paraId: 'body-0', match: 'Needle', before: 'A ', after: ' here' },
  ]);
});

test('counts overlapping occurrences and returns no matches for an empty query', () => {
  expect(findParagraphs(reader({ body: ['aaa', 'aa'] }), 'aa')).toEqual([
    { paraId: 'body-1', match: 'aa', before: '', after: '' },
  ]);
  expect(findParagraphs(reader({ body: ['anything'] }), '')).toEqual([]);
});

test('defaults to twenty matches and honors custom limits', () => {
  const document = reader({ body: Array.from({ length: 25 }, () => 'match') });
  expect(findParagraphs(document, 'match')).toHaveLength(20);
  expect(findParagraphs(document, 'match', { limit: 2 })).toHaveLength(2);
  expect(findParagraphs(document, 'match', { limit: 0 })).toEqual([]);
});

test('keeps forty characters of context on each side', () => {
  const before = 'a'.repeat(50);
  const after = 'b'.repeat(50);
  expect(findParagraphs(reader({ body: [`${before}Needle${after}`] }), 'needle')).toEqual([
    { paraId: 'body-0', match: 'Needle', before: 'a'.repeat(40), after: 'b'.repeat(40) },
  ]);
});
