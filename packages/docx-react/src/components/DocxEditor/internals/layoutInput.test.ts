import { expect, test } from 'bun:test';
import { sameLayoutInput } from './layoutInput';

test('equal layout input strings match', () => {
  const input = '{"bodyStory":"body","renderEnv":{}}';
  expect(sameLayoutInput(input, input)).toBe(true);
});

test('layout inputs match regardless of nested object key order', () => {
  const held = JSON.stringify({
    bodyStory: 'body',
    renderEnv: {
      compatibilityFlags: { word2013: true, word2010: false },
      tocStyleIds: ['TOC1', 'TOC2'],
      revisionPreview: { r1: 'accepted', r2: 'rejected' },
    },
  });
  const request = JSON.stringify({
    renderEnv: {
      revisionPreview: { r2: 'rejected', r1: 'accepted' },
      compatibilityFlags: { word2010: false, word2013: true },
      tocStyleIds: ['TOC1', 'TOC2'],
    },
    bodyStory: 'body',
  });
  expect(sameLayoutInput(held, request)).toBe(true);
});

test('layout inputs with different values do not match', () => {
  expect(sameLayoutInput(
    '{"renderEnv":{"revisionPreview":{"r1":"accepted"}}}',
    '{"renderEnv":{"revisionPreview":{"r1":"rejected"}}}'
  )).toBe(false);
});

test('layout inputs with a missing versus present key do not match', () => {
  expect(sameLayoutInput('{"renderEnv":{}}', '{"renderEnv":{"revisionPreview":null}}')).toBe(false);
});

test('layout inputs preserve array order', () => {
  expect(sameLayoutInput('{"tocStyleIds":["TOC1","TOC2"]}', '{"tocStyleIds":["TOC2","TOC1"]}')).toBe(false);
});

test('invalid layout input JSON does not match', () => {
  expect(sameLayoutInput('invalid', '{"renderEnv":{}}')).toBe(false);
  expect(sameLayoutInput('{"renderEnv":{}}', 'invalid')).toBe(false);
  expect(sameLayoutInput('invalid', 'also invalid')).toBe(false);
});
