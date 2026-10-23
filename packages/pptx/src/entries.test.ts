import { expect, test } from 'bun:test';
import * as pptx from './index';
import * as editor from './editor';
import * as worker from './worker';
import * as collaboration from './collaboration';

test('root exports preserve both editor and worker APIs', () => {
  const exports = { ...editor, ...worker };
  expect(Object.keys(pptx).sort()).toEqual(Object.keys(exports).sort());
  for (const [name, value] of Object.entries(exports)) {
    expect(pptx[name as keyof typeof pptx]).toBe(value);
  }
});

test('collaboration entry matches the root collaboration exports', () => {
  for (const [name, value] of Object.entries(collaboration)) {
    expect(pptx[name as keyof typeof pptx]).toBe(value);
  }
});
