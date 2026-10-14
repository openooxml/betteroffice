import { expect, test } from 'bun:test';
import * as pptx from './index';
import * as editor from './editor';
import * as peerless from './peerless';

test('root exports preserve both editor and peerless APIs', () => {
  const exports = { ...editor, ...peerless };
  expect(Object.keys(pptx).sort()).toEqual(Object.keys(exports).sort());
  for (const [name, value] of Object.entries(exports)) {
    expect(pptx[name as keyof typeof pptx]).toBe(value);
  }
});
