import { beforeAll, describe, expect, it } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from './index';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');

describe('YrsSession searchText', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  it('returns every paragraph-local match in deterministic story order', async () => {
    const session = await createYrsSession({ clientId: 76001 });
    try {
      const body = session.createStory('body', 'Alpha alpha ALPHA');
      const header = session.createStory('header:rId1', 'before alpha after');

      expect(session.searchText('alpha')).toEqual([
        { story: 'body', paraId: body.paraId, start: 0, end: 5, text: 'Alpha' },
        { story: 'body', paraId: body.paraId, start: 6, end: 11, text: 'alpha' },
        { story: 'body', paraId: body.paraId, start: 12, end: 17, text: 'ALPHA' },
        { story: 'header:rId1', paraId: header.paraId, start: 7, end: 12, text: 'alpha' },
      ]);
    } finally {
      session.destroy();
    }
  });

  it('supports case-sensitive and limited queries without changing the session', async () => {
    const session = await createYrsSession({ clientId: 76002 });
    try {
      session.createStory('body', 'Alpha alpha ALPHA');
      const before = session.encodeStateVector();

      expect(session.searchText('Alpha', { caseSensitive: true })).toHaveLength(1);
      expect(session.searchText('alpha', { limit: 2 }).map((match) => match.start)).toEqual([0, 6]);
      expect(session.searchText('')).toEqual([]);
      expect(session.searchText('alpha', { limit: 0 })).toEqual([]);
      expect(() => session.searchText('alpha', { limit: -1 })).toThrow(RangeError);
      expect(session.encodeStateVector()).toEqual(before);
    } finally {
      session.destroy();
    }
  });

  it('preserves Unicode simple-folding matches and UTF-16 offsets', async () => {
    const text = readFileSync(resolve(import.meta.dir,
      '../../../../crates/docx-edit/tests/fixtures/unicode-search.txt'), 'utf8').trimEnd();
    const session = await createYrsSession({ clientId: 76004 });
    try {
      session.createStory('body', text);
      const cases: Array<[string, string[]]> = [
        ['i', ['I', 'i']], ['I', ['I', 'i']], ['ı', ['ı']], ['İ', ['İ']],
        ['ΐ', ['ΐ', 'ΐ']], ['ΐ', ['ΐ', 'ΐ']],
        ['ΰ', ['ΰ', 'ΰ']], ['ΰ', ['ΰ', 'ΰ']],
        ['ﬅ', ['ﬅ', 'ﬆ']], ['ﬆ', ['ﬅ', 'ﬆ']],
      ];
      for (const [query, expected] of cases) {
        const matches = session.searchText(query);
        expect(matches.map((match) => match.text)).toEqual(expected);
        for (const match of matches) {
          expect(match.start).toBe(text.indexOf(match.text));
          expect(match.end).toBe(match.start + match.text.length);
        }
        expect(session.searchText(query, { caseSensitive: true }).map((match) => match.text))
          .toEqual([query]);
      }
    } finally {
      session.destroy();
    }
  });

  it('does not match across inline embeds and preserves their offset unit', async () => {
    const session = await createYrsSession({ clientId: 76003 });
    try {
      const body = session.createStory('body', 'foobar');
      session.insertPageBreak({ story: 'body', paraId: body.paraId, offset: 3 });

      expect(session.searchText('foobar')).toEqual([]);
      expect(session.searchText('foo')).toEqual([
        { story: 'body', paraId: body.paraId, start: 0, end: 3, text: 'foo' },
      ]);
      expect(session.searchText('bar')).toEqual([
        { story: 'body', paraId: body.paraId, start: 4, end: 7, text: 'bar' },
      ]);
    } finally {
      session.destroy();
    }
  });
});
