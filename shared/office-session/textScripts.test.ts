import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { textScripts } from './textScripts';

const fixture = JSON.parse(readFileSync(
  fileURLToPath(new URL('../../crates/ooxml-text/tests/fixtures/text-scripts.json', import.meta.url)), 'utf8',
)) as { ranges: Array<[number, number, string[]]>; mixes: Array<{ text: string; scripts: string[] }> };

test('every code point falls in the script bucket the native engines give it', () => {
  expect(fixture.ranges.length).toBeGreaterThan(40);
  const wrong: Array<[number, string[], string[]]> = [];
  for (const [start, end, scripts] of fixture.ranges) {
    for (let point = start; point <= end; point += 1) {
      if (point >= 0xd800 && point <= 0xdfff) continue;
      const found = textScripts(String.fromCodePoint(point));
      if (JSON.stringify(found) !== JSON.stringify(scripts)) wrong.push([point, found, scripts]);
    }
  }
  expect(wrong.slice(0, 5)).toEqual([]);
});

test('mixed text lists its buckets in the native order', () => {
  expect(fixture.mixes.length).toBe(600);
  for (const { text, scripts } of fixture.mixes) expect([text, textScripts(text)]).toEqual([text, scripts]);
});
