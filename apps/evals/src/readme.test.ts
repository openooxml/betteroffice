import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { loadData } from './build';
import { aggregate } from './stats';

test('headline suite matches the README model and both Evals columns', async () => {
  const readme = await readFile(resolve(import.meta.dir, '../../../README.md'), 'utf8');
  const section = readme.match(/^## Evals[ \t]*\r?$[\s\S]*?(?=^## |<!--|(?![\s\S]))/m)?.[0];
  expect(section).toBeDefined();
  const model = section!.match(/`([^`]+)` via `codex exec`/)?.[1];
  expect(model).toBeDefined();
  const table = section!.match(/^\|.*\|[ \t]*\r?$(?:\r?\n^\|.*\|[ \t]*\r?$)+/m)?.[0];
  expect(table).toBeDefined();
  const [headers, , ...rows] = table!.split(/\r?\n/).map(line => line.trim().slice(1, -1).split('|').map(cell => cell.replace(/\*\*|`/g, '').trim()));
  const suite = (await loadData()).suites[0]!;
  expect(suite.arms.map(a => a.label)).toEqual(headers!.slice(1));
  for (const row of rows) {
    const format = row[0]!.toLowerCase();
    for (const [index, arm] of suite.arms.entries()) {
      const stats = aggregate(suite, { model, arm: arm.id, format });
      expect(`${stats.passes}/${stats.n}`).toBe(row[index + 1]!);
    }
  }
});
