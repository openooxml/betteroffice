import { expect, test } from 'bun:test';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { loadData } from './build';
import { aggregate, newestFirst, publishedSuite, tiers } from './stats';

test('newest suite matches the README model and Evals columns by tier', async () => {
  const readme = await readFile(resolve(import.meta.dir, '../../../README.md'), 'utf8');
  const section = readme.match(/^## Evals[ \t]*\r?$[\s\S]*?(?=^## |<!--|(?![\s\S]))/m)?.[0];
  expect(section).toBeDefined();
  const model = section!.match(/`([^`]+)` via `codex exec`/)?.[1];
  expect(model).toBeDefined();
  const tables = [...section!.matchAll(/^\|.*\|[ \t]*\r?$(?:\r?\n^\|.*\|[ \t]*\r?$)+/gm)].map(match => match[0]);
  const suite = publishedSuite(newestFirst((await loadData()).suites)[0]!);
  const groups = tiers(suite);
  expect(tables).toHaveLength(groups.length);
  expect(section).not.toMatch(/complex/i);
  expect(section).not.toContain('Sweep 1');
  for (const [tableIndex, tier] of (groups.length > 1 ? groups : [undefined]).entries()) {
    const [headers, , ...rows] = tables[tableIndex]!.split(/\r?\n/).map(line => line.trim().slice(1, -1).split('|').map(cell => cell.replace(/\*\*|`/g, '').trim()));
    const arms = headers!.slice(1).map(label => suite.arms.find(a => a.label === label));
    expect(arms).toHaveLength(suite.arms.length);
    expect(arms.every(Boolean)).toBe(true);
    expect(new Set(arms.map(arm => arm!.id)).size).toBe(suite.arms.length);
    expect(rows.map(row => row[0]!.toLowerCase()).sort()).toEqual([...new Set(suite.tasks.filter(t => !tier || (t.tier ?? 'basic') === tier).map(t => t.format))].sort());
    for (const row of rows) {
      const format = row[0]!.toLowerCase();
      for (const [index, arm] of arms.entries()) {
        const stats = aggregate(suite, { model, arm: arm!.id, format, tier });
        expect(`${stats.passes}/${stats.n}`).toBe(row[index + 1]!);
      }
    }
  }
});
