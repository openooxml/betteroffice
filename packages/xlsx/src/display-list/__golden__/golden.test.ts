/**
 * Golden regression test.
 *
 * For every corpus scenario, build the snapshot, canonicalize it, and assert it
 * equals the committed golden. A nonzero diff is an intended equivalence
 * (regenerate) or a regression — stop and investigate.
 *
 * Regenerate the goldens (only when a diff is intended and reviewed):
 *   GOLDEN_UPDATE=1 bun test src/display-list/__golden__
 * A normal run never writes files.
 */

import { describe, test, expect } from 'bun:test';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

import { corpus } from './corpus';
import { serialize } from './serialize';

const GOLDEN_DIR = join(import.meta.dir, 'golden');
const UPDATE = process.env.GOLDEN_UPDATE === '1';

// first line where two multi-line strings differ, for a readable hint.
function firstDiff(
  expected: string,
  actual: string
): { line: number; expected: string; actual: string } {
  const e = expected.split('\n');
  const a = actual.split('\n');
  const max = Math.max(e.length, a.length);
  for (let i = 0; i < max; i++) {
    if (e[i] !== a[i]) return { line: i + 1, expected: e[i] ?? '<eof>', actual: a[i] ?? '<eof>' };
  }
  return { line: 0, expected: '', actual: '' };
}

describe('golden display-list corpus', () => {
  for (const scenario of corpus) {
    test(scenario.name, () => {
      const actual = serialize(scenario.build());
      const file = join(GOLDEN_DIR, `${scenario.name}.json`);

      if (UPDATE) {
        mkdirSync(GOLDEN_DIR, { recursive: true });
        writeFileSync(file, actual);
        return;
      }

      let expected: string;
      try {
        expected = readFileSync(file, 'utf8');
      } catch {
        throw new Error(
          `Missing golden for scenario "${scenario.name}". ` +
            `Generate it with: GOLDEN_UPDATE=1 bun test src/display-list/__golden__`
        );
      }

      if (actual !== expected) {
        const d = firstDiff(expected, actual);
        throw new Error(
          `Golden mismatch for scenario "${scenario.name}" (${scenario.pins}) at line ${d.line}:\n` +
            `  golden: ${d.expected}\n` +
            `  actual: ${d.actual}\n` +
            `A nonzero diff is an intended equivalence (regenerate) or a regression — ` +
            `stop and investigate. Regenerate with: ` +
            `GOLDEN_UPDATE=1 bun test src/display-list/__golden__`
        );
      }

      expect(actual).toBe(expected);
    });
  }
});
