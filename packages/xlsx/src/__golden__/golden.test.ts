/**
 * Golden regression test for the editing seams.
 *
 * Mirror of the display-list golden harness: build each scenario, canonicalize,
 * assert equality with the committed golden. A nonzero diff is an intended
 * equivalence (regenerate) or a regression — stop and investigate.
 *
 * Regenerate (only when a diff is intended and reviewed), from packages/core:
 *   GOLDEN_UPDATE=1 bun test src/__golden__
 */

import { describe, test, expect } from 'bun:test';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';

import { corpus } from './corpus';
import { serialize } from '../display-list/__golden__/serialize';

const GOLDEN_DIR = join(import.meta.dir, 'golden');
const UPDATE = process.env.GOLDEN_UPDATE === '1';

describe('golden editing-seam corpus', () => {
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
            `Generate it with: GOLDEN_UPDATE=1 bun test src/__golden__`
        );
      }

      expect(actual).toBe(expected);
    });
  }
});
