import { describe, expect, test } from 'bun:test';
import * as yrs from './index';
import { EXPLICIT_OFF_SAVE_CASES } from './__fixtures__/explicitOffProperties';

for (const group of new Set(EXPLICIT_OFF_SAVE_CASES.map((entry) => entry.group))) {
  describe(group, () => {
    for (const { name, run } of EXPLICIT_OFF_SAVE_CASES.filter((entry) => entry.group === group)) {
      test(name, async () => {
        const { path, savedAs, expected } = await run(yrs);
        expect({ path }).toEqual({ path: 'replaced' });
        expect({ savedAs }).toEqual({ savedAs: expected });
      });
    }
  });
}
