import { describe, expect, test } from 'bun:test';
import * as yrs from './index';
import { EDITED_TEXT, EXPLICIT_OFF_SAVE_CASES } from './__fixtures__/explicitOffProperties';

for (const group of new Set(EXPLICIT_OFF_SAVE_CASES.map((entry) => entry.group))) {
  describe(group, () => {
    for (const { name, open, run } of EXPLICIT_OFF_SAVE_CASES.filter((entry) => entry.group === group)) {
      if (open) {
        test.todo(`${name}; ${open}`);
        continue;
      }
      test(name, async () => {
        const { path, text, savedAs, expected } = await run(yrs);
        expect({ path }).toEqual({ path: 'replaced' });
        expect({ text }).toEqual({ text: EDITED_TEXT });
        expect({ savedAs }).toEqual({ savedAs: expected });
      });
    }
  });
}
