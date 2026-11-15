import { describe, expect, test } from 'bun:test';
import type * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { EXPLICIT_OFF_SAVE_CASES } from '../../packages/docx/src/yrs/__fixtures__/explicitOffProperties';

for (const group of new Set(EXPLICIT_OFF_SAVE_CASES.map((entry) => entry.group))) {
  describe(`native: ${group}`, () => {
    for (const { name, run } of EXPLICIT_OFF_SAVE_CASES.filter((entry) => entry.group === group)) {
      test(name, async () => {
        const { path, savedAs, expected } = await run(nativeYrs as unknown as typeof wasmYrs);
        expect({ path }).toEqual({ path: 'replaced' });
        expect({ savedAs }).toEqual({ savedAs: expected });
      });
    }
  });
}
