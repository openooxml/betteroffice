import { describe, expect, test } from 'bun:test';
import type * as wasmYrs from '../../packages/docx/src/yrs';
import * as nativeYrs from '@betteroffice/docx-native/yrs';
import { EDITED_TEXT, EXPLICIT_OFF_SAVE_CASES } from '../../packages/docx/src/yrs/__fixtures__/explicitOffProperties';

for (const group of new Set(EXPLICIT_OFF_SAVE_CASES.map((entry) => entry.group))) {
  describe(`native: ${group}`, () => {
    for (const { name, open, run } of EXPLICIT_OFF_SAVE_CASES.filter((entry) => entry.group === group)) {
      const body = async () => {
        const { path, text, savedAs, expected } = await run(nativeYrs as unknown as typeof wasmYrs);
        expect({ path }).toEqual({ path: 'replaced' });
        expect({ text }).toEqual({ text: EDITED_TEXT });
        expect({ savedAs }).toEqual({ savedAs: expected });
      };
      if (open) test.todo(`${name}; ${open}`, body);
      else test(name, body);
    }
  });
}
