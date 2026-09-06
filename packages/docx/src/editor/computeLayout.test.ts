import { beforeAll, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';

import { preloadEditWasm } from '../wasm/edit';
import { createYrsSession } from '../yrs';
import {
  computeLayout,
  getLayoutKernelInputs,
  type ComputeLayoutInputs,
} from './computeLayout';

const WASM = resolve(import.meta.dir, '../wasm/generated/edit/docx_edit_bg.wasm');
const FONT = resolve(
  import.meta.dir,
  '../../../../crates/ooxml-text/tests/fonts/LiberationSans-Regular.ttf'
);

describe('computeLayout retained kernel inputs', () => {
  beforeAll(() => preloadEditWasm(new Uint8Array(readFileSync(WASM))));

  test('the measured arena is fetched lazily, once, and only on demand', () => {
    const layout = { pages: [] };
    let kernelFetches = 0;
    const session = {
      layoutDocumentWithRegionsRetainedJson: () =>
        JSON.stringify({ layout, notesConverged: true }),
      residentWorkerProbe: () => ({ layoutRevision: 1 }),
      retainedKernelInputsJson: (expectedLayoutRevision: number) => {
        expect(expectedLayoutRevision).toBe(1);
        kernelFetches += 1;
        return JSON.stringify({ measured: [{ block: { kind: 'paragraph' } }], options: { pageGap: 24 } });
      },
    };
    const computation = computeLayout({
      document: null,
      pageGap: 24,
      session: session as never,
      renderEnv: {},
      measurement: { fontChains: {}, defaults: { fontSize: 11, fontFamily: 'Calibri' }, authoritativeShaping: true } as never,
    });
    expect(computation.notesConverged).toBe(true);
    const inputs = getLayoutKernelInputs(computation.layout);
    expect(inputs).toBeDefined();
    expect(kernelFetches).toBe(0);
    expect(inputs!.measured.length).toBe(1);
    expect(inputs!.options).toEqual({ pageGap: 24 });
    expect(kernelFetches).toBe(1);
  });

  test('an older layout cannot fetch a newer retained arena', async () => {
    const session = await createYrsSession({ clientId: 247 });
    try {
      const { paraId } = session.createStory('body', 'first');
      const fontId = session.registerFont(new Uint8Array(readFileSync(FONT)));
      const inputs: ComputeLayoutInputs = {
        document: null,
        pageGap: 24,
        session,
        renderEnv: {},
        measurement: {
          fontChains: { 'calibri|0|0': [fontId] },
          defaults: { fontSize: 11, fontFamily: 'Calibri' },
          compat: { noLeading: false, doNotExpandShiftReturn: false },
          authoritativeShaping: true,
        },
      };

      const first = computeLayout(inputs);
      session.insertText({ story: 'body', paraId, offset: 5 }, ' second');
      const second = computeLayout(inputs);
      const firstKernel = getLayoutKernelInputs(first.layout);
      const secondKernel = getLayoutKernelInputs(second.layout);

      expect(() => firstKernel!.measured).toThrow(
        'retained layout revision mismatch: expected 1, current 2'
      );
      expect(secondKernel!.measured).toHaveLength(1);
    } finally {
      session.destroy();
    }
  });
});
