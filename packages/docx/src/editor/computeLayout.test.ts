import { describe, expect, test } from 'bun:test';
import { computeLayout, getLayoutKernelInputs } from './computeLayout';

describe('computeLayout retained kernel inputs', () => {
  test('the measured arena is fetched lazily, once, and only on demand', () => {
    const layout = { pages: [] };
    let kernelFetches = 0;
    const session = {
      layoutDocumentWithRegionsRetainedJson: () =>
        JSON.stringify({ layout, notesConverged: true }),
      retainedKernelInputsJson: () => {
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
});
