import { afterEach, expect, spyOn, test } from 'bun:test';
import { resetPptxEngineWarningsForTests, warnPptxEngine } from './engineWarnings';

afterEach(resetPptxEngineWarningsForTests);

test('engine warnings are shared by separate package instances', async () => {
  resetPptxEngineWarningsForTests();
  const duplicate = await import(new URL('./engineWarnings.ts?copy', import.meta.url).href);
  const warn = spyOn(console, 'warn').mockImplementation(() => {});
  try {
    for (const reason of ['legacy-sync-api', 'deprecated', 'no-worker', 'no-offscreen-canvas'] as const) {
      warnPptxEngine(reason);
      duplicate.warnPptxEngine(reason);
    }
    expect(warn).toHaveBeenCalledTimes(4);
  } finally { warn.mockRestore(); }
});
